# History of decisions

Chronological log of what was decided while building the CSR reporting
pipeline, and why. The current state is described in the main `README.md`;
this file keeps the reasoning behind it. Each entry points to the commit
where it landed (`git show <commit>` for the details).

## 21/09/2026 — From the 12-tab workbooks to a reference list + long-format data (`2c2d6f9`, `cf5e043`)

- **Starting point**: each BU sent a 12-tab Excel workbook (one tab per
  month) mixing indicators to fill in and indicators computed by formulas,
  with the indicator list duplicated on every tab.
- **One reference list** (`Indicator_reference_CSR.xlsx`): the 6 BUs were
  verified to share exactly the same 39 indicators, in the same order, on
  all 12 months, so the list was extracted once (from BAF, February, the
  most complete tab). Each indicator gets a `Kind`: `calculated` for the 9
  CSO rows that held an Excel formula, `input` otherwise.
- **Ref.1 reclassified `calculated` = SUM(Ref.2..Ref.9)**: it never had a
  formula, but the 3 totals typed by hand matched that sum exactly, and it
  fills in totals left empty despite detailed leaks.
- **Long-format raw data** (`Raw_data_CSR.xlsx`): input values only, one row
  per (BU, Month, ID), all BUs in one file.
- **Cleaning rules, never silent and never guessed**: values read with
  `data_only=True` (referents often typed formulas like `=598089/1000`);
  decimal comma and numeric text converted with a "Data quality note";
  unreadable text (e.g. `28,9kWh`) left empty and flagged for a fix at the
  source.
- **Data entry file per BU**: only the referent's own `input` indicators,
  locked except the value/comment cells, built on an `.xlsm` template whose
  `Workbook_Open` macro recolors the current month and recalculates
  "Tracking" on every open (so the file stays correct without being
  regenerated). Pre-filled with known values; values already typed in the
  current file win over the history, so regenerating never loses an entry.
- **"Annual summary" tab** (asked by the referents): all months side by side
  as Excel formulas (live, not a snapshot), variations > 20% in red
  (`IF(...="","",...)` so an empty month is not read as 0 and does not
  trigger a false -100%).
- **Validation against the original workbooks**: tests comparing the
  migrated values with the originals.

## 22/09/2026 — Multiple years and Safety data (`d290c9e`)

- **"Year" column everywhere, key (BU, Year, Month, ID)**: without it,
  August 2027 would silently overwrite August 2026. Data entry files become
  one per year (`<BU>_data_entry_CSR_<year>.xlsm`), so a referent never
  confuses this year's file with an old one, and a new year starts with
  nothing pre-filled from the previous one. Existing files migrated by
  `migrate_2026_09_add_year.py`.
- **Safety integrated from the systems that own it**, not through the
  referents' files (Saf.* belong to H&S / HR):
  - Saf.4.2 (hours worked) from `Working Hours 2026.xlsx` (directly under
    `General/`, confirmed on disk — an earlier guess was one level too
    deep); a blank cell = month not declared, not 0.
  - Saf.2 / Saf.3 / Saf.5 from the BlueKanGo export (in
    `Monthly reporting/extract_bluekango/`, confirmed on disk), latest file
    picked automatically since each export is a new timestamped file.
  - **Agency workers excluded** from every Safety KPI ("Bioline employees"
    in the definitions) — applied to Saf.1 too, as an assumption.
  - **A month with no accident = 0**, not missing, for every BU/month present
    in Working Hours (used as the "month is active" signal).
  - **Saf.1 (days without accident) not restricted to the year**: the
    counter must not reset on 1 January.
- **BU countries for public holidays**: BIB -> Spain and Viridaxis ->
  Belgium are best guesses, never confirmed.

## 23/09/2026 — Sums only, Env.1 and FTE out (`9df4ed3`, `622a5db`)

- **The calculation engine computes sums only, never ratios.** Results are
  rolled up to group level; a sum rolls up correctly, a ratio does not
  (average of averages). The 7 ratios (Wat.2, Ene.10, Ene.11, Was.3, Was.4,
  Saf.6, Saf.7) stay in the reference list as `calculated` but are computed
  downstream (Fabric / Power BI) from the raw components. Before this, only
  Wat.2/Ene.11/Was.4 were missing (no Env.1); the other 4 were computed
  locally.
- **Env.1 (Sales) excluded entirely**: the pipeline never has real SAP
  figures; Sales is joined in Fabric with the ratios. Removed from the
  reference list and the history by `migrate_2026_09_remove_env1.py`.
- **Saf.4 (FTE) excluded entirely**: no longer tracked anywhere. With
  Saf.4.2 now always real, the old Saf.6/Saf.7 fallback (Env.1 x Saf.4.1 x
  8) was dropped.
- **Ene.6.1, Ene.7.1, Saf.4.1 reclassified `calculated`**: they were typed by
  hand every month but nobody kept them up to date. The two conversion
  factors are constants, Saf.4.1 (working days) is derived from a
  public-holiday calendar (`holidays` package). They get their own mechanism
  (`CONTEXTUAL_VALUES`), separate from `FORMULAS`, because they depend on
  the BU and period, not on other entered values. Migrated by
  `migrate_2026_09_reclassify_conversion_factors.py`.

## 24/09/2026 — Comparison tool (`1baaa76`)

- `compare_archived_vs_consolidated.py`: compares every indicator in the
  original workbooks with the consolidated results, telling apart real
  differences, absences expected by design (exclusions, ratios) and
  unexpected absences. Used to measure the effect of each pipeline change.

## 25/09/2026 — Long-format completion, CO2 emissions (`5ae8c8d`, `bcd79a0`)

- **Completion table in long format** (BU, Month, %), so Power BI can filter
  and group directly.
- **Carb.1-5 and Ene.12-15 added** (46 indicators): CO2 emissions from
  energy (electricity, natural gas, fuel, LPG, total) and their per-BU
  emission factors. Part of the intended catalog but absent from every raw
  workbook, so appended directly to the reference list
  (`migrate_2026_09_add_carbon_indicators.py`). Quantity x factor is
  additive, so they stay in the sums-only engine.

## 02/10/2026 — Move to Fabric (`7eb9b76`, `8d2292b`, `2c7f20d`, `202ffeb`)

- **Consolidation ported to a Fabric notebook** (`nb_consolidate_csr_data`):
  same logic as `integrate_data_entry.py` + `integrate_safety_data.py` +
  `csr_calc_engine.py` + the detection part of `consolidation_report.py`.
  It reads SharePoint live through a OneLake shortcut (no copies), stores
  the raw data in the Delta table `CSR_raw_data` (seeded once from
  `Raw_data_CSR.xlsx`) and writes the results as Delta tables.
- **Anomaly report**: the Outlook email cannot run in Fabric (no Outlook
  desktop), so anomalies go to the `CSR_anomalies` table (append, with
  `Run_date`), for the Data Pipeline or Activator to send the email.
- **BlueKanGo latest export picked by the timestamp in its name** rather
  than the modification date, which is less reliable through a OneLake
  shortcut. The Working Hours file name follows the target year.
- **Factors moved to `CSR_parameters.xlsx`**, keyed by (Parameter, BU,
  Year), out of `config.py` and the notebook: one source for both, editable
  without code, with a Source column for audit, and keyed by year so
  publishing a new factor never recomputes past emissions. A missing year
  falls back to the latest earlier year (never a later one) and is flagged.
  Created by `migrate_2026_10_create_parameters_file.py` (Year = 2026).
- **Code and data separated**: the repo is cloned outside OneDrive (OneDrive
  locked the `.venv` and broke `uv sync`); the data stays on SharePoint, at
  the same paths, so the Fabric shortcut did not change.
- **Double run validated**: 2808 values compared between the notebook and
  the local pipeline, the only differences being entries made after the
  last local run.
- **Switch to production**: the notebook writes the real tables
  (`CSR_indicators_report`, `CSR_completion_report` — renamed from
  `CSR_completion_tracking` to match the table the Dataflow produced),
  replacing the Dataflow's queries. The local consolidation scripts stopped
  being run.
- **`holidays` pinned to 0.105** in Fabric, the same as `uv.lock`: holiday
  dates estimated in advance (e.g. Idd-ul-Fitr in Kenya) differ between
  versions — BAF March 2026 had 21 working days with 0.105 and 22 with
  Fabric's default version.
- **Repo cleanup**: migration scripts and the former local consolidation
  moved to `archive/`; `Raw_data_CSR.xlsx` is no longer updated (history up
  to 02/10/2026, still used to pre-fill the data entry files and as the
  notebook's reseed source).

## 02/10/2026 — Ratios in the dashboard, sales from the semantic model, notebooks as .ipynb

- **Ratios computed in the Power BI dashboard, not in Fabric**: the second
  notebook computes no indicator at all; ratios are computed on the fly in the dashboard, at the level of
  aggregation displayed. This supersedes the earlier plan of computing
  them per BU in that notebook.
- **Sales (Env.1) read from the commercial dashboard's semantic model**
  ("Bioline Agrosciences sales - Copy", measure "Sales | Commercial",
  scenario 1 = actuals) with `sempy`'s `evaluate_measure`, rather than
  summed from the raw sales table or a separate DAX query: it gives the same
  figure as the commercial report. Summing the raw table (2.5 million rows)
  gave about twice the figure (BUK July 2026: 4.04 M€ vs 2.02 M€).
  The measure is taken as is, whatever its currency handling (a DAX query
  filtered on EURO gave a slightly different figure: 2 032 475 € vs
  2 023 347 € for BUK July 2026).
- **SAP entities -> BUs**: both DUDUTECH entities -> BAF, BIOLINE
  AGROSCIENCES MEXICANA -> BUS (confirmed).
- **The production notebook is `nb_consolidate_data_for_CSR_report`**
  (the draft `nb_consolidate_sap_and_csr_data` is abandoned). It writes the
  two tables the dashboard reads, whose schema is kept as is:
  `CSR_gold_reporting` (Entity, BU, Year, Month number, MonthName, ID,
  Value, Date — sales as `Env.1` in k€ and `Sales_€` in €, one row per SAP
  entity) and `CSR_gold_tracking` (completion, same date columns). Date =
  1st of the month, for the relation with d_Calendar.
- **Cleanup of that notebook**: CSR indicators were joined to the sales
  wide, per entity, which duplicated every CSR value of BAF and BUS (2
  entities each) and dropped the months without sales (inner join). Sales
  and CSR indicators are now stacked instead: CSR rows have no entity, and
  every month is kept. Empty values are no longer written. The dashboard
  filters by BU only, never by Entity (confirmed): the Entity column is
  kept for traceability of the sales, not used in visuals.
- **Notebooks versioned as `.ipynb` exported from Fabric**, replacing the
  `.py` transcriptions whose cell layout no longer matched Fabric (and whose
  sales notebook did not match what ran at all). Exported without outputs,
  which contain real data, using `import-fabric-notebook` from
  bioline_utils (generic, not specific to this project); a test fails if a
  notebook still has outputs.

## 02/10/2026 — Daily pipeline in Fabric

- **One pipeline, scheduled daily at 10:30**, after the sales finish
  updating in the commercial model (around 10:00): dataflow
  `df_CSR_dimensions` (renamed from `df_consolidated_CSR_file`, which now
  only holds Dim_Indicator and Dim_BU) → `nb_consolidate_csr_data` →
  `nb_consolidate_data_for_CSR_report` → refresh of `ms_csr_reporting`,
  each step only on success of the previous one. The model is refreshed by
  the pipeline rather than by its own schedule (12:00), so the dashboard is
  up to date as soon as the data is.
- **Dataflows `df_extract_BlueKanGo` and `df_working_hours` removed**: the
  consolidation notebook reads the BlueKanGo export and the Working Hours
  file directly from SharePoint.
- **`holidays` installed through a Fabric Environment**, not `%pip`: `%pip`
  is blocked when a pipeline runs the notebook (the first pipeline run
  failed on it). The Environment adds `holidays==0.105` from PyPI, which
  overrides the 0.48 built into Fabric.

## 05/10/2026 — Sales filters aligned with the commercial report

- **The sales query now applies the commercial report's filters explicitly**:
  scenario `d_Scenary[Desc_Scenary] = "Actual"`, currency
  `d_Currency[Currency] = "EURO"` and forex method
  `z_Aux Calc Method Forex[Method Forex] = "FOREX BFC (Weighted Average)"`,
  grouped by `d_Calendar[Beginning of the month]`. Without the currency and
  forex filters, the measure did not give the right sales figures, BAF in
  particular. This supersedes the earlier choice of taking the measure "as
  is" with only `Cod_Scenary = 1`.

## 05/10/2026 — CO2 from refrigerant leaks, Carb.5 becomes the total

- **Carb.5.2 = CO2e from refrigerant leaks**: kg leaked per fluid (Ref.2-9)
  x its global warming potential, / 1000 for tCO2e. Still additive, so it
  stays in the consolidation notebook.
- **Carb.5.1 = CO2 from energy** (Carb.1-4), i.e. the former Carb.5.
  **Carb.5 is now the total**, Carb.5.1 + Carb.5.2 ("Total CO2 emissions",
  tCO2e). A dashboard visual that used Carb.5 as energy CO2 now shows the
  total. The notebook checks Carb.5 = Carb.5.1 + Carb.5.2 for every BU and
  month (check "Consistency" in CSR_anomalies).
- **GWPs: IPCC AR5, 100-year** (the GHG Protocol reference), as no GWP was
  available from the group: R449A 1282, R32 677, R404A 3943, R410A 1924,
  R407C 1624, R448A 1273, R134A 1300. R600A (isobutane) is not in AR5's
  table: 3, the EU F-Gas regulation value. Stored in CSR_parameters.xlsx
  (GWP_<fluid>, one row per BU, Year 2026), like the energy factors. The
  GWPs are not indicators of the reference list (internal Ref.X.GWP
  values), so they are not written to CSR_indicators_report.
- **Sub-totals named Carb.5.1 / Carb.5.2**, not Carb.6 / Carb.7: Carb.6 is
  already used in the Power BI dashboard (carbon intensity).
- Validated locally on the 2026 data (05/10/2026): every other indicator
  unchanged, Carb.5.1 equal to the former Carb.5, Carb.5 = Carb.5.1 + Carb.5.2 on
  all 72 BU-months.

## 07/10/2026 — Sums of empty components stay empty

- **A calculated sum is empty when all its components are empty**, instead
  of 0: Ene.9 and Carb.5 showed 0 for months not entered yet, because every
  empty input was read as 0 (`z()`). Now `total()` skips empty terms but
  returns empty if none is filled, and `product()` (quantity x factor) is
  empty as soon as one factor is. A partly entered month still gives a
  partial sum (the missing inputs are flagged in CSR_anomalies); a 0 actually
  typed stays 0.
- The Carb.5 consistency check follows the same rule (an empty sub-total
  counts as 0, both empty = empty), and its Detail shows "empty" instead of
  a blank cell for a missing value.

## 07/10/2026 — BlueKanGo export read by column name

- **Columns located by their header name, not their position**: the
  column order depends on who exports the file from BlueKanGo. The header
  row is the first of the first 10 rows that contains every expected name
  (case, extra spaces and apostrophe type ignored); a missing column raises
  an error listing the headers found, rather than reading the wrong column.
  "Oui"/"OUI" values compared case-insensitively for the same reason.
- Validated on the two exports on SharePoint (21/09 and 05/10/2026):
  accident events identical to the position-based version.

## 07/10/2026 — Check the pipeline after each notebook import

- After the merges of PR #11-13, `CSR_indicators_report` still had Ene.9 /
  Ref.1 / Carb.5 at 0 for empty months and no Carb.5.1 / Carb.5.2: the
  pipeline was most likely still running an older notebook item, earlier
  imports having created new items instead of replacing it (a Notebook
  activity points to an item, not a name). Hence the post-import check in
  the README (Setup): reselect the notebook in each activity, delete or
  rename the old item, run and check a value.

## 07/10/2026 — BU countries confirmed, table for the Power BI map

- **BU countries confirmed**: BIB -> Spain, Viridaxis -> Belgium (the
  22/09 guesses), BUS -> United States (site in Camarillo, California;
  BUS's Mexican SAP entity only matters for sales).
- **`CSR_bu_country` table** (BU, CountryCode, Country), written by
  `nb_consolidate_csr_data` from `BU_COUNTRY`, the mapping already used for
  the public holidays of Saf.4.1, so the map and the working days cannot
  diverge. One row per BU, related to `Dim_BU` in Power BI, rather than a
  country column repeated on every indicator row. English country name in
  addition to the ISO code, as the Power BI map geocodes names better.
- **Region and MapLocation columns**: with the country only, the map puts
  BUS in the middle of the US. `BU_SITE` gives a more precise site where
  needed (BUS: Camarillo, California); `MapLocation` = "City, Region,
  Country" for those BUs, the country name otherwise.

## Open points

- BlueKanGo label "Bioline Viridaxis" never seen in an export, unverified.
- Agency-worker accidents excluded from Saf.1 too: assumption to confirm.
- `nb_consolidate_data_for_CSR_report`: cleaned version to run and
  validate in Fabric (check the dashboard visuals for BAF and BUS, whose
  CSR values are no longer doubled).
- `CSR_indicators_report` and `CSR_completion_report` read by the dashboard
  notebook on 02/10/2026 had no Year column, i.e. were still the Dataflow's
  version: the Dataflow queries must be removed so they stop overwriting
  the consolidation notebook's tables.
- Anomaly email still to be set up in the Fabric Data Pipeline / Activator.
- Refrigerant GWPs: AR5 chosen by default; to replace if the group
  publishes its own values (add rows for the new year in
  CSR_parameters.xlsx).
