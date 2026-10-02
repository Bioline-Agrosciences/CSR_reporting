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
  notebook (`nb_consolidate_sap_and_csr_data`) computes no indicator at
  all; ratios are computed on the fly in the dashboard, at the level of
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
- **`CSR_gold_reporting` keeps the schema the dashboard already reads**:
  Entity, BU, Year, Month (number), MonthName, ID, Value, Date. Env.1 stays
  one row per SAP entity (Entity filled, the dashboard sums by BU); CSR
  indicators have no entity.
- **Notebooks versioned as `.ipynb` exported from Fabric**, replacing the
  `.py` transcriptions whose cell layout no longer matched Fabric (and whose
  sales notebook did not match what ran at all). Exported without outputs,
  which contain real data, using `import-fabric-notebook` from
  bioline_utils (generic, not specific to this project); a test fails if a
  notebook still has outputs.

## Open points

- BU countries BIB -> ES and Viridaxis -> BE (public holidays for Saf.4.1)
  never confirmed.
- BlueKanGo label "Bioline Viridaxis" never seen in an export, unverified.
- Agency-worker accidents excluded from Saf.1 too: assumption to confirm.
- `nb_consolidate_sap_and_csr_data`: cleaned version to run and validate
  in Fabric.
- Anomaly email still to be set up in the Fabric Data Pipeline / Activator.
