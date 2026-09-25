# Monthly CSR Reporting — Bioline Agrosciences

Automates the monthly collection and consolidation of CSR data (energy,
water, waste, safety, workforce...) across Bioline's 6 business units (BUs):
Viridaxis, BAF, BFR, BIB, BUK, BUS.

## In plain language

Every month, each BU's CSR referent needs to report a set of indicators
(electricity used, water consumed, accidents, refrigerant leaks, etc.).
Historically this meant sending around a heavy 12-tab Excel workbook that
mixed indicators to fill in with indicators calculated by formulas — easy to
break, hard to keep in sync across 6 BUs, and hard to tell who still owed
data for a given month.

This project replaces that with:

- **One shared list of indicators** (kept in one file, not duplicated 6
  times) that says, for each indicator, whether a referent fills it in by
  hand or whether it's calculated automatically, and who is responsible for it.
- **One clean, small file per BU per month** for the referent to fill in —
  only the indicators that are actually theirs to enter, nothing they could
  accidentally break.
- **One calculation engine** that recomputes every additive calculated
  indicator (totals) from the entered data — so a formula never needs to be
  copied into a spreadsheet cell by hand again.

Not every indicator is owned by a CSR referent, though: Safety (accidents,
hours worked) is owned by the H&S manager / HR, and is integrated
separately, straight from the systems that already own it (BlueKanGo, the
Working Hours file) — see below.

**This local pipeline computes sums only, never ratios, and never touches
Sales (decision of 23/09/2026):**
- Ratio/intensity indicators (Wat.2, Ene.10, Ene.11, Was.3, Was.4, Saf.6,
  Saf.7) are NOT computed here. A sum rolls up correctly to a group total
  (all 6 BUs' total energy = the sum of each BU's total energy), a ratio
  does not (averaging or summing 6 BUs' "% renewable energy" produces a
  number that looks plausible but is mathematically meaningless — the
  classic "average of averages" mistake). These indicators still exist in
  the reference list, still nobody types them in by hand, they're just
  computed downstream instead (Power BI/Fabric), from the same raw
  components, at whatever level of aggregation is actually meaningful.
- Env.1 ("Sales") is excluded entirely — not read, not stored, not carried
  over from history. This pipeline (CSR referents + BlueKanGo + Working
  Hours) never has real sales/SAP figures to begin with; that figure is
  joined downstream, in Fabric, alongside the ratios above.
- Saf.4 ("FTE" / headcount) is likewise excluded entirely — no longer
  tracked or used anywhere in this pipeline.
- Ene.6.1, Ene.7.1 (LPG/Fuel conversion factors) and Saf.4.1 (working days
  in the month) used to be hand-typed every month, but nothing was actually
  keeping them up to date. They're now computed fresh on every run instead
  (`csr_calc_engine.py`'s `CONTEXTUAL_VALUES`): the two conversion factors
  come straight from `config.py` (`LPG_CONVERSION_FACTOR`,
  `FUEL_CONVERSION_FACTOR` — one value per BU, so a single BU can be
  corrected later without touching the others, even though every BU uses
  the same value today), and Saf.4.1 from a public-holiday-aware calendar
  count (the `holidays` package) for the country each BU reports from
  (`config.BU_COUNTRY`).

**Carb.1-5 — CO2 emissions from energy consumption (added 25/09/2026):**
these 9 indicators (Ene.12-15, the per-BU CO2 emission factors, and
Carb.1-5, the emissions themselves) were part of the intended catalog from
the start but got left out of the initial extraction — unlike every other
indicator here, they're not read from the original raw workbooks at all
(confirmed absent from all 6 files, every one of the 12 monthly tabs), so
they're appended directly in `extract_reference_and_data.py`'s
`CARBON_INDICATORS` instead. Still additive, so still fine to compute here:
a quantity (energy consumed) times a factor sums correctly to a group
total, unlike a ratio.
- `Carb.1 = Ene.1 * Ene.12` (electricity), `Carb.2 = Ene.5 * Ene.13`
  (natural gas), `Carb.3 = Ene.7 * Ene.14` (fuel), `Carb.4 = Ene.6 * Ene.15`
  (LPG), `Carb.5 = Carb.1 + Carb.2 + Carb.3 + Carb.4` (total).
- Ene.12-15 are per-BU CO2 emission factors, from `config.py`
  (`ELECTRICITY_EMISSION_FACTOR`, `NATURAL_GAS_EMISSION_FACTOR`,
  `FUEL_EMISSION_FACTOR`, `LPG_EMISSION_FACTOR`) — electricity genuinely
  varies by country (grid mix), the other 3 happen to be identical across
  BUs today but are still a dict per BU, same reasoning as the conversion
  factors above.

## The pipeline

Folders used by the pipeline are deliberately **not inside this
project** — they live directly on the shared SharePoint drive instead (this
whole project already sits inside the same synced library), so nothing
ever needs to be copied in or sent out by hand:

- `config.RAW_DATA_DIR` = `.../General/Monthly reporting/Archives/`
- `config.DATA_ENTRY_DIR` = `.../General/Monthly reporting/`
- `config.WORKING_HOURS_FILE` = `.../General/Working Hours 2026.xlsx` (the
  year is currently hardcoded in the filename in `config.py` — update it by
  hand when Bioline starts a new year's file)
- `config.ACCIDENTS_DIR` = `.../General/Monthly reporting/extract_bluekango/`
  (the folder the BlueKanGo accidents export lands in)

```
                      ONE-TIME HISTORICAL MIGRATION
        RAW_DATA_DIR (shared drive)  →  input_data/ (this project)
   .../Monthly reporting/Archives/*.xlsx ────────►  input_data/
    (old 12-tab-per-BU workbooks)        extract_reference_and_data.py
                                              Indicator_reference_CSR.xlsx
                                              Raw_data_CSR.xlsx
                                                                  │
                      ┌───────────────────────────────────────────┘
                      │            RECURRING MONTHLY CYCLE
                      ▼
           generate_data_entry_file.py
                      │
                      ▼
     DATA_ENTRY_DIR/<BU>_data_entry_CSR_<year>.xlsm  ── written directly into the
                                                   shared SharePoint folder
                                                   (General/Monthly reporting)
                                                   each CSR referent already
                                                   has standing access to
                      │
        (the referent fills in the current month's orange cells, in place)
                      │
                      ▼
           integrate_data_entry.py          WORKING_HOURS_FILE +
                      │                      ACCIDENTS_DIR (BlueKanGo
                      │                      export) — shared drive
                      │                                │
                      │                      integrate_safety_data.py
                      │                                │
                      ▼                                ▼
        input_data/Raw_data_CSR.xlsx  ◄─────────────────┘  (updated)
                      │
                      ▼
             csr_calc_engine.py
                      │
                      ▼
     output_data/Consolidated_results_CSR.xlsx  (final result, all BUs/months)
```

The migration on the left (`extract_reference_and_data.py`) runs **once**,
to convert the old-format workbooks into the new format. After that, it is
never run again — the monthly cycle on the right only ever touches the new
format.

Both `Raw_data_CSR.xlsx` and `Consolidated_results_CSR.xlsx` carry a "Year"
column (added 22/09/2026) — every row is keyed by (BU, Year, Month,
Indicator), not just (BU, Month), so the pipeline can keep running year
after year without a new year's August silently overwriting the previous
one's. This is also why the data entry file is one file PER YEAR (see step 1
below) rather than a single file reused forever.

## Project layout

| Folder / file | What it holds |
|---|---|
| `input_data/` | The indicator reference list + the consolidated raw data, all BUs. The "database" the pipeline runs on. Never committed to git, except the empty macro template. |
| `output_data/` | The final consolidated results + the anomaly report attachment. Never committed to git. |
| `scripts/` | All the code. Not a Python package on purpose — just scripts you run directly with `uv run`. |
| `tests/` | Automated tests (pytest) for the logic in `scripts/`. |
| `.github/workflows/` | CI: runs the test suite on every push/PR. |

The original raw workbooks and the data entry files are **not** in this
project — see `scripts/config.py` for `RAW_DATA_DIR` and `DATA_ENTRY_DIR`,
both pointing directly at the shared SharePoint drive. Same for
`WORKING_HOURS_FILE` and `ACCIDENTS_DIR`.

## Getting started

Requirements: [uv](https://docs.astral.sh/uv/) (Python package/version
manager). uv takes care of installing the right Python version and
dependencies.

```
uv sync
```

Then run any script with `uv run scripts/<name>.py`, from the project root.

## Running the monthly cycle

1. **Generate the data entry files** (regenerate any time — it re-reads
   the indicator list and the latest known data, so it's always safe to
   re-run):
   ```
   uv run scripts/generate_data_entry_file.py          # all 6 BUs
   uv run scripts/generate_data_entry_file.py BAF BFR   # just a couple of BUs
   ```
   This writes `<BU>_data_entry_CSR_<year>.xlsm` (one file PER YEAR — a new
   one starts automatically the first time this is run in a new year,
   nothing pre-filled from the year before) directly into
   `config.DATA_ENTRY_DIR` (General/Monthly reporting on the shared
   SharePoint drive) — referents already have access to it, nothing needs
   to be emailed or sent around.

2. **The referent fills in their file, in place on SharePoint.** Only the
   "Value of the month" and "Comment" cells for the current month (colored
   orange) are editable. Opening the file may show an Excel security banner
   ("Enable Content") — this needs to be clicked once, it's what lets the
   file automatically highlight the right month and keep the completion
   tracker up to date on every future open, without the file needing to be
   regenerated.

3. **Integrate what's been filled in**, once the referents have entered
   their data:
   ```
   uv run scripts/integrate_data_entry.py          # all BUs whose file is found
   uv run scripts/integrate_data_entry.py BAF       # just one BU
   ```

4. **Integrate Safety data** — independent of step 3 (different sources:
   BlueKanGo + the Working Hours file, not the referent's data entry file),
   order between the two doesn't matter, but both need to have run before
   step 5:
   ```
   uv run scripts/integrate_safety_data.py
   ```
   Reads the Working Hours file (`config.WORKING_HOURS_FILE`) and the
   latest BlueKanGo accidents export (`config.ACCIDENTS_DIR`), and updates
   the same `input_data/Raw_data_CSR.xlsx` with Saf.1 (days without an
   accident), Saf.2/Saf.3 (non-lost-time / lost-time injuries), Saf.4.2
   (hours worked) and Saf.5 (days lost) — these Safety indicators are owned
   by the H&S manager/HR, not the CSR referent, so they never go through
   the data entry file from step 1-3.

5. **Recompute the consolidated results:**
   ```
   uv run scripts/csr_calc_engine.py
   ```
   Writes `output_data/Consolidated_results_CSR.xlsx` (two tabs: "Results",
   every indicator/BU/month, and "Completion", one row per BU x month with
   its % completion for the current year, scoped to the CSR-referent-owned
   indicators — the same tracking each BU's own "Tracking" tab shows
   individually, consolidated here across all 6 BUs in one place; long/tidy
   format, so BU and Month are both plain columns a PivotTable or Power BI
   can filter and group by directly), and
   emails the anomaly report (see below).

## The indicator reference list

`input_data/Indicator_reference_CSR.xlsx` is the single source of truth for
what gets tracked. Each row is one indicator, with:

- **Kind**: `input` (a referent types it in) or `calculated` (the engine
  computes it — never shown in the file the referent fills in). `calculated`
  covers three different mechanisms under the hood: a sum of other entered
  values (`csr_calc_engine.py`'s `FORMULAS`), a value derived from context
  alone — a constant or a calendar, never anyone's entry
  (`CONTEXTUAL_VALUES`), or computed downstream instead, not by this
  pipeline at all (the ratios, see above).
- **Responsible**: who actually owns that number — usually "CSR referent",
  but some indicators belong to Finance, HR, or the H&S manager instead;
  those are entered elsewhere, not through this pipeline's data entry file.
  Safety (Saf.*, owned by H&S manager/HR) is the one case this pipeline
  handles too, via `integrate_safety_data.py` (see step 4 above) rather
  than a data entry file. Sales (Env.1) and FTE (Saf.4) don't appear in
  this reference list at all anymore (see above) — Env.1 is handled
  entirely downstream, in Fabric/Power BI, and Saf.4 isn't tracked
  anywhere in this pipeline any longer.

39 indicators originally; 37 remain after Env.1 and Saf.4 were excluded
(23/09/2026), plus 9 more appended on top (25/09/2026 — Ene.12-15 and
Carb.1-5, see above): 46 in total.

Add, edit, or remove a row here and every script picks it up automatically
on its next run — nothing else to change.

## Data confidentiality

Real Bioline data never gets committed to git — the raw workbooks and data
entry files live entirely outside the project (on the shared SharePoint
drive, see above), and `input_data/`/`output_data/` are excluded (see
`.gitignore`), except for the empty macro template. Only code, tests, and
configuration are tracked.

## Anomaly report email

Every `csr_calc_engine.py` run also emails a summary of anything worth a
second look: data quality corrections, missing values, and large
month-over-month variations. A per-BU count table goes in the email itself;
the full row-by-row detail is attached as
`output_data/Anomaly_report_CSR_<date>.xlsx`. Sent through the local Outlook
desktop app — no password stored anywhere.

**Outlook must already be open** when you run `csr_calc_engine.py` — the
script only attaches to an already-running Outlook, it never launches one
itself (launching it from a script can hang indefinitely with no error,
waiting on a UI it can't show). If Outlook isn't open, the report is still
printed to the console and the detail workbook still gets written to
`output_data/`, it just doesn't get emailed.

Configure who receives it, or turn it off entirely, in `scripts/config.py` (`REPORT_RECIPIENTS`,
`SEND_ERROR_REPORT_EMAIL`).

## Tests

```
uv run pytest
```

Every push and pull request also runs the test suite automatically via
GitHub Actions (see `.github/workflows/tests.yml`).

## One-off utility scripts (not part of the monthly cycle)

- `scripts/locate_safety_files.py` — searches your OneDrive for the Working
  Hours file and the BlueKanGo export, to help fill in `WORKING_HOURS_FILE`
  and `ACCIDENTS_DIR` in `config.py`. Run once, or again if either file
  ever moves.
- `scripts/migrate_2026_09_remove_env1.py` — one-time cleanup for the
  23/09/2026 decision above: removes the Env.1 row already sitting in
  `Indicator_reference_CSR.xlsx` and every historical Env.1 row already in
  `Raw_data_CSR.xlsx` (both backed up first). Already run once on the real
  files; only needed again on a fresh setup that still has an older
  `Indicator_reference_CSR.xlsx`/`Raw_data_CSR.xlsx` predating this decision.
- `scripts/migrate_2026_09_reclassify_conversion_factors.py` — same idea,
  for the Saf.4 (FTE) exclusion and the Ene.6.1/Ene.7.1/Saf.4.1
  reclassification above. Already run once on the real files.
- `scripts/migrate_2026_09_add_carbon_indicators.py` — one-time cleanup for
  the 25/09/2026 decision above: regenerates
  `Indicator_reference_CSR.xlsx` so it includes the 9 carbon/emission-factor
  indicators (Ene.12-15, Carb.1-5) that were omitted from the initial
  extraction (backed up first). No change needed to `Raw_data_CSR.xlsx` —
  none of these 9 are ever hand-entered.
- `scripts/nb_consolidate_sap_and_csr_data.py` — not run from here at all
  (a Fabric/Spark notebook, kept in this repo for reference and version
  history): recomputes the 7 ratio indicators and joins the real Sales/SAP
  figure downstream, from the raw components this pipeline still produces.
- `scripts/compare_archived_vs_consolidated.py` — diagnostic tool, run by
  hand whenever you want to see exactly how much a pipeline change moved
  the numbers: compares every indicator as it originally appeared in the
  raw per-BU workbooks (`RAW_DATA_DIR`) against `output_data/
  Consolidated_results_CSR.xlsx`, for `HISTORICAL_YEAR` only. Tells apart
  genuine differences, absences that are expected by design (Env.1/Saf.4
  exclusions, ratio indicators no longer computed here — derived live from
  the current `FORMULAS`/`CONTEXTUAL_VALUES`/`EXCLUDED_IDS`, so it can't
  silently drift out of sync), and unexpected absences worth investigating.
  Prints a per-ID summary and a sample of rows to the console, and always
  writes the full detail to a timestamped `output_data/
  Comparison_archived_vs_consolidated_<date>.xlsx` ("Summary" and "Details"
  sheets).
  ```
  uv run scripts/compare_archived_vs_consolidated.py          # every indicator
  uv run scripts/compare_archived_vs_consolidated.py Saf       # only IDs starting with "Saf"
  ```
