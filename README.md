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
- **One calculation engine** that recomputes every calculated indicator
  (totals, ratios) from the entered data — so a formula never needs to be
  copied into a spreadsheet cell by hand again.

## The pipeline

Two folders used by the pipeline are deliberately **not inside this
project** — they live directly on the shared SharePoint drive instead (this
whole project already sits inside the same synced library), so nothing
ever needs to be copied in or sent out by hand:

- `config.RAW_DATA_DIR` = `.../General/Monthly reporting/Archives/`
- `config.DATA_ENTRY_DIR` = `.../General/Monthly reporting/`

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
     DATA_ENTRY_DIR/<BU>_data_entry_CSR.xlsm  ── written directly into the
                                                   shared SharePoint folder
                                                   (General/Monthly reporting)
                                                   each CSR referent already
                                                   has standing access to
                      │
        (the referent fills in the current month's orange cells, in place)
                      │
                      ▼
           integrate_data_entry.py
                      │
                      ▼
        input_data/Raw_data_CSR.xlsx  (updated)
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
both pointing directly at the shared SharePoint drive.

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
   This writes `<BU>_data_entry_CSR.xlsm` directly into
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

4. **Recompute the consolidated results:**
   ```
   uv run scripts/csr_calc_engine.py
   ```
   Writes `output_data/Consolidated_results_CSR.xlsx`.

## The indicator reference list

`input_data/Indicator_reference_CSR.xlsx` is the single source of truth for
what gets tracked. Each row is one indicator, with:

- **Kind**: `input` (a referent types it in) or `calculated` (the engine
  computes it — never shown in the file the referent fills in).
- **Responsible**: who actually owns that number — usually "CSR referent",
  but some indicators belong to Finance, HR, or the H&S manager instead;
  those are entered elsewhere, not through this pipeline's data entry file.

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
