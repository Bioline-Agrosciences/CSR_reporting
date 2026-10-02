"""
Paths shared by every script in the project — a single place to change if
the folder layout ever moves.

The CODE and the DATA live in two different places (02/10/2026):

  - The code (this repo, cloned from GitHub) lives on the local disk, e.g.
    C:/Users/<me>/Documents/CSR/monthly_reporting — never inside a
    OneDrive-synced folder: OneDrive locks files mid-sync, which breaks
    `uv sync` on the .venv ("Accès refusé", os error 5), and syncs thousands
    of .venv files to SharePoint for nothing.

  - The data stays on the shared SharePoint library "CSR referents", synced
    on this computer by OneDrive, under its "General" folder (GENERAL_DIR
    below) — found from the user's home folder the same way bioline_utils'
    sharepoint_dir() does it:
        C:/Users/<me>/Bioline Agrosciences Group/CSR referents - Documents/General
    Override with the CSR_SHAREPOINT_GENERAL_DIR environment variable if the
    library is synced somewhere else on a given computer.

    General/
    ├── Reporting_Automation/monthly_reporting/   <- DATA_DIR (same place as when the code lived
    │   ├── input_data/                              here too, so the Fabric shortcut paths don't
    │   │                                            change): indicator reference + long-format raw
    │   │                                            data + CSR_parameters.xlsx
    │   └── output_data/                          <- consolidated results : OUTPUT of csr_calc_engine.py
    ├── Monthly reporting/                        <- data entry files (+ Archives/, extract_bluekango/)
    └── Working Hours 2026.xlsx

Only the empty macro template (templates/data_entry_template.xlsm) stays in
the repo: it's part of the code, not data.

Folders read from / written directly to the shared SharePoint drive — no
manual copying in either direction, no local staging folder:

  - RAW_DATA_DIR = .../General/Monthly reporting/Archives/
    The 6 original raw workbooks (12-tab-per-BU, never modified by this
    project). Read ONCE by extract_reference_and_data.py, to convert them
    into the new format (input_data/) — the historical backfill. Not part
    of the regular monthly cycle: once that conversion is done, this folder
    and that script are never touched again.

    NOTE — this will need updating once Bioline retires the "Archives"
    subfolder (expected around Sept. 2026): the 6 files will then live
    directly in .../General/Monthly reporting/ instead, one level up. Since
    this only matters for that one-time historical read, there is no rush
    to change it before then.

  - DATA_ENTRY_DIR = .../General/Monthly reporting/
    Where generate_data_entry_file.py writes each BU's monthly file
    (<BU>_data_entry_CSR_<year>.xlsm) and integrate_data_entry.py reads it back
    from, once the referent has filled it in. Every BU's CSR referent
    already has standing access to this folder — nothing needs to be
    emailed or sent around, and there is no separate "data_entry/" folder
    inside this project.

Two more such SharePoint-direct paths, added for integrate_safety_data.py
(Safety indicators are NOT owned by the CSR referent, so they don't come
through DATA_ENTRY_DIR — see that script's docstring):

  - WORKING_HOURS_FILE = .../General/Working Hours 2026.xlsx
    Directly under General/, NOT under Monthly reporting/ — confirmed on
    disk 22/09/2026 (a previous guess had it one level too deep; if this
    line reverts to a MONTHLY_REPORTING_DIR-based guess again, that's wrong,
    re-apply this fix rather than trusting the guess).

  - ACCIDENTS_DIR = .../General/Monthly reporting/extract_bluekango/
    One level deeper than Monthly reporting/ itself — confirmed on disk
    22/09/2026 (a previous guess pointed at Monthly reporting/ directly,
    missing the extract_bluekango/ subfolder).

This file computes these paths once (regardless of the directory `uv run ...`
is launched from), and creates the input_data/output_data folders if they
don't exist yet — only when the SharePoint library is actually synced on
this computer (never on the CI runner, where it isn't).
"""
import os
from pathlib import Path

# This file lives at <repo root>/scripts/config.py.
PROJECT_ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_DIR = PROJECT_ROOT / "templates"

SHAREPOINT_ORG = "Bioline Agrosciences Group"
SHAREPOINT_LIBRARY = "CSR referents - Documents"
GENERAL_DIR = Path(os.environ.get("CSR_SHAREPOINT_GENERAL_DIR")
                   or Path.home() / SHAREPOINT_ORG / SHAREPOINT_LIBRARY / "General")

MONTHLY_REPORTING_DIR = GENERAL_DIR / "Monthly reporting"
RAW_DATA_DIR = MONTHLY_REPORTING_DIR / "Archives"
DATA_ENTRY_DIR = MONTHLY_REPORTING_DIR
DATA_DIR = GENERAL_DIR / "Reporting_Automation" / "monthly_reporting"
INPUT_DIR = DATA_DIR / "input_data"
OUTPUT_DIR = DATA_DIR / "output_data"

if GENERAL_DIR.is_dir():
    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Names of the 6 original raw Excel files, exactly as provided by Bioline.
# Used by extract_reference_and_data.py (required) and by csr_calc_engine.py
# (only for the optional one-off validation).
RAW_FILES = {
    "Viridaxis": "Bioline monthly reporting - Viridaxis.xlsx",
    "BAF": "Bioline monthly reporting - BAF.xlsx",
    "BFR": "Bioline monthly reporting - BFR.xlsx",
    "BIB": "Bioline monthly reporting - BIB.xlsx",
    "BUK": "Bioline monthly reporting - BUK.xlsx",
    "BUS": "Bioline monthly reporting - BUS.xlsx",
}

# --- Safety data sources (used by integrate_safety_data.py) ----------------
# Confirmed on disk 22/09/2026 (both were off by one folder level from the
# original guess — see the docstring above). If you ever need to relocate
# them again, `uv run scripts/locate_safety_files.py` searches your OneDrive
# for both and prints paths to paste in here.
#
# Working hours: one fixed file, updated in place every month (wide format:
# one row per BU, one column per month). Lives directly under General/, a
# level up from the rest of the monthly reporting stuff.
WORKING_HOURS_FILE = GENERAL_DIR / "Working Hours 2026.xlsx"

# BlueKanGo accidents export: a NEW file each time (name carries an export
# timestamp, e.g. "Accidents_du_travail_Bioline_20260921-165507.xlsx"), so we
# point at a FOLDER + pattern and always pick the most recently modified
# match rather than a fixed filename. One level under Monthly reporting/, in
# the extract_bluekango/ subfolder.
ACCIDENTS_DIR = MONTHLY_REPORTING_DIR / "extract_bluekango"
ACCIDENTS_FILE_PATTERN = "Accidents_du_travail_Bioline_*.xlsx"

# --- Constants for the 3 "calculated, not entered" indicators (decision of
# 23/09/2026 — see csr_calc_engine.py's CONTEXTUAL_VALUES) ------------------
# Country each BU reports from, used to look up public holidays (via the
# `holidays` package) when computing Saf.4.1 "Number of working days for the
# month". ISO 3166-1 alpha-2 codes. BIB ("Iberia") -> Spain and Viridaxis ->
# Belgium are our best guess, not confirmed with Aurélie — flag it if wrong.
BU_COUNTRY = {
    "Viridaxis": "BE",
    "BAF": "KE",
    "BFR": "FR",
    "BIB": "ES",
    "BUK": "GB",
    "BUS": "US",
}

# Energy conversion factors (Ene.6.1, Ene.7.1) and CO2 emission factors
# (Ene.12-15, feeding Carb.1-5) are NOT here anymore (02/10/2026): they
# live in input_data/CSR_parameters.xlsx, one value per (BU, Year), shared
# with the Fabric notebook — see scripts/csr_parameters.py.

# Month-over-month variation threshold above which the "Annual summary" tab
# (generated by generate_data_entry_file.py) highlights a cell in red, and
# above which csr_calc_engine.py's anomaly report flags a value as an
# outlier. 0.20 = 20%.
VARIATION_THRESHOLD = 0.20

# Anomaly report emailed automatically at the end of every
# csr_calc_engine.py run (data quality corrections, missing values, large
# variations) — see scripts/consolidation_report.py. Sent through the local
# Outlook desktop app, no password stored anywhere.
SEND_ERROR_REPORT_EMAIL = True  # set to False to stop sending it

REPORT_RECIPIENTS = [
    "athebault@biolineagrosciences.com",
]
