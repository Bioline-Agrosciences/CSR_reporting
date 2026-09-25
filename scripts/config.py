"""
Paths shared by every script in the project — a single place to change if
the folder layout ever moves.

Expected project structure (see README.md at the root):

    monthly_reporting/                  <- project root (holds pyproject.toml, .venv, ...)
    ├── input_data/                     <- indicator reference + long-format raw data (1 single
    │                                      consolidated file, all BUs) : OUTPUT of
    │                                      extract_reference_and_data.py, INPUT of
    │                                      generate_data_entry_file.py, integrate_data_entry.py,
    │                                      integrate_safety_data.py and csr_calc_engine.py
    ├── output_data/                    <- consolidated results : OUTPUT of csr_calc_engine.py
    └── scripts/                        <- the code (this file + the scripts) — no src/, no
        ├── config.py                      __init__.py: these aren't modules of a package
        ├── extract_reference_and_data.py  meant to be imported from outside, just scripts
        ├── generate_data_entry_file.py     that get run directly.
        ├── integrate_data_entry.py
        ├── integrate_safety_data.py
        ├── csr_calc_engine.py
        └── consolidation_report.py

Two folders on purpose are NOT inside this project — they are read from /
written directly to the shared SharePoint drive instead, because this whole
project already lives inside the same synced library
(.../CSR referents - Documents/General/...), so those folders are just as
reachable as anything inside the project itself. No manual copying in either
direction, no local staging folder:

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

This file computes these paths once, based on this file's own location
(regardless of the directory `uv run ...` is launched from), and creates the
input_data/output_data folders if they don't exist yet.
"""
from pathlib import Path

# This file lives at <root>/scripts/config.py:
#   parents[0] = scripts
#   parents[1] = <project root> (.../General/Reporting_Automation/monthly_reporting)
#   parents[2] = .../General/Reporting_Automation
#   parents[3] = .../General
PROJECT_ROOT = Path(__file__).resolve().parents[1]

MONTHLY_REPORTING_DIR = PROJECT_ROOT.parents[1] / "Monthly reporting"
RAW_DATA_DIR = MONTHLY_REPORTING_DIR / "Archives"
DATA_ENTRY_DIR = MONTHLY_REPORTING_DIR
INPUT_DIR = PROJECT_ROOT / "input_data"
OUTPUT_DIR = PROJECT_ROOT / "output_data"

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
WORKING_HOURS_FILE = PROJECT_ROOT.parents[1] / "Working Hours 2026.xlsx"

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

# Energy conversion factors — physical constants, not something anyone
# measures or enters monthly. Kept per-BU (even though every BU uses the same
# value today) so a single BU can be corrected later without touching the
# others. Change here, nowhere else — csr_calc_engine.py picks it up on its
# next run automatically.
LPG_CONVERSION_FACTOR = {bu: 13.8 for bu in RAW_FILES}   # kWh per kg of LPG
FUEL_CONVERSION_FACTOR = {bu: 10.0 for bu in RAW_FILES}  # kWh per liter of fuel

# CO2 emission factors (25/09/2026) — used by csr_calc_engine.py's Carb.1-5
# (CO2 emissions from energy consumption). Kept per-BU, same reasoning as
# the conversion factors above: electricity genuinely varies by country
# (grid mix), the other 3 happen to be identical across BUs today but are
# still a dict per BU so one can be corrected without touching the others.
ELECTRICITY_EMISSION_FACTOR = {  # tCO2/kWh — varies with each country's grid mix
    "Viridaxis": 0.0001325071644737640,
    "BAF": 0.0000793441816171164,
    "BFR": 0.0000350000000000000,
    "BIB": 0.0001868802584036050,
    "BUK": 0.0002096964205652360,
    "BUS": 0.0003595537086916850,
}
NATURAL_GAS_EMISSION_FACTOR = {bu: 0.0002049620000000000 for bu in RAW_FILES}  # tCO2/kWh
FUEL_EMISSION_FACTOR = {bu: 0.0026187600000000000 for bu in RAW_FILES}         # tCO2/L
LPG_EMISSION_FACTOR = {bu: 0.0024154420279113500 for bu in RAW_FILES}          # tCO2/kg

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
