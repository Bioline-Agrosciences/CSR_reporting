"""
Paths and constants for generate_data_entry_file.py.

The code (this repo) is cloned on the local disk, outside OneDrive. The data
lives on the "CSR referents" SharePoint library, synced on this computer by
OneDrive:

    C:/Users/<me>/Bioline Agrosciences Group/CSR referents - Documents/General
    ├── Reporting_Automation/monthly_reporting/input_data/   <- INPUT_DIR
    │       Indicator_reference_CSR.xlsx, CSR_parameters.xlsx, Raw_data_CSR.xlsx
    └── Monthly reporting/                                     <- DATA_ENTRY_DIR
            <BU>_data_entry_CSR_<year>.xlsm

Set the CSR_SHAREPOINT_GENERAL_DIR environment variable if the library is
synced somewhere else on a given computer.
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
DATA_ENTRY_DIR = MONTHLY_REPORTING_DIR
DATA_DIR = GENERAL_DIR / "Reporting_Automation" / "monthly_reporting"
INPUT_DIR = DATA_DIR / "input_data"

BU_LIST = ["Viridaxis", "BAF", "BFR", "BIB", "BUK", "BUS"]

# Month-over-month variation above which the "Annual summary" tab highlights
# a cell in red. Same threshold as the anomaly check in nb_consolidate_csr_data.
VARIATION_THRESHOLD = 0.20
