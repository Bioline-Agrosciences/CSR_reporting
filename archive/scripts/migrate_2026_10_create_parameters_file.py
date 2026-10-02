"""
ONE-TIME setup (02/10/2026) — run this ONCE, then never again (same spirit
as the migrate_2026_09_*.py scripts, this is not part of the recurring
monthly cycle).

Why: the energy conversion factors (Ene.6.1, Ene.7.1) and CO2 emission
factors (Ene.12-15) used to be hardcoded dicts in config.py — one value per
BU, no year, duplicated by hand in the Fabric notebook. They now live in
input_data/CSR_parameters.xlsx, one row per (Parameter, BU, Year) — see
csr_parameters.py for the rationale.

What this script does: creates input_data/CSR_parameters.xlsx from the
values config.py held until now (copied below, tagged Year=2026 — the only
year this pipeline has processed so far, so the results stay strictly
identical). Never overwrites an existing file: once it exists, it's the
source of truth, edited by hand from then on.

Usage
-----
    uv run scripts/migrate_2026_10_create_parameters_file.py

Dependencies: pandas, openpyxl.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
from csr_parameters import PARAMETER_COLS, PARAMETERS_PATH, PARAMETERS_SHEET, validate_parameters

YEAR = 2026
BUS = list(config.RAW_FILES)

# Values exactly as they were in config.py before this migration.
ELECTRICITY = {
    "Viridaxis": 0.0001325071644737640,
    "BAF": 0.0000793441816171164,
    "BFR": 0.0000350000000000000,
    "BIB": 0.0001868802584036050,
    "BUK": 0.0002096964205652360,
    "BUS": 0.0003595537086916850,
}
FRAMEWORK = "Group CSR indicator framework, provided 25/09/2026"
PARAMETERS = [
    # (Parameter, {BU: value}, Unit, Source)
    ("LPG_CONVERSION_FACTOR", {bu: 13.8 for bu in BUS}, "kWh/kg", "Physical constant (previously config.py)"),
    ("FUEL_CONVERSION_FACTOR", {bu: 10.0 for bu in BUS}, "kWh/L", "Physical constant (previously config.py)"),
    ("ELECTRICITY_EMISSION_FACTOR", ELECTRICITY, "tCO2/kWh", FRAMEWORK + " (country grid mix)"),
    ("NATURAL_GAS_EMISSION_FACTOR", {bu: 0.0002049620000000000 for bu in BUS}, "tCO2/kWh", FRAMEWORK),
    ("FUEL_EMISSION_FACTOR", {bu: 0.0026187600000000000 for bu in BUS}, "tCO2/L", FRAMEWORK),
    ("LPG_EMISSION_FACTOR", {bu: 0.0024154420279113500 for bu in BUS}, "tCO2/kg", FRAMEWORK),
]


def build_parameters() -> pd.DataFrame:
    rows = [{"Parameter": name, "BU": bu, "Year": YEAR, "Value": value, "Unit": unit, "Source": source}
            for name, by_bu, unit, source in PARAMETERS for bu, value in by_bu.items()]
    df = pd.DataFrame(rows, columns=PARAMETER_COLS)
    validate_parameters(df)
    return df


def main():
    if PARAMETERS_PATH.exists():
        print(f"{PARAMETERS_PATH.name} already exists — not overwriting it (it's the source of truth now; "
              f"edit it by hand).")
        return
    df = build_parameters()
    with pd.ExcelWriter(PARAMETERS_PATH) as writer:
        df.to_excel(writer, sheet_name=PARAMETERS_SHEET, index=False)
        ws = writer.sheets[PARAMETERS_SHEET]
        ws.freeze_panes = "A2"
        for col, width in zip("ABCDEF", (30, 11, 7, 24, 10, 60)):
            ws.column_dimensions[col].width = width
        for cell in ws["D"][1:]:
            cell.number_format = "0.0000000000000000"  # show every digit of the emission factors
    print(f"{PARAMETERS_PATH} created: {len(df)} row(s) "
          f"({df.Parameter.nunique()} parameters x {df.BU.nunique()} BUs, year {YEAR}).")


if __name__ == "__main__":
    main()
