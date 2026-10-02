"""
ONE-TIME cleanup (23/09/2026) — run this ONCE, then never again (like
migrate_2026_09_add_year.py, this is not part of the recurring monthly
cycle).

Why: Aurélie decided this pipeline should never track Env.1 ("Sales") at
all — it has no real sales/SAP figures to begin with, that indicator is
joined downstream in Fabric instead (see extract_reference_and_data.py and
csr_calc_engine.py for the full rationale, including why the ratio
indicators that used to divide by it aren't computed here anymore either).
extract_reference_and_data.py and integrate_data_entry.py were updated so
Env.1 is never picked up again going forward, but that doesn't retroactively
clean up:
  - the Env.1 row already sitting in input_data/Indicator_reference_CSR.xlsx
    (written once by extract_reference_and_data.py, before this decision);
  - every historical Env.1 row already sitting in
    input_data/Raw_data_CSR.xlsx (same reason).
This script removes both — explicitly, per Aurélie's instruction not to
carry old Env.1 values forward even as history.

What this script does, in order:
  1. Reads input_data/Indicator_reference_CSR.xlsx and rewrites it without
     the Env.1 row. Backed up first
     (input_data/Indicator_reference_CSR_pre_env1_removal_backup.xlsx) —
     never overwritten if that backup already exists, so running this twice
     by mistake never destroys the pre-cleanup state.
  2. Same for input_data/Raw_data_CSR.xlsx: every row with ID == "Env.1"
     (any BU, any Year, any Month) is dropped, backed up first the same way.
Both steps are no-ops (printed, not silent) if the file is missing or
already has no Env.1 to remove.

Usage
-----
    uv run scripts/migrate_2026_09_remove_env1.py

Dependencies: pandas, openpyxl.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
from extract_reference_and_data import export_reference
from integrate_data_entry import export_consolidated

EXCLUDED_ID = "Env.1"


def clean_reference():
    path = config.INPUT_DIR / "Indicator_reference_CSR.xlsx"
    if not path.exists():
        print(f"{path} not found — nothing to clean.")
        return

    df = pd.read_excel(path, sheet_name="Reference")
    if EXCLUDED_ID not in set(df["ID"]):
        print(f"{path.name} has no {EXCLUDED_ID} row — nothing to do.")
        return

    backup_path = config.INPUT_DIR / "Indicator_reference_CSR_pre_env1_removal_backup.xlsx"
    if backup_path.exists():
        print(f"Backup {backup_path.name} already exists — not overwriting it "
              f"(this cleanup may have partially run before; check manually if unsure).")
    else:
        df.to_excel(backup_path, sheet_name="Reference", index=False)
        print(f"Backed up the pre-cleanup file to {backup_path.name}.")

    df = df[df.ID != EXCLUDED_ID].reset_index(drop=True)
    export_reference(df)
    print(f"{path.name}: {EXCLUDED_ID} row removed ({len(df)} indicators remain).")


def clean_raw_data():
    path = config.INPUT_DIR / "Raw_data_CSR.xlsx"
    if not path.exists():
        print(f"{path} not found — nothing to clean.")
        return

    df = pd.read_excel(path, sheet_name="Raw_data")
    removed_count = int((df.ID == EXCLUDED_ID).sum())
    if removed_count == 0:
        print(f"{path.name} has no {EXCLUDED_ID} rows — nothing to do.")
        return

    backup_path = config.INPUT_DIR / "Raw_data_CSR_pre_env1_removal_backup.xlsx"
    if backup_path.exists():
        print(f"Backup {backup_path.name} already exists — not overwriting it "
              f"(this cleanup may have partially run before; check manually if unsure).")
    else:
        df.to_excel(backup_path, sheet_name="Raw_data", index=False)
        print(f"Backed up the pre-cleanup file to {backup_path.name}.")

    df = df[df.ID != EXCLUDED_ID].reset_index(drop=True)
    export_consolidated(df)
    print(f"{path.name}: {removed_count} {EXCLUDED_ID} row(s) removed.")


def main():
    print("Step 1/2 — Indicator_reference_CSR.xlsx")
    clean_reference()
    print("\nStep 2/2 — Raw_data_CSR.xlsx")
    clean_raw_data()
    print("\nDone. You can now use the pipeline scripts normally.")


if __name__ == "__main__":
    main()
