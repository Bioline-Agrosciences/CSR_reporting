"""
ONE-TIME migration (22/09/2026) — run this ONCE, before using the updated
pipeline scripts for real, then never again (like extract_reference_and_data.py,
this is not part of the recurring monthly cycle).

Why: the pipeline now tags every row with a "Year" (BU, Year, Month, ID
instead of just BU, Month, ID — see README.md), and the data entry file is
now one file PER YEAR (<BU>_data_entry_CSR_<year>.xlsm instead of
<BU>_data_entry_CSR.xlsm). Without this migration:
  - input_data/Raw_data_CSR.xlsx has no "Year" column yet -> every updated
    script would crash with a KeyError on "Year" the moment it runs.
  - the 6 real data entry files on SharePoint are still named without a
    year -> integrate_data_entry.py would no longer find them (it now looks
    for the _2026 name), so anything a referent already typed in and hasn't
    been integrated yet would silently stop being picked up.

What this script does, in order:
  1. Reads input_data/Raw_data_CSR.xlsx (old schema) and rewrites it with
     Year=2026 added to every row (safe: every row in it today was
     necessarily collected during 2026, the only year this pipeline has
     ever run). The old file is backed up first
     (input_data/Raw_data_CSR_pre_year_migration_backup.xlsx) — never
     overwritten if that backup already exists, so running this twice by
     mistake never destroys the original.
  2. Renames each config.DATA_ENTRY_DIR/<BU>_data_entry_CSR.xlsm (if found)
     to <BU>_data_entry_CSR_2026.xlsm, in place on the SharePoint folder —
     preserves everything a referent has already typed in (including
     anything not yet integrated), and keeps the file referents already
     have open/bookmarked pointing at the same content, just under its new
     name. Skipped (with a message) for any BU where the old-named file
     doesn't exist, or the new-named one already does.

Usage
-----
    uv run scripts/migrate_2026_09_add_year.py

Dependencies: pandas, openpyxl.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
from integrate_data_entry import export_consolidated

YEAR = 2026


def migrate_raw_data():
    path = config.INPUT_DIR / "Raw_data_CSR.xlsx"
    if not path.exists():
        print(f"{path} not found — nothing to migrate (normal on a brand-new setup).")
        return

    df = pd.read_excel(path, sheet_name="Raw_data")
    if "Year" in df.columns:
        print(f"{path} already has a 'Year' column — nothing to do.")
        return

    backup_path = config.INPUT_DIR / "Raw_data_CSR_pre_year_migration_backup.xlsx"
    if backup_path.exists():
        print(f"Backup {backup_path.name} already exists — not overwriting it "
              f"(this migration may have partially run before; check manually if unsure).")
    else:
        df.to_excel(backup_path, sheet_name="Raw_data", index=False)
        print(f"Backed up the pre-migration file to {backup_path.name}.")

    df.insert(1, "Year", YEAR)  # right after "BU", matching the column order used everywhere else
    export_consolidated(df)
    print(f"{path} migrated: {len(df)} row(s) tagged Year={YEAR}.")


def migrate_data_entry_files():
    for bu in config.RAW_FILES:
        old_path = config.DATA_ENTRY_DIR / f"{bu}_data_entry_CSR.xlsm"
        new_path = config.DATA_ENTRY_DIR / f"{bu}_data_entry_CSR_{YEAR}.xlsm"
        if new_path.exists():
            print(f"{bu}: {new_path.name} already exists — skipped.")
            continue
        if not old_path.exists():
            print(f"{bu}: {old_path.name} not found — skipped (nothing to rename).")
            continue
        old_path.rename(new_path)
        print(f"{bu}: {old_path.name} -> {new_path.name}")


def main():
    print("Step 1/2 — Raw_data_CSR.xlsx")
    migrate_raw_data()
    print("\nStep 2/2 — data entry files")
    migrate_data_entry_files()
    print("\nDone. You can now use the pipeline scripts normally.")


if __name__ == "__main__":
    main()
