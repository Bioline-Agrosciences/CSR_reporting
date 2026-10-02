"""
ONE-TIME cleanup (23/09/2026) — run this ONCE, then never again (same spirit
as migrate_2026_09_remove_env1.py, this is not part of the recurring monthly
cycle).

Why: two more decisions on top of the Env.1 one:
  - Saf.4 ("FTE" / headcount) is no longer tracked anywhere in this
    pipeline — nothing computes or consumes it (see extract_reference_and_data.py's
    EXCLUDED_IDS).
  - Ene.6.1, Ene.7.1 (LPG/Fuel conversion factors) and Saf.4.1 (working days
    in the month) are reclassified from "input" (hand-typed every month) to
    "calculated" — csr_calc_engine.py now derives all 3 fresh every run, from
    config.py constants / a public-holiday calendar (see its
    CONTEXTUAL_VALUES), never from a value someone re-types.

Neither change is retroactively applied by just editing the scripts: the
Saf.4 row and the (now-stale) Ene.6.1/Ene.7.1/Saf.4.1 hand-typed rows are
still sitting in the real input_data/Indicator_reference_CSR.xlsx and
input_data/Raw_data_CSR.xlsx, written before these decisions.

What this script does, in order:
  1. Regenerates input_data/Indicator_reference_CSR.xlsx from scratch (via
     extract_reference_and_data.build_reference() + export_reference() —
     safe to do any time: this file is fully derived from the reference tab
     of the original raw workbooks, nothing accumulates into it over time
     the way Raw_data_CSR.xlsx does). Requires the original raw workbooks to
     still be reachable at config.RAW_DATA_DIR (they are, on the shared
     drive — see config.py). The old file is backed up first
     (Indicator_reference_CSR_pre_reclassification_backup.xlsx) — never
     overwritten if that backup already exists.
  2. Removes every row with ID in {"Saf.4", "Ene.6.1", "Ene.7.1", "Saf.4.1"}
     (any BU, any Year, any Month) from input_data/Raw_data_CSR.xlsx — Saf.4
     because it's excluded entirely, the other 3 because they're computed
     fresh every run now, not stored as history. Backed up first the same
     way (Raw_data_CSR_pre_reclassification_backup.xlsx).
Both steps are no-ops (printed, not silent) if a file is missing.

Usage
-----
    uv run scripts/migrate_2026_09_reclassify_conversion_factors.py

Dependencies: pandas, openpyxl.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
from extract_reference_and_data import build_reference, export_reference
from integrate_data_entry import export_consolidated

REMOVED_FROM_RAW_DATA = {"Saf.4", "Ene.6.1", "Ene.7.1", "Saf.4.1"}


def rebuild_reference():
    path = config.INPUT_DIR / "Indicator_reference_CSR.xlsx"
    backup_path = config.INPUT_DIR / "Indicator_reference_CSR_pre_reclassification_backup.xlsx"

    if path.exists():
        if backup_path.exists():
            print(f"Backup {backup_path.name} already exists — not overwriting it "
                  f"(this cleanup may have partially run before; check manually if unsure).")
        else:
            old_df = pd.read_excel(path, sheet_name="Reference")
            old_df.to_excel(backup_path, sheet_name="Reference", index=False)
            print(f"Backed up the pre-cleanup file to {backup_path.name}.")
    else:
        print(f"{path.name} not found — nothing to back up, building it fresh.")

    new_df = build_reference()
    export_reference(new_df)
    print(f"{path.name} regenerated: {len(new_df)} indicators "
          f"({(new_df.Kind == 'input').sum()} input / {(new_df.Kind == 'calculated').sum()} calculated).")


def clean_raw_data():
    path = config.INPUT_DIR / "Raw_data_CSR.xlsx"
    if not path.exists():
        print(f"{path} not found — nothing to clean.")
        return

    df = pd.read_excel(path, sheet_name="Raw_data")
    to_remove = df.ID.isin(REMOVED_FROM_RAW_DATA)
    removed_count = int(to_remove.sum())
    if removed_count == 0:
        print(f"{path.name} has none of {sorted(REMOVED_FROM_RAW_DATA)} — nothing to do.")
        return

    backup_path = config.INPUT_DIR / "Raw_data_CSR_pre_reclassification_backup.xlsx"
    if backup_path.exists():
        print(f"Backup {backup_path.name} already exists — not overwriting it "
              f"(this cleanup may have partially run before; check manually if unsure).")
    else:
        df.to_excel(backup_path, sheet_name="Raw_data", index=False)
        print(f"Backed up the pre-cleanup file to {backup_path.name}.")

    df = df[~to_remove].reset_index(drop=True)
    export_consolidated(df)
    print(f"{path.name}: {removed_count} row(s) removed ({sorted(REMOVED_FROM_RAW_DATA)}).")


def main():
    print("Step 1/2 — Indicator_reference_CSR.xlsx")
    rebuild_reference()
    print("\nStep 2/2 — Raw_data_CSR.xlsx")
    clean_raw_data()
    print("\nDone. You can now use the pipeline scripts normally.")


if __name__ == "__main__":
    main()
