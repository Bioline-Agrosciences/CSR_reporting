"""
ONE-TIME cleanup (25/09/2026) — run this ONCE, then never again (same spirit
as the other migrate_2026_09_*.py scripts, this is not part of the recurring
monthly cycle).

Why: Carb.1-5 (CO2 emissions from energy consumption) and Ene.12-15 (the
per-BU CO2 emission factors that feed them — electricity, natural gas, fuel,
LPG) were left out when this pipeline was first built. Not a newly-invented
metric: they're calculated indicators (Kind="calculated", computed by
csr_calc_engine.py's FORMULAS/CONTEXTUAL_VALUES) that were part of the
intended catalog from the start, simply omitted from the initial
extraction. Unlike every other row in the reference list, they can't be
read from the original raw workbooks — confirmed (25/09/2026, all 12
monthly tabs of the BAF file inspected directly) that none of these 9 IDs
appear anywhere in any of the 6 files at all. See
extract_reference_and_data.py's CARBON_INDICATORS for where they are
appended instead of read from a sheet.

What this script does: regenerates input_data/Indicator_reference_CSR.xlsx
from scratch (extract_reference_and_data.build_reference() now includes
these 9 rows on every call) — the old file is backed up first
(Indicator_reference_CSR_pre_carbon_indicators_backup.xlsx), never
overwritten if that backup already exists. Requires the original raw
workbooks to still be reachable at config.RAW_DATA_DIR, same as the other
migrate_2026_09_*.py scripts. No change needed to Raw_data_CSR.xlsx: none
of these 9 indicators are ever hand-entered, so there is nothing to remove
from it.

Usage
-----
    uv run scripts/migrate_2026_09_add_carbon_indicators.py

Dependencies: pandas, openpyxl.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
from extract_reference_and_data import CARBON_INDICATOR_IDS, build_reference, export_reference


def main():
    path = config.INPUT_DIR / "Indicator_reference_CSR.xlsx"
    backup_path = config.INPUT_DIR / "Indicator_reference_CSR_pre_carbon_indicators_backup.xlsx"

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
    added = new_df[new_df.ID.isin(CARBON_INDICATOR_IDS)]
    print(f"{path.name} regenerated: {len(new_df)} indicators total, "
          f"{len(added)} carbon/emission-factor indicator(s) added: {sorted(added.ID)}.")


if __name__ == "__main__":
    main()
