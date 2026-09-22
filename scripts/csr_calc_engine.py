"""
Centralized calculation engine — CSR reporting (Bioline Agrosciences)
===================================================================

Principle
---------
The indicator reference list (input_data/Indicator_reference_CSR.xlsx)
describes the list of indicators ONCE, their unit, who enters them, and for
calculated indicators, the formula (in plain language, "Formula
(documentation)" column). Verified identical across the 6 BUs (Viridaxis,
BAF, BFR, BIB, BUK, BUS) and the 12 months.

The raw data (input_data/Raw_data_CSR.xlsx, one consolidated file, all BUs
AND all years) contains only ENTERED values (Responsible != CSO with a
formula), in long format: one row per (BU, Year, Month, ID). Extracted with
parsing of malformed data cases (decimal comma, text) — see the "Data
quality note" column. The "Year" column (added 22/09/2026) is what lets the
pipeline keep running year after year without a new year's August silently
overwriting the previous one's.

This script is the ONLY place where the calculation logic is actually
executed (the FORMULAS dict below). Adding, changing, or removing a
calculated indicator happens here and nowhere else — no more touching an
Excel tab per month per BU.

To add a new calculated indicator:
  1. Add a row to the reference list (ID, Topic, KPI, Unit, Kind="calculated", ...).
  2. Add an entry to the FORMULAS dict below: "My.ID": lambda v: ...
  That's it — it will be calculated for every BU and every month on the next run.

Usage
-----
    uv run scripts/csr_calc_engine.py

Prerequisite: extract_reference_and_data.py must already have been run
(it produces the files in input_data/). The validation against the original
Excel files (validate_against_originals) is an optional one-off check: it
only runs if raw_data/ still contains the 6 original files.

Dependencies: pandas, openpyxl.
"""

import sys
from datetime import date
from pathlib import Path

import pandas as pd

# Robust import of config.py, whether this script is launched via
# `uv run path/to/file.py`, `python file.py` from any directory, or `python -m ...`.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import config
import consolidation_report

REFERENCE_PATH = config.INPUT_DIR / "Indicator_reference_CSR.xlsx"
RESULT_PATH = config.OUTPUT_DIR / "Consolidated_results_CSR.xlsx"

BU_LIST = list(config.RAW_FILES.keys())

MONTH_ORDER = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]

# KNOWN and documented discrepancies between the original file and the
# recalculation — the engine corrects these cases, this is not a defect on
# its part (see the verification report delivered with this prototype for
# the detail of each case). Only used by validate_against_originals (a
# one-off check).
# Keys stay (BU, Month, ID) without a Year — validate_against_originals only
# ever compares against HISTORICAL_VALIDATION_YEAR (2026), the one and only
# year the original raw files covered, so a Year element here would be pure
# redundancy.
KNOWN_CORRECTIONS = {
    ("Viridaxis", "June", "Ene.9"): "original Excel formula was broken (#REF!)",
    ("Viridaxis", "July", "Ene.9"): "original formula pointed to June (uncorrected copy-paste)",
    ("Viridaxis", "July", "Ene.11"): "cascading consequence of the Ene.9 error above (Ene.11 = Ene.9 / Env.1)",
    ("BFR", "March", "Ene.9"): "cell replaced by a frozen value (formula lost) in the original file",
    ("BFR", "March", "Ene.10"): "cell replaced by a frozen value (formula lost) in the original file",
    ("BFR", "March", "Ene.11"): "cell replaced by a frozen value (formula lost) in the original file",
    ("BFR", "March", "Was.3"): "cell replaced by a frozen value (formula lost) in the original file",
    ("BFR", "March", "Was.4"): "cell replaced by a frozen value (formula lost) in the original file",
    ("BFR", "March", "Saf.6"): "cell replaced by a frozen value (formula lost) in the original file",
    ("BFR", "March", "Saf.7"): "cell replaced by a frozen value (formula lost) in the original file",
    ("Viridaxis", "July", "Ref.1"): "hand-entered value (0) inconsistent with Ref.2-9 (a 23kg leak not reflected in the total)",
    ("BUK", "June", "Ref.1"): "hand-entered value (0) inconsistent with Ref.2-9 (a 1.5kg leak not reflected in the total)",
}

# The 6 original raw files (validate_against_originals) only ever covered
# 2026 — same constant as extract_reference_and_data.py's HISTORICAL_YEAR,
# duplicated here rather than imported to keep this script runnable on its
# own without depending on the one-time migration script.
HISTORICAL_VALIDATION_YEAR = 2026

# ---------------------------------------------------------------------------
# THE reference list of calculation formulas — the only business logic of
# the whole pipeline.
# ---------------------------------------------------------------------------
def z(x):
    """Like Excel in SUM()/a + or * operator: an empty cell (None/NaN)
    counts as 0 when it's a term in a sum/product, instead of cascading a
    "missing" value upward."""
    return 0 if x is None or (isinstance(x, float) and pd.isna(x)) else x


def _hours_worked(v):
    """Hours worked and paid: Saf.4.2 if provided, otherwise a flat estimate
    (Sales x working days x 8h) — reproduces exactly the original Excel
    formula (=IF(J37="", J2*J36*8, J37))."""
    saf42 = v("Saf.4.2")
    if saf42 in (None, "") or (isinstance(saf42, float) and pd.isna(saf42)):
        return z(v("Env.1")) * z(v("Saf.4.1")) * 8
    return saf42


FORMULAS = {
    "Wat.2": lambda v: z(v("Wat.1")) / v("Env.1"),
    "Ene.9": lambda v: (z(v("Ene.1")) + z(v("Ene.2")) + z(v("Ene.3")) + z(v("Ene.4")) + z(v("Ene.5"))
                         + z(v("Ene.6")) * z(v("Ene.6.1"))
                         + z(v("Ene.7")) * z(v("Ene.7.1"))
                         + z(v("Ene.8")) * z(v("Ene.7.1"))),
    "Ene.10": lambda v: (z(v("Ene.2")) + z(v("Ene.3")) + z(v("Ene.4"))) / v("Ene.9"),
    "Ene.11": lambda v: v("Ene.9") / v("Env.1"),
    "Ref.1": lambda v: sum(z(v(f"Ref.{i}")) for i in range(2, 10)),
    "Was.3": lambda v: z(v("Was.2")) / v("Was.1"),
    "Was.4": lambda v: z(v("Was.1")) / v("Env.1"),
    "Saf.6": lambda v: (z(v("Saf.2")) + z(v("Saf.3"))) / _hours_worked(v) * 1_000_000,
    "Saf.7": lambda v: z(v("Saf.5")) / _hours_worked(v) * 1_000,
}


def compute_bu_month(raw_values: dict) -> dict:
    """Computes every "calculated" indicator (see FORMULAS) for ONE BU and
    ONE month, from the entered values (raw_values). Some formulas depend on
    the result of another calculated formula (e.g. Ene.11 needs Ene.9
    already computed): the function therefore retries the formulas that
    failed, until every one that can be computed has been — the ones that
    truly cannot, for lack of input data, are simply left out of the result
    rather than crashing the script."""
    values = dict(raw_values)
    remaining = dict(FORMULAS)
    progress = True
    while remaining and progress:
        progress = False
        for kpi_id, formula in list(remaining.items()):
            try:
                values[kpi_id] = formula(lambda i: values[i])
                del remaining[kpi_id]
                progress = True
            except (KeyError, ZeroDivisionError, TypeError):
                continue
    return values


def load_reference_data() -> pd.DataFrame:
    """Loads the indicator reference list
    (input_data/Indicator_reference_CSR.xlsx), produced once by
    extract_reference_and_data.py — used here to look up the Topic/KPI/Unit
    of each indicator to add to the final result."""
    return pd.read_excel(REFERENCE_PATH, sheet_name="Reference")


def load_all_raw_data() -> pd.DataFrame:
    """A single consolidated file (all BUs combined, "BU" column already
    present) — see extract_reference_and_data.py for the rationale."""
    return pd.read_excel(config.INPUT_DIR / "Raw_data_CSR.xlsx", sheet_name="Raw_data")


def build_completion_table(raw: pd.DataFrame, reference: pd.DataFrame, current_year: int) -> pd.DataFrame:
    """Per-BU, per-month % completion of the indicators actually owned by the
    CSR referent (Kind == "input" AND Responsible == "CSR referent" — the
    same filter generate_data_entry_file.py uses to decide what goes in each
    BU's data entry file; deliberately narrower than "every input indicator",
    which would also include Safety/Finance/HR-owned ones that were never
    the referent's to fill in). Restricted to `current_year`: a prior year is
    closed history, not something to keep tracking active progress on, and
    mixing years into one BU/Month cell wouldn't mean anything anyway.

    Consolidates, across all 6 BUs in one matrix, what each BU's own
    "Tracking" tab already shows individually — no need to open 6 files to
    see who's behind. Every one of the 12 months is included (not just
    months already due): a future month simply reads 0%, same convention as
    the per-BU Tracking tab."""
    raw = raw[raw.Year == current_year]
    referent_ids = set(reference[(reference.Kind == "input") & (reference.Responsible == "CSR referent")]["ID"])
    total = len(referent_ids)

    filled = raw[raw.ID.isin(referent_ids) & raw["Value"].notna()]
    filled_counts = filled.groupby(["BU", "Month"]).size()

    rows = []
    for bu in sorted(raw["BU"].unique()):
        row = {"BU": bu}
        for month in MONTH_ORDER:
            count = filled_counts.get((bu, month), 0)
            row[month] = f"{round(100 * count / total)}%" if total else "0%"
        rows.append(row)
    return pd.DataFrame(rows)


def run(reference: pd.DataFrame, raw: pd.DataFrame) -> pd.DataFrame:
    """Computes the full reporting: for every BU, year and month present in
    the raw data, computes every "calculated" indicator from the entered
    values (via compute_bu_month — formulas never mix values from two
    different years, since the grouping is per (BU, Year, Month)), then
    gathers everything — entered AND calculated — into a single consolidated
    table, sorted by BU / year / month / indicator. This is the table that
    main() then writes to output_data/Consolidated_results_CSR.xlsx."""
    meta = reference.set_index("ID")[["Topic", "KPI", "Unit", "Kind"]]
    results = []
    for (bu, year, month), grp in raw.groupby(["BU", "Year", "Month"], sort=False):
        raw_values = dict(zip(grp["ID"], grp["Value"]))
        computed = compute_bu_month(raw_values)
        for kpi_id, value in computed.items():
            if kpi_id not in meta.index:
                continue
            row = meta.loc[kpi_id]
            results.append({
                "BU": bu, "Year": year, "Month": month, "ID": kpi_id,
                "Topic": row["Topic"], "KPI": row["KPI"], "Unit": row["Unit"],
                "Kind": row["Kind"], "Value": value,
            })
    out = pd.DataFrame(results)
    out["Month"] = pd.Categorical(out["Month"], categories=MONTH_ORDER, ordered=True)
    return out.sort_values(["BU", "Year", "Month", "ID"]).reset_index(drop=True)


def validate_against_originals(consolidated: pd.DataFrame) -> None:
    """One-off check: compares the recalculation against the values the
    Excel formulas already displayed in the original raw files. Optional —
    only runs if raw_data/ still contains those 6 files (not needed for a
    normal monthly run, only to re-validate the engine after a change to
    FORMULAS). Always compares against HISTORICAL_VALIDATION_YEAR (2026) —
    the only year those 6 original files ever covered — regardless of what
    other years now sit in `consolidated`."""
    import openpyxl

    missing = [bu for bu in BU_LIST if not (config.RAW_DATA_DIR / config.RAW_FILES[bu]).exists()]
    if missing:
        print(f"\n(Validation skipped: original raw files missing from {config.RAW_DATA_DIR}/ "
              f"for {', '.join(missing)} — normal for regular monthly use.)")
        return

    calc_ids = list(FORMULAS.keys())
    checked, real_mismatches, known = 0, [], []
    for bu in BU_LIST:
        path = config.RAW_DATA_DIR / config.RAW_FILES[bu]
        wb = openpyxl.load_workbook(path, data_only=True)
        id_row = {}
        ws_feb = wb["February"]
        for r in range(2, ws_feb.max_row + 1):
            idv = ws_feb.cell(row=r, column=2).value
            if idv:
                id_row[idv] = r
        for month in MONTH_ORDER:
            if month not in wb.sheetnames or month == "January":
                continue  # January never had the calculated rows in any file
            ws = wb[month]
            for kpi_id in calc_ids:
                row = id_row.get(kpi_id)
                if not row:
                    continue
                original = ws.cell(row=row, column=10).value
                if not isinstance(original, (int, float)):
                    continue  # empty or an error (#REF!, #DIV/0!): nothing to compare as-is
                mine_series = consolidated[(consolidated.BU == bu) & (consolidated.Year == HISTORICAL_VALIDATION_YEAR)
                                            & (consolidated.Month == month) & (consolidated.ID == kpi_id)]["Value"]
                if mine_series.empty:
                    continue
                mine = mine_series.iloc[0]
                checked += 1
                is_bad = pd.isna(mine) or abs(mine - original) > 1e-6
                if is_bad:
                    key = (bu, month, kpi_id)
                    if key in KNOWN_CORRECTIONS:
                        known.append((bu, month, kpi_id, original, mine, KNOWN_CORRECTIONS[key]))
                    else:
                        real_mismatches.append((bu, month, kpi_id, original, mine))

    print(f"\nValidation vs original files: {checked} values compared.")
    print(f"  {len(known)} expected correction(s) (broken formula in the source file, "
          f"documented):")
    for m in known:
        print(f"    {m[0]}/{m[1]}/{m[2]}: file={m[3]} -> recalculated={m[4]:.2f}  ({m[5]})")
    print(f"  {len(real_mismatches)} UNEXPLAINED discrepancy/discrepancies:")
    for m in real_mismatches:
        print("    MISMATCH:", m)


def main():
    """Entry point: loads the reference list and the consolidated raw data
    (all BUs, all years), computes every indicator for every BU/year/month,
    prints a summary (number of rows computed, a preview of the Ref.1
    indicator), also runs an optional consistency check against the old
    Excel files if they're still present, emails the anomaly report (see
    consolidation_report.py), then writes output_data/Consolidated_results_CSR.xlsx
    ("Results": every indicator/BU/month; "Completion": a BU x month %
    completion matrix for the CSR-referent-owned indicators, current year
    only — see build_completion_table). This is the script to run every
    month once the new entries have been integrated (via
    integrate_data_entry.py and integrate_safety_data.py)."""
    reference = load_reference_data()
    raw = load_all_raw_data()

    consolidated = run(reference, raw)
    print(f"{len(consolidated)} rows computed in total for {raw.BU.nunique()} BUs, "
          f"{raw.Year.nunique()} year(s) ({consolidated.Value.notna().sum()} non-empty values).")

    # Ref.1 recalculated vs sum of parts, before/after the recalculation (reference)
    ref1_check = consolidated[consolidated.ID == "Ref.1"][["BU", "Year", "Month", "Value"]]
    print("\nRef.1 recalculated (excerpt, non-empty values):")
    print(ref1_check[ref1_check.Value.notna()].to_string(index=False))

    validate_against_originals(consolidated)

    current_year = date.today().year
    current_month = MONTH_ORDER[date.today().month - 1]
    consolidation_report.run_and_send(raw, consolidated, reference, current_year, current_month)

    completion = build_completion_table(raw, reference, current_year)
    with pd.ExcelWriter(RESULT_PATH) as writer:
        consolidated.to_excel(writer, sheet_name="Results", index=False)
        completion.to_excel(writer, sheet_name="Completion", index=False)
    print(f"\nConsolidated results written to {RESULT_PATH} "
          f"(\"Results\" + \"Completion\" — % filled per BU/month for CSR-referent indicators, {current_year})")


if __name__ == "__main__":
    main()
