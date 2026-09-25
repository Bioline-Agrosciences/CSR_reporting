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

ONLY ADDITIVE (sum) FORMULAS LIVE HERE — no ratios, no divisions (decision
of 23/09/2026). This pipeline computes results at BU level AND rolls them
up to group level (all 6 BUs together): a sum rolls up correctly (total
energy for the group = sum of each BU's total energy), but a ratio does
NOT — averaging or summing 6 BUs' "% renewable energy" or "frequency rate"
produces a number that looks plausible but is mathematically meaningless
(the classic "average of averages" mistake: it ignores each BU's weight/
denominator). Ratio/intensity indicators — Wat.2, Ene.10, Ene.11, Was.3,
Was.4, Saf.6, Saf.7 — are therefore NOT computed here. They still exist in
the reference list (Kind="calculated", nobody types them in by hand), but
whoever needs them recomputes them downstream, FROM THE RAW COMPONENTS,
at whatever level of aggregation is actually meaningful (e.g. as Power BI
DAX measures — see the architecture notes shared with Aurélie). Producing
them here, at BU level only, would invite exactly the mistake this
decision avoids: someone summing or averaging them later without
realizing they can't be.

Env.1 ("Sales") is entirely absent from this pipeline for the same
family of reasons, plus a practical one: this pipeline (CSR referents +
BlueKanGo + Working Hours) never has real sales/SAP figures to begin with.
It is not read, not stored, not carried over from history — see
extract_reference_and_data.py for where it's excluded at the source, and
where a group-level Sales figure actually gets joined in (downstream, in
Fabric, alongside the ratios above).

A THIRD category, alongside FORMULAS (sums of already-entered values) and
"not computed at all" (the ratios above): CONTEXTUAL_VALUES, for the 3
indicators that were historically hand-typed every month (Ene.6.1, Ene.7.1,
Saf.4.1) but are neither something a referent should have to enter NOR a
sum of other entered values — they come from context that's already known
without anyone entering anything (a physical constant, or a public-holiday
calendar), see CONTEXTUAL_VALUES below for why this needs its own
mechanism (a FORMULAS lambda has no idea which BU or month it's running
for; these do, and that's exactly what they need).

To add a new (additive) calculated indicator:
  1. Add a row to the reference list (ID, Topic, KPI, Unit, Kind="calculated", ...).
  2. Add an entry to the FORMULAS dict below: "My.ID": lambda v: ... — a
     sum of z(v(...)) terms, never a division.
  That's it — it will be calculated for every BU and every month on the next run.

To add a new contextual value (a constant or something derivable from
BU/year/month alone, never from other entered values): add an entry to
CONTEXTUAL_VALUES instead — "My.ID": lambda bu, year, month: ...

Usage
-----
    uv run scripts/csr_calc_engine.py

Prerequisite: extract_reference_and_data.py must already have been run
(it produces the files in input_data/). The validation against the original
Excel files (validate_against_originals) is an optional one-off check: it
only runs if raw_data/ still contains the 6 original files.

Dependencies: pandas, openpyxl.
"""

import calendar
import sys
from datetime import date
from pathlib import Path

import holidays as holidays_lib
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
# one-off check). Only entries for IDs still in FORMULAS matter here (the
# ratio indicators removed from FORMULAS on 23/09/2026 — Ene.10, Ene.11,
# Was.3, Was.4, Saf.6, Saf.7 — had entries of their own; removed along with
# them, since validate_against_originals can no longer compute those IDs to
# compare in the first place).
# Keys stay (BU, Month, ID) without a Year — validate_against_originals only
# ever compares against HISTORICAL_VALIDATION_YEAR (2026), the one and only
# year the original raw files covered, so a Year element here would be pure
# redundancy.
KNOWN_CORRECTIONS = {
    ("Viridaxis", "June", "Ene.9"): "original Excel formula was broken (#REF!)",
    ("Viridaxis", "July", "Ene.9"): "original formula pointed to June (uncorrected copy-paste)",
    ("BFR", "March", "Ene.9"): "cell replaced by a frozen value (formula lost) in the original file",
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


FORMULAS = {
    "Ene.9": lambda v: (z(v("Ene.1")) + z(v("Ene.2")) + z(v("Ene.3")) + z(v("Ene.4")) + z(v("Ene.5"))
                         + z(v("Ene.6")) * z(v("Ene.6.1"))
                         + z(v("Ene.7")) * z(v("Ene.7.1"))
                         + z(v("Ene.8")) * z(v("Ene.7.1"))),
    "Ref.1": lambda v: sum(z(v(f"Ref.{i}")) for i in range(2, 10)),
}


def working_days_in_month(country_code: str, year: int, month: int) -> int:
    """Number of weekdays (Mon-Fri) in a given month that are NOT a public
    holiday in `country_code` (ISO 3166-1 alpha-2, e.g. "FR", "KE") —
    computed fresh from the `holidays` package's calendar every time, never
    hand-typed or stored. This is Saf.4.1's new definition (23/09/2026): the
    historical hand-entered values (see config.BU_COUNTRY) were essentially
    identical across every BU regardless of country, which this replaces
    with an actually country-aware count."""
    country_holidays = holidays_lib.country_holidays(country_code, years=year)
    _, days_in_month = calendar.monthrange(year, month)
    return sum(
        1 for day in range(1, days_in_month + 1)
        if date(year, month, day).weekday() < 5 and date(year, month, day) not in country_holidays
    )


# Indicators computed from CONTEXT (which BU, which year/month) rather than
# from other entered values (that's FORMULAS' job) or not computed here at
# all (the ratios, see module docstring). Added 23/09/2026, replacing what
# used to be 3 separate "input" indicators nobody was actually keeping
# up to date every month (Ene.6.1, Ene.7.1: a physical constant that never
# changes; Saf.4.1: a public-holiday calendar count, previously approximated
# by hand and identically across every BU regardless of country). Injected
# into each (BU, Year, Month) group's raw_values BEFORE compute_bu_month
# runs (see run()) — so FORMULAS entries that already reference them (Ene.9
# uses Ene.6.1/Ene.7.1) keep working unchanged, now resolving to these
# instead of to a hand-typed value.
CONTEXTUAL_VALUES = {
    "Ene.6.1": lambda bu, year, month: config.LPG_CONVERSION_FACTOR[bu],
    "Ene.7.1": lambda bu, year, month: config.FUEL_CONVERSION_FACTOR[bu],
    # `month` arrives as a name ("January"), same as everywhere else in this
    # pipeline — working_days_in_month needs the 1-12 number instead.
    "Saf.4.1": lambda bu, year, month: working_days_in_month(
        config.BU_COUNTRY[bu], year, MONTH_ORDER.index(month) + 1),
}


def compute_bu_month(raw_values: dict) -> dict:
    """Computes every "calculated" indicator (see FORMULAS) for ONE BU and
    ONE month, from the entered values (raw_values). A formula can depend on
    the result of another calculated formula (none currently do, now that
    FORMULAS only holds sums — but the retry loop stays generic in case a
    future addition ever needs it): the function retries the formulas that
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
    the per-BU Tracking tab.

    Long/tidy format: one row per (BU, Month), with BU and Month as plain
    columns (not one column per month) so a PivotTable or Power BI can
    filter and group by either one directly."""
    raw = raw[raw.Year == current_year]
    referent_ids = set(reference[(reference.Kind == "input") & (reference.Responsible == "CSR referent")]["ID"])
    total = len(referent_ids)

    filled = raw[raw.ID.isin(referent_ids) & raw["Value"].notna()]
    filled_counts = filled.groupby(["BU", "Month"]).size()

    rows = []
    for bu in sorted(raw["BU"].unique()):
        for month in MONTH_ORDER:
            count = filled_counts.get((bu, month), 0)
            rows.append({
                "BU": bu,
                "Month": month,
                "completion_percent": round(100 * count / total) if total else 0,
            })
    return pd.DataFrame(rows)


def run(reference: pd.DataFrame, raw: pd.DataFrame) -> pd.DataFrame:
    """Computes the full reporting: for every BU, year and month present in
    the raw data, computes every "calculated" indicator from the entered
    values (via compute_bu_month — formulas never mix values from two
    different years, since the grouping is per (BU, Year, Month)), then
    gathers everything — entered, calculated, AND contextual (see
    CONTEXTUAL_VALUES) — into a single consolidated table, sorted by BU /
    year / month / indicator. This is the table that main() then writes to
    output_data/Consolidated_results_CSR.xlsx."""
    meta = reference.set_index("ID")[["Topic", "KPI", "Unit", "Kind"]]
    results = []
    for (bu, year, month), grp in raw.groupby(["BU", "Year", "Month"], sort=False):
        raw_values = dict(zip(grp["ID"], grp["Value"]))
        for kpi_id, contextual_fn in CONTEXTUAL_VALUES.items():
            raw_values[kpi_id] = contextual_fn(bu, year, month)
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
