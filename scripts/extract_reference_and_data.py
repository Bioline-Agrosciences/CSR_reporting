"""
Extraction of the indicator reference list + the long-format raw data, from
the 12 monthly tabs of the original raw Excel files (one per BU).
===============================================================================

This script performs the "old format -> new format" conversion:
  - old format: 12 tabs per BU, mixing rows to be filled in and calculated
    rows, with the indicator reference list duplicated on every tab.
  - new format (output of this script, written to input_data/):
      1. Indicator_reference_CSR.xlsx  — the reference list, once, with each
         indicator's "Kind" (input / calculated) and, for calculated ones,
         the formula as documentation (see csr_calc_engine.py for the
         executable version of these formulas).
      2. Raw_data_CSR.xlsx — INPUT values only, for all BUs, in long format
         [BU, Year, Month, ID, Value, Comment], with a "Data quality note"
         column tracking anything that was automatically corrected (e.g.
         decimal comma) or still needs a manual fix (unparsable text).

How the reference list was built
---------------------------------
Verified beforehand (see the verification report delivered separately) that
the 6 BUs have EXACTLY the same catalog of 39 indicators, in the same order,
across the 12 months. The reference list is therefore extracted once from a
single reference tab (BAF, February — the most complete one), deduplicating
by ID. Each ID is classified "calculated" if it matches one of the 9 CSO
rows that, in the original file, contain an Excel calculation formula (e.g.
Ene.9 = total energy). Ref.1 ("Total refrigerants leaks") is a special case
handled separately: in the original files this row never had a formula (the
cell was always empty), but 3 real cases where a total had been entered by
hand matched exactly the sum of Ref.2 to Ref.9 — it is therefore reclassified
here as "calculated" = SUM(Ref.2..Ref.9), which also fills in several cases
where the total had been left empty despite real leaks recorded in detail.

Env.1 ("Sales") and Saf.4 ("FTE" / headcount) are EXCLUDED entirely
(decisions of 23/09/2026, see EXCLUDED_IDS below) — this pipeline never has
real sales/SAP figures to begin with (Env.1), and FTE is simply no longer
tracked or used anywhere in this pipeline (Saf.4). That leaves 37 of the
original 39 indicators, PLUS 9 more appended on top (CARBON_INDICATORS,
added 25/09/2026 — Ene.12-15 and Carb.1-5, CO2 emissions from energy
consumption, omitted from the initial extraction because they're absent
from the raw workbooks entirely, see CARBON_INDICATORS below): 46
indicators in total.

CALC_IDS carries TWO kinds of exception on top of the 9 original
CSO-formula rows:
  - Ref.1 (see above): never had a formula in the original files, hand-
    verified to equal SUM(Ref.2..Ref.9), reclassified "calculated" so a
    referent never re-types a total the pipeline can derive itself.
  - Ene.6.1, Ene.7.1, Saf.4.1 (decision of 23/09/2026): these WERE hand-
    typed "input" rows in the original files, but none of them is actually
    something anyone should be measuring or re-typing every month — Ene.6.1
    /Ene.7.1 are physical conversion constants (kWh per kg of LPG / per
    liter of fuel), Saf.4.1 is a public-holiday-aware working-days count
    that a calendar can derive on its own. All 3 are reclassified
    "calculated" and computed fresh every run by
    csr_calc_engine.py's CONTEXTUAL_VALUES (see that module — a different
    mechanism from FORMULAS, since these come from BU/year/month context,
    not from other entered values).

How the raw data was cleaned
-------------------------------
Loaded with data_only=True (openpyxl) to get a cell's already-computed value,
including when the cell itself is a formula typed by a referent to express
their entry (e.g. "=598089/1000" or "=5+36.5+6+12" — very common in the
original files, and no issue since only the result is needed). `parse_raw_value`
then handles the remaining cases:
  - a clean number -> kept as is;
  - text with a French decimal comma ("209,3") -> converted to a dot, with a
    traceability note (do NOT do this conversion silently anywhere else in
    the pipeline: it must stay visible);
  - standard numeric text -> converted;
  - unparsable text (e.g. "28,9kWh", which mixes a number and a unit)
    -> left empty, with a "NOT PARSABLE" note inviting a fix at the source
    rather than guessing a value.

Usage
-----
    uv run scripts/extract_reference_and_data.py

Prerequisite: the 6 original raw Excel files must be present in raw_data/
(see config.py for the expected names).

Dependencies: pandas, openpyxl.
"""

import re
import sys
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.styles import PatternFill, Font, Alignment, Protection
from openpyxl.utils import get_column_letter

# Robust import of config.py, whether this script is launched via
# `uv run path/to/file.py`, `python file.py` from any directory, or `python -m ...`.
sys.path.insert(0, str(Path(__file__).resolve().parent))
import config

# ---------------------------------------------------------------------------
# Configuration specific to this script (paths come from config.py)
# ---------------------------------------------------------------------------
REFERENCE_BU = "BAF"          # BU used to extract the reference list (doesn't matter
REFERENCE_MONTH = "February"  # which, as long as the catalog is identical everywhere)

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]

# The 6 original raw files given to us only ever covered 2026 (no year
# column in them at all — each was a single-year workbook). Hardcoded here
# rather than derived from today's date: this script is a one-time
# historical migration (see module docstring, "never run again" once done),
# not part of the recurring monthly cycle, so it must always tag its output
# as 2026 regardless of when it happens to be re-run.
HISTORICAL_YEAR = 2026

# The 9 rows whose "Monthly value" is a calculation formula (Responsible=CSO)
# in the original files, PLUS Ref.1, Ene.6.1, Ene.7.1, Saf.4.1 (see docstring
# above for why each of these last 4 is an exception, and of a different
# kind). Not just historical anymore, unlike the note that used to be here —
# Ene.6.1/Ene.7.1/Saf.4.1's reclassification IS a "stop treating this as
# input" decision, same spirit as Ref.1's.
CALC_IDS = {"Wat.2", "Ene.9", "Ene.10", "Ene.11", "Ref.1", "Was.3", "Was.4", "Saf.6", "Saf.7",
            "Ene.6.1", "Ene.7.1", "Saf.4.1"}

# Indicators this pipeline never tracks at all (decisions of 23/09/2026) —
# skipped entirely in build_reference(), so they never reach
# Indicator_reference_CSR.xlsx and, as a direct consequence, never reach
# Raw_data_CSR.xlsx either (build_raw_long only extracts IDs that made it
# into the reference list as entry_ids).
#   - Env.1 ("Sales"): this pipeline (CSR referents + BlueKanGo + Working
#     Hours) has no real sales/SAP figures — that figure is joined
#     downstream, in Fabric, where the real numbers live.
#   - Saf.4 ("FTE" / headcount): no longer tracked or used anywhere in this
#     pipeline — nothing computes or consumes it.
EXCLUDED_IDS = {"Env.1", "Saf.4"}

# CO2 emissions from energy consumption (25/09/2026): Ene.12-15 (per-BU CO2
# emission factors — electricity, natural gas, fuel, LPG) and Carb.1-5 (the
# emissions themselves, quantity * factor — see csr_calc_engine.py's
# FORMULAS/CONTEXTUAL_VALUES). These 9 indicators were part of the intended
# catalog from the start and simply got left out of the initial extraction
# — not a newly-invented metric. Unlike every other row in the reference
# list, they can't be read from a raw workbook: confirmed (25/09/2026, all
# 12 monthly tabs of the BAF file inspected directly) that none of these 9
# IDs appear anywhere in the original raw files at all — no ID, no topic,
# no KPI name. So they're appended here directly, all Kind="calculated",
# rather than looped from a sheet like the rest of build_reference().
CARBON_INDICATORS = [
    {"ID": "Ene.12", "Topic": "Energy", "KPI": "Electricity - FE", "Unit": "tCO2/kWh",
     "Kind": "calculated", "Responsible": "CSO",
     "Formula (documentation)": "Constant: config.ELECTRICITY_EMISSION_FACTOR[BU] "
                                 "(varies by country's grid mix)",
     "Definition": "CO2 emission factor for electricity, per BU's country grid mix.",
     "Calculation detail": "", "Consistency check": "",
     "Source of data": "Group CSR indicator framework (country grid emission factors), provided 25/09/2026.",
     "Notes": ""},
    {"ID": "Ene.13", "Topic": "Energy", "KPI": "Natural Gas - FE", "Unit": "tCO2/kWh",
     "Kind": "calculated", "Responsible": "CSO",
     "Formula (documentation)": "Constant: config.NATURAL_GAS_EMISSION_FACTOR[BU]",
     "Definition": "CO2 emission factor for natural gas.",
     "Calculation detail": "", "Consistency check": "",
     "Source of data": "Group CSR indicator framework, provided 25/09/2026.", "Notes": ""},
    {"ID": "Ene.14", "Topic": "Energy", "KPI": "Fioul - average FE", "Unit": "tCO2/L",
     "Kind": "calculated", "Responsible": "CSO",
     "Formula (documentation)": "Constant: config.FUEL_EMISSION_FACTOR[BU]",
     "Definition": "CO2 emission factor for fuel (fioul).",
     "Calculation detail": "", "Consistency check": "",
     "Source of data": "Group CSR indicator framework, provided 25/09/2026.", "Notes": ""},
    {"ID": "Ene.15", "Topic": "Energy", "KPI": "LPG - average FE", "Unit": "tCO2/kg",
     "Kind": "calculated", "Responsible": "CSO",
     "Formula (documentation)": "Constant: config.LPG_EMISSION_FACTOR[BU]",
     "Definition": "CO2 emission factor for LPG.",
     "Calculation detail": "", "Consistency check": "",
     "Source of data": "Group CSR indicator framework, provided 25/09/2026.", "Notes": ""},
    {"ID": "Carb.1", "Topic": "Energy", "KPI": "CO2 emissions - Electricity", "Unit": "tCO2",
     "Kind": "calculated", "Responsible": "CSO", "Formula (documentation)": "Ene.1 * Ene.12",
     "Definition": "CO2 emissions from electricity consumption.",
     "Calculation detail": "", "Consistency check": "",
     "Source of data": "Computed from Ene.1 and Ene.12.", "Notes": ""},
    {"ID": "Carb.2", "Topic": "Energy", "KPI": "CO2 emissions - Natural gas", "Unit": "tCO2",
     "Kind": "calculated", "Responsible": "CSO", "Formula (documentation)": "Ene.5 * Ene.13",
     "Definition": "CO2 emissions from natural gas consumption.",
     "Calculation detail": "", "Consistency check": "",
     "Source of data": "Computed from Ene.5 and Ene.13.", "Notes": ""},
    {"ID": "Carb.3", "Topic": "Energy", "KPI": "CO2 emissions - Fuel", "Unit": "tCO2",
     "Kind": "calculated", "Responsible": "CSO", "Formula (documentation)": "Ene.7 * Ene.14",
     "Definition": "CO2 emissions from fuel consumption.",
     "Calculation detail": "", "Consistency check": "",
     "Source of data": "Computed from Ene.7 and Ene.14.", "Notes": ""},
    {"ID": "Carb.4", "Topic": "Energy", "KPI": "CO2 emissions - LPG", "Unit": "tCO2",
     "Kind": "calculated", "Responsible": "CSO", "Formula (documentation)": "Ene.6 * Ene.15",
     "Definition": "CO2 emissions from LPG consumption.",
     "Calculation detail": "", "Consistency check": "",
     "Source of data": "Computed from Ene.6 and Ene.15.", "Notes": ""},
    {"ID": "Carb.5", "Topic": "Energy", "KPI": "Total CO2 emissions (energy)", "Unit": "tCO2",
     "Kind": "calculated", "Responsible": "CSO",
     "Formula (documentation)": "Carb.1 + Carb.2 + Carb.3 + Carb.4",
     "Definition": "Total CO2 emissions from energy consumption (electricity, natural gas, fuel, LPG).",
     "Calculation detail": "", "Consistency check": "",
     "Source of data": "Sum of Carb.1-4.", "Notes": ""},
]
CARBON_INDICATOR_IDS = {row["ID"] for row in CARBON_INDICATORS}

FORMULA_DESC = {
    "Wat.2": "Wat.1 / Env.1",
    "Ene.9": "Ene.1 + Ene.2 + Ene.3 + Ene.4 + Ene.5 + Ene.6*Ene.6.1 + Ene.7*Ene.7.1 + Ene.8*Ene.7.1",
    "Ene.10": "(Ene.2 + Ene.3 + Ene.4) / Ene.9",
    "Ene.11": "Ene.9 / Env.1",
    "Ref.1": "SUM(Ref.2..Ref.9)",
    "Was.3": "Was.2 / Was.1",
    "Was.4": "Was.1 / Env.1",
    "Saf.6": "(Saf.2+Saf.3) / (Saf.4.2 if provided, otherwise Env.1*Saf.4.1*8) * 1,000,000",
    "Saf.7": "Saf.5 / (Saf.4.2 if provided, otherwise Env.1*Saf.4.1*8) * 1,000",
    # NOTE — Wat.2, Ene.10, Ene.11, Was.3, Was.4, Saf.6, Saf.7: documentation
    # of the ORIGINAL Excel formula only. csr_calc_engine.py no longer
    # executes these (23/09/2026, ratios can't be summed to a group total) —
    # they're computed downstream instead, from the same raw components.
    "Ene.6.1": "Constant: config.LPG_CONVERSION_FACTOR[BU] (kWh per kg of LPG)",
    "Ene.7.1": "Constant: config.FUEL_CONVERSION_FACTOR[BU] (kWh per liter of fuel)",
    "Saf.4.1": "Working days in the month, public holidays excluded, per config.BU_COUNTRY[BU] "
               "(see csr_calc_engine.working_days_in_month)",
}
REFERENCE_NOTES = {
    "Ref.1": ("Corrected: never actually calculated in the original files (the cell was empty "
              "most months, even when Ref.2-9 held real leaks). Reclassified as a calculated "
              "indicator = sum of Ref.2 to Ref.9. Verified against 3 cases where a value had been "
              "entered by hand (BAF January, Viridaxis August, BIB August): matches the sum exactly "
              "in all 3 cases. 2 cases where the hand-entered value (0) contradicted the sum of the "
              "detail rows (Viridaxis July, BUK June) — likely a forgotten total update, see the "
              "verification report."),
}

FRENCH_DECIMAL_RE = re.compile(r"^\s*-?\d+,\d+\s*$")

ORANGE = "FFFFC000"
YELLOW_NOTE = "FFFFF2CC"
HEADER_FILL = "1F3864"


def parse_raw_value(v):
    """Returns (numeric_value_or_None, note) from a "Monthly value" cell
    already resolved by data_only=True (so formulas internal to the file are
    already numbers — only malformed text arrives here as a str)."""
    if v is None:
        return None, None
    if isinstance(v, (int, float)):
        return float(v), None
    if isinstance(v, str):
        s = v.strip()
        if FRENCH_DECIMAL_RE.match(s):
            return float(s.replace(",", ".")), "Corrected: decimal comma (“" + v + "”) -> dot"
        try:
            return float(s), "Corrected: numeric text -> number"
        except ValueError:
            return None, f"NOT PARSABLE, needs a manual fix in the source file: “{v}”"
    return None, f"Unexpected type ({type(v).__name__}): {v!r}"


def _raw_path(bu: str) -> Path:
    """Returns the path to the original raw Excel file for this BU (e.g.
    raw_data/Bioline monthly reporting - BAF.xlsx), and stops the script
    immediately with a clear message if this file doesn't exist — rather than
    a confusing error further down the code."""
    path = config.RAW_DATA_DIR / config.RAW_FILES[bu]
    if not path.exists():
        raise FileNotFoundError(
            f"Raw file not found for {bu}: {path}\n"
            f"-> Check that the 6 original Excel files are in {config.RAW_DATA_DIR}/"
        )
    return path


def build_reference() -> pd.DataFrame:
    """Builds the indicator reference list (a single one, valid for all 6 BUs)
    by reading the full indicator catalog from a SINGLE reference tab (BAF,
    February — see REFERENCE_BU/REFERENCE_MONTH). For each indicator, records
    whether it is to be filled in by a referent ("input") or calculated
    automatically ("calculated"), and for calculated indicators, records the
    formula in plain language (see FORMULA_DESC). IDs in EXCLUDED_IDS (Env.1,
    Saf.4) are skipped entirely — never written to the reference list at
    all. CARBON_INDICATORS (Ene.12-15, Carb.1-5) are then appended on top —
    they can't be read from the sheet like the rest, since they're absent
    from the raw workbooks entirely (see that constant's docstring). The
    result is then written to Indicator_reference_CSR.xlsx by
    export_reference."""
    wb_ref = openpyxl.load_workbook(_raw_path(REFERENCE_BU), data_only=False)
    ws_ref = wb_ref[REFERENCE_MONTH]
    rows = []
    for r in range(2, ws_ref.max_row + 1):
        idv = ws_ref.cell(row=r, column=2).value
        if not idv or idv in EXCLUDED_IDS:
            continue
        is_calc = idv in CALC_IDS
        rows.append({
            "ID": idv,
            "Topic": ws_ref.cell(row=r, column=1).value,
            "KPI": ws_ref.cell(row=r, column=3).value,
            "Unit": ws_ref.cell(row=r, column=9).value,
            "Kind": "calculated" if is_calc else "input",
            "Responsible": ws_ref.cell(row=r, column=4).value,
            "Formula (documentation)": FORMULA_DESC.get(idv, "") if is_calc else "",
            "Definition": ws_ref.cell(row=r, column=5).value,
            "Calculation detail": ws_ref.cell(row=r, column=6).value,
            "Consistency check": ws_ref.cell(row=r, column=7).value,
            "Source of data": ws_ref.cell(row=r, column=8).value,
            "Notes": REFERENCE_NOTES.get(idv, ""),
        })
    rows.extend(CARBON_INDICATORS)
    return pd.DataFrame(rows)


def build_raw_long(entry_ids: set) -> pd.DataFrame:
    """Walks through the 6 original raw Excel files (every month, every BU)
    and extracts only the values ENTERED by a referent (those whose ID is in
    entry_ids, so never calculated indicators): one row per (BU, Year, Month,
    Indicator), with the cleaned value (see parse_raw_value) and the
    referent's comment, if any. This table then becomes Raw_data_CSR.xlsx,
    the input database for the calculation engine (csr_calc_engine.py). Every
    row is tagged Year=HISTORICAL_YEAR (2026) — the only year these original
    files ever covered."""
    records = []
    for bu in config.RAW_FILES:
        wb = openpyxl.load_workbook(_raw_path(bu), data_only=True)
        for month in MONTHS:
            if month not in wb.sheetnames:
                continue
            ws = wb[month]
            for r in range(2, ws.max_row + 1):
                idv = ws.cell(row=r, column=2).value
                if idv not in entry_ids:
                    continue
                value, note = parse_raw_value(ws.cell(row=r, column=10).value)
                records.append({
                    "BU": bu, "Year": HISTORICAL_YEAR, "Month": month, "ID": idv, "Value": value,
                    "Comment": ws.cell(row=r, column=11).value,
                    "Data quality note": note,
                })
    df = pd.DataFrame(records)
    df["Month"] = pd.Categorical(df["Month"], categories=MONTHS, ordered=True)
    return df.sort_values(["BU", "Year", "Month", "ID"]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# Writing the output Excel workbooks (into input_data/)
# ---------------------------------------------------------------------------
def write_table(ws, df):
    """Writes a table (pandas DataFrame) into an Excel sheet: bold header row
    on a dark blue background, one row per data row, column widths adjusted
    to content, and the first row frozen so it stays visible while scrolling.
    Generic function reused by export_reference and export_raw_long — the
    formatting specific to each file (colors, cell locking...) is added on
    top, after this call."""
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")
    for row in df.itertuples(index=False):
        ws.append(list(row))
        # Source text starting with "=" (e.g. a description like "= (a/b)*1000")
        # would otherwise be mistaken for a formula by Excel/openpyxl.
        for c in ws[ws.max_row]:
            if isinstance(c.value, str) and c.value.startswith("="):
                c.data_type = "s"
    last_row = df.shape[0] + 1
    last_col = get_column_letter(df.shape[1])
    for col_cells in ws.columns:
        length = max((len(str(c.value)) if c.value is not None else 0) for c in col_cells)
        ws.column_dimensions[col_cells[0].column_letter].width = min(max(length + 2, 10), 60)
    ws.freeze_panes = "A2"
    return last_row, last_col


def export_reference(reference_df: pd.DataFrame):
    """Writes the indicator reference list to
    input_data/Indicator_reference_CSR.xlsx, as a structured Excel table
    (handy for sorting/filtering). This is THE reference file for the whole
    pipeline: adding, changing, or removing an indicator happens here, and
    generate_data_entry_file.py + csr_calc_engine.py pick it up automatically
    on their next run, with nothing else to change."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reference"
    last_row, last_col = write_table(ws, reference_df)
    tbl = Table(displayName="Indicator_Reference", ref=f"A1:{last_col}{last_row}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tbl)
    for r in range(2, last_row + 1):
        for c in range(1, reference_df.shape[1] + 1):
            ws.cell(row=r, column=c).alignment = Alignment(vertical="top", wrap_text=(c in (7, 8, 9, 10)))
    wb.save(config.INPUT_DIR / "Indicator_reference_CSR.xlsx")


def export_raw_long(raw_df: pd.DataFrame):
    """A single consolidated file (all BUs together, with a "BU" column
    first) rather than one file per BU: simpler to manage, and the
    calculation engine merges the BUs back into a single table as soon as it
    loads them anyway. Stays correct as long as re-integrating entries (see
    integrate_data_entry.py) is called as a SINGLE script that processes the
    BUs one after another (never several concurrent processes writing this
    same file in parallel — which this script does not do)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Raw_data"
    last_row, last_col = write_table(ws, raw_df)

    value_col = list(raw_df.columns).index("Value") + 1
    comment_col = list(raw_df.columns).index("Comment") + 1
    note_col = list(raw_df.columns).index("Data quality note") + 1

    # Same convention as the v2 workbook: everything locked except the
    # "Value" cell (orange = to be filled in) and "Comment".
    for r in range(2, last_row + 1):
        for c in range(1, raw_df.shape[1] + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)
        vcell = ws.cell(row=r, column=value_col)
        ccell = ws.cell(row=r, column=comment_col)
        vcell.protection = Protection(locked=False)
        ccell.protection = Protection(locked=False)
        vcell.fill = PatternFill(start_color=ORANGE, end_color=ORANGE, fill_type="solid")
        if ws.cell(row=r, column=note_col).value:
            ws.cell(row=r, column=note_col).fill = PatternFill(
                start_color=YELLOW_NOTE, end_color=YELLOW_NOTE, fill_type="solid")

    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                         showErrorMessage=True, errorTitle="Invalid value",
                         error="Please enter a number >= 0.")
    ws.add_data_validation(dv)
    dv.add(f"{get_column_letter(value_col)}2:{get_column_letter(value_col)}{last_row}")

    tbl = Table(displayName="Raw_Data_CSR", ref=f"A1:{last_col}{last_row}")
    tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tbl)

    ws.protection = SheetProtection(sheet=True, password=None, formatCells=False, formatColumns=False,
                                     formatRows=False, autoFilter=False, sort=False,
                                     selectLockedCells=False, selectUnlockedCells=False,
                                     objects=False, scenarios=False)

    wb.save(config.INPUT_DIR / "Raw_data_CSR.xlsx")


def main():
    """Entry point: builds the indicator reference list and the long-format
    raw data from the 6 original Excel files, prints a summary (indicator
    count, cells with a data quality note to review), then writes both
    output files to input_data/. Run once for the initial conversion (or if
    the original files change) — the regular monthly workflow goes through
    generate_data_entry_file.py and integrate_data_entry.py afterwards, not
    through this script."""
    reference_df = build_reference()
    entry_ids = set(reference_df[reference_df.Kind == "input"]["ID"])
    calculated_count = (reference_df.Kind == "calculated").sum()
    print(f"Reference list: {len(reference_df)} indicators "
          f"({len(entry_ids)} input / {calculated_count} calculated)")

    raw_df = build_raw_long(entry_ids)
    print(f"Raw data: {raw_df.shape[0]} rows across {raw_df.BU.nunique()} BUs")
    flagged = raw_df[raw_df["Data quality note"].notna()]
    print(f"\n{len(flagged)} cell(s) with a data quality note:")
    print(flagged[["BU", "Month", "ID", "Value", "Data quality note"]].to_string(index=False))

    export_reference(reference_df)
    export_raw_long(raw_df)
    print(f"\nFiles written to {config.INPUT_DIR}/: Indicator_reference_CSR.xlsx "
          f"+ Raw_data_CSR.xlsx (all BUs combined).")


if __name__ == "__main__":
    main()
