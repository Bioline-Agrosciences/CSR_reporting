"""
Generates each BU's data entry file for the CSR referent:
config.DATA_ENTRY_DIR/<BU>_data_entry_CSR_<year>.xlsm (one file per BU and
per year), written straight into the SharePoint folder the referents
already have access to. The Fabric notebook nb_consolidate_csr_data reads
these files back.

The workbook contains:
  - "Instructions" for the referent;
  - "Tracking": % completion per month;
  - one tab per month, listing only the indicators of the reference list
    with Kind = "input" and Responsible = "CSR referent". Only "Value of the
    month" and "Comment" are editable; values already known are pre-filled;
  - "Annual summary": all months side by side (Excel formulas, so it updates
    as the referent types), with month-over-month variations above
    config.VARIATION_THRESHOLD in red.

It is built on templates/data_entry_template.xlsm, whose Workbook_Open macro
recolors the current month's tab and recalculates "Tracking" every time the
file is opened, so the file stays correct for months without regenerating
it. Regenerate it when the reference list changes (an indicator added or
removed appears/disappears on the next run, for all BUs); a new year's file
starts automatically the first time it runs in a new year.

The "ID" column (first column) is what the notebook uses to read the values
back, so rows can be moved around without breaking anything.

Usage
-----
    uv run scripts/generate_data_entry_file.py              # all 6 BUs
    uv run scripts/generate_data_entry_file.py BAF BFR       # a selection

Dependencies: pandas, openpyxl.
"""
import sys
from datetime import date
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.protection import SheetProtection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill, Font, Alignment, Protection
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]

ORANGE = "FFFFC000"
HEADER_FILL = "1F3864"
CURRENT_TAB_COLOR = "FFC000"
ANNUAL_SUMMARY_TAB_COLOR = "375623"
VARIATION_FILL = "FFF4C7C3"

# .xlsm template carrying the macro (recolors the current month's tab and
# recalculates "Tracking" on open, with no need to regenerate the file) —
# see the Workbook_Open sub in its VBA project. Every data entry file is
# built on top of this template (openpyxl keep_vba=True) to carry the macro
# along with it. Part of the code (versioned in the repo, empty of any data),
# not of the SharePoint data folder.
TEMPLATE_PATH = config.TEMPLATE_DIR / "data_entry_template.xlsm"

# Context columns shown to the referent (read-only), in this order. "ID" is
# added on top, as the first column. nb_consolidate_csr_data (Cell 3) relies
# on this layout to find "Value of the month" and "Comment": update it too if
# this list changes.
DISPLAY_COLS = ["Topic", "KPI", "Unit", "Definition", "Calculation detail", "Source of data"]

# Column width (in characters) by context column name — lets "Calculation
# detail" reuse exactly the same width as "Source of data" instead of
# depending on its position in DISPLAY_COLS.
CONTEXT_COL_WIDTHS = {
    "Topic": 22, "KPI": 36, "Unit": 10, "Definition": 46,
    "Calculation detail": 22, "Source of data": 22,
}


def _is_filled(v) -> bool:
    """An empty Excel cell comes back sometimes as None, sometimes as NaN
    (float) depending on how pandas read it — this little helper avoids
    duplicating that check everywhere we need to know whether an indicator
    has a value (the % completion calculation in compute_completion, among
    others)."""
    return v is not None and not (isinstance(v, float) and pd.isna(v))


def load_data_entry_reference() -> pd.DataFrame:
    """Loads the indicator reference list and keeps only the rows to be
    filled in by the CSR referent themselves: Kind == "input" (never
    "calculated", formulas owned by the calculation engine, which a referent
    could overwrite) AND Responsible == "CSR referent" (the "input"
    indicators owned by another role — CSO, H&S manager, Finance, HR
    services — are filled in elsewhere, not in this file)."""
    ref = pd.read_excel(config.INPUT_DIR / "Indicator_reference_CSR.xlsx", sheet_name="Reference")
    return ref[(ref["Kind"] == "input") & (ref["Responsible"] == "CSR referent")].reset_index(drop=True)


def load_consolidated() -> pd.DataFrame:
    """input_data/Raw_data_CSR.xlsx — all BUs and all years, one row per
    (BU, Year, Month, ID). No longer updated since the consolidation moved
    to Fabric (the live data is the CSR_raw_data Delta table): it holds the
    history up to 02/10/2026."""
    path = config.INPUT_DIR / "Raw_data_CSR.xlsx"
    if not path.exists():
        return pd.DataFrame(columns=["BU", "Year", "Month", "ID", "Value", "Comment", "Data quality note"])
    return pd.read_excel(path, sheet_name="Raw_data")


def load_existing_values(bu: str, year: int, consolidated: pd.DataFrame) -> dict:
    """{(Month, ID): (Value, Comment)} already known for this BU, for THIS
    YEAR ONLY, merged from TWO sources:

    1. input_data/Raw_data_CSR.xlsx (see load_consolidated), filtered on
       this BU and this year (a prior year's August must never pre-fill this
       year's August tab).
    2. config.DATA_ENTRY_DIR/<BU>_data_entry_CSR_<year>.xlsm — the CURRENT
       data entry file for this year, if it already exists.

    Source 2 wins in case of conflict: regenerating the file must never
    erase what a referent has already typed. Values are copied as-is (even
    uncleaned text such as "56,7"): cleaning happens in the notebook.
    """
    values: dict = {}

    for row in consolidated[(consolidated.BU == bu) & (consolidated.Year == year)].itertuples():
        values[(row.Month, row.ID)] = (row.Value, row.Comment)

    entry_path = config.DATA_ENTRY_DIR / f"{bu}_data_entry_CSR_{year}.xlsm"
    if entry_path.exists():
        wb = openpyxl.load_workbook(entry_path, data_only=True)
        value_col = len(DISPLAY_COLS) + 2
        comment_col = value_col + 1
        for month in MONTHS:
            if month not in wb.sheetnames:
                continue
            ws = wb[month]
            for r in range(2, ws.max_row + 1):
                idv = ws.cell(row=r, column=1).value  # "ID" column
                if not idv:
                    continue
                value = ws.cell(row=r, column=value_col).value      # "Value of the month"
                comment = ws.cell(row=r, column=comment_col).value  # "Comment"
                if value is not None or comment is not None:
                    values[(month, idv)] = (value, comment)

    return values


def compute_completion(entry_records: list, existing: dict) -> list:
    """Computes, for each month, how many indicators out of the total
    already have a value filled in (existing) — used to fill the "Tracking"
    tab (completion rate per month), to see at a glance where data entry is
    still missing."""
    total = len(entry_records)
    completion = []
    for month in MONTHS:
        filled = sum(1 for rec in entry_records if _is_filled(existing.get((month, rec["ID"]), (None, None))[0]))
        completion.append((month, filled, total))
    return completion


def write_instructions_sheet(wb, bu: str, year: int):
    """Creates the very first tab of the data entry file, "Instructions": a
    simple text explaining to the CSR referent how to fill in the file
    (which cells to change, what each tab is for, etc.). No calculation
    here, only text and formatting."""
    ws = wb.create_sheet("Instructions")
    ws.column_dimensions["A"].width = 100
    lines = [
        f"Monthly CSR reporting — {bu} — {year}",
        "",
        f"This file is for {year} only — a new one, with {year + 1} in its name, will be "
        f"generated at the start of next year; this one stays as an archive, nothing is "
        f"overwritten.",
        "",
        "How to fill in this file:",
        "- This file contains a macro: when you open it, Excel may show a security bar "
        "(\"Macros have been disabled\" / \"Enable Content\") — click it, otherwise the tab "
        "colors and the \"Tracking\" tab won't update.",
        "- One tab per month. The tab for the current month is colored orange — this updates "
        "automatically every time you open the file, based on today's date, even if you keep "
        "using the same file for months without receiving a new one.",
        "- Only the \"Value of the month\" and \"Comment\" cells are editable (orange background) — "
        "everything else is locked.",
        "- Values already entered in previous months are pre-filled: check them and correct if "
        "needed, no need to re-enter everything.",
        "- A value must be a number >= 0. If you don't know the exact value yet, leave the cell "
        "empty rather than writing an approximate text (e.g. \"about 30\") — it won't be taken "
        "into account automatically.",
        "- Automatically calculated indicators (totals, ratios) are not in this file: they are "
        "recalculated from your entries, you have nothing to do for them.",
        "- The \"Tracking\" tab shows the completion rate per month — it also refreshes "
        "automatically each time you open the file.",
        f"- The last tab, \"Annual summary\", lists all months side by side for each "
        f"indicator, with large variations (more than {round(config.VARIATION_THRESHOLD * 100)}% "
        f"from one month to the next) highlighted in red. It updates automatically as soon as you "
        f"enter a value, no need to wait for a new version of the file to be sent.",
        "",
        "This file is regenerated periodically from the indicator reference list: if an indicator "
        "is added or removed, it will automatically appear/disappear the next time the file is "
        "sent — you don't need to do anything on your side.",
    ]
    for line in lines:
        ws.append([line])
    for r in range(1, len(lines) + 1):
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
    ws.cell(row=1, column=1).font = Font(bold=True, size=14)
    ws.sheet_properties.tabColor = "1F3864"


def write_tracking_sheet(wb, completion: list):
    """Creates the "Tracking" tab: a summary table with, for each month, the
    number of indicators already filled in out of the total and the
    corresponding % completion (computed beforehand by compute_completion).
    Lets you immediately see which month(s) have fallen behind, without
    having to open every monthly tab one by one."""
    ws = wb.create_sheet("Tracking")
    ws.append(["Month", "Indicators filled", "Total indicators", "% completed"])
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")
    for month, filled, total in completion:
        pct = round(100 * filled / total) if total else 0
        ws.append([month, filled, total, f"{pct}%"])
    for col, width in zip("ABCD", [16, 22, 16, 12]):
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A2"


def write_month_sheet(wb, month: str, entry_records: list, existing: dict, is_current: bool):
    """Creates the tab for a given month: one row per indicator to fill in,
    with its context (Topic, KPI, Unit, Definition, Calculation detail,
    Source of data) read-only, and two editable columns ("Value of the
    month", "Comment") pre-filled with the already-known value where
    applicable (existing). Locks every other cell, colors the tab orange if
    it's the current month (is_current), and requires an entered value to
    be a number >= 0."""
    ws = wb.create_sheet(month)
    if is_current:
        ws.sheet_properties.tabColor = CURRENT_TAB_COLOR

    headers = ["ID"] + DISPLAY_COLS + ["Value of the month", "Comment"]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")

    for rec in entry_records:
        value, comment = existing.get((month, rec["ID"]), (None, None))
        ws.append([rec["ID"]] + [rec[c] for c in DISPLAY_COLS] + [value, comment])

    last_row = len(entry_records) + 1
    id_col = 1
    value_col = len(DISPLAY_COLS) + 2
    comment_col = value_col + 1

    for r in range(2, last_row + 1):
        for c in range(1, comment_col + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)
        vcell = ws.cell(row=r, column=value_col)
        ccell = ws.cell(row=r, column=comment_col)
        vcell.protection = Protection(locked=False)
        ccell.protection = Protection(locked=False)
        vcell.fill = PatternFill(start_color=ORANGE, end_color=ORANGE, fill_type="solid")
        for c in range(2, value_col):
            ws.cell(row=r, column=c).alignment = Alignment(vertical="top", wrap_text=True)

    dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                         showErrorMessage=True, errorTitle="Invalid value",
                         error="Please enter a number >= 0.")
    ws.add_data_validation(dv)
    dv.add(f"{get_column_letter(value_col)}2:{get_column_letter(value_col)}{last_row}")

    ws.column_dimensions[get_column_letter(id_col)].width = 10
    widths = {c + 2: CONTEXT_COL_WIDTHS[name] for c, name in enumerate(DISPLAY_COLS)}
    widths[value_col] = 14
    widths[comment_col] = 32
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "B2"
    ws.row_dimensions[1].height = 30

    ws.protection = SheetProtection(sheet=True, password=None, formatCells=False, formatColumns=False,
                                     formatRows=False, autoFilter=False, sort=False,
                                     selectLockedCells=False, selectUnlockedCells=False,
                                     objects=False, scenarios=False)


def write_annual_summary_sheet(wb, entry_records: list):
    """Last tab, requested by the CSR referents: all months side by side for
    each indicator, to easily spot large variations.

    This is NOT a Python snapshot frozen at generation time: every month
    cell is an Excel formula (e.g. ='September'!G5) pointing to the "Value of
    the month" cell of the same indicator in the corresponding tab — since
    the 12 monthly tabs share exactly the same row order (one indicator = one
    row, identical everywhere), row r of the Annual summary always
    corresponds to row r of every monthly tab. Result: the Annual summary
    updates itself in Excel/LibreOffice as soon as the referent types a
    value, with no need to regenerate the file.

    Conditional formatting highlights, for each month from the second one
    onward, a variation of more than config.VARIATION_THRESHOLD compared to
    the previous month — the same logic as the original BAF v2 workbook.
    """
    ws = wb.create_sheet("Annual summary")
    ws.sheet_properties.tabColor = ANNUAL_SUMMARY_TAB_COLOR

    context_cols = ["Topic", "KPI", "Unit"]
    headers = ["ID"] + context_cols + MONTHS
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color=HEADER_FILL, end_color=HEADER_FILL, fill_type="solid")

    # "Value of the month" column in EVERY monthly tab (see write_month_sheet):
    # same position everywhere, so only one column letter needs computing.
    month_value_col_letter = get_column_letter(len(DISPLAY_COLS) + 2)

    for i, rec in enumerate(entry_records):
        row = i + 2
        row_values = [rec["ID"]] + [rec[c] for c in context_cols]
        # IF(...="","",...) rather than a direct reference: otherwise a month
        # not filled in yet would come back as "0" (Excel displays 0 for a
        # reference to an empty cell), which would trigger a false large-
        # variation alert (e.g. 1234 followed by an unfilled "0" = wrongly
        # detected as -100%).
        row_values += [f'=IF(\'{month}\'!{month_value_col_letter}{row}="","",\'{month}\'!{month_value_col_letter}{row})'
                       for month in MONTHS]
        ws.append(row_values)

    last_row = len(entry_records) + 1
    first_month_col = len(context_cols) + 2   # after ID + Topic/KPI/Unit
    last_month_col = first_month_col + len(MONTHS) - 1

    for r in range(2, last_row + 1):
        for c in range(1, last_month_col + 1):
            ws.cell(row=r, column=c).protection = Protection(locked=True)
        for c in range(2, len(context_cols) + 2):
            ws.cell(row=r, column=c).alignment = Alignment(vertical="top", wrap_text=True)

    # Highlight variations above the threshold, month by month, compared to
    # the previous month. ISNUMBER (rather than <>"") cleanly excludes months
    # not filled in yet (which contain "", not a number) without risking a
    # #VALUE! error if either side isn't numeric.
    fill = PatternFill(start_color=VARIATION_FILL, end_color=VARIATION_FILL, fill_type="solid")
    for col in range(first_month_col + 1, last_month_col + 1):
        cur, prev = get_column_letter(col), get_column_letter(col - 1)
        formula = (f'AND(ISNUMBER({prev}2),ISNUMBER({cur}2),{prev}2<>0,'
                   f'ABS({cur}2-{prev}2)/ABS({prev}2)>{config.VARIATION_THRESHOLD})')
        ws.conditional_formatting.add(f"{cur}2:{cur}{last_row}", FormulaRule(formula=[formula], fill=fill))

    ws.column_dimensions["A"].width = 10
    widths = {2: 22, 3: 34, 4: 10}
    for col in range(first_month_col, last_month_col + 1):
        widths[col] = 11
    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = get_column_letter(first_month_col) + "2"
    ws.row_dimensions[1].height = 30

    ws.protection = SheetProtection(sheet=True, password=None, formatCells=False, formatColumns=False,
                                     formatRows=False, autoFilter=False, sort=False,
                                     selectLockedCells=False, selectUnlockedCells=False,
                                     objects=False, scenarios=False)


def generate_for_bu(bu: str, data_entry_reference: pd.DataFrame, consolidated: pd.DataFrame,
                     year: int = None):
    """Builds and saves config.DATA_ENTRY_DIR/<BU>_data_entry_CSR_<year>.xlsm:
    gathers already-known values for that year, then assembles the tabs in
    order (Instructions, Tracking, one tab per month, Annual summary) and
    prints a summary (indicator count, completion of the current month).

    Built on top of TEMPLATE_PATH (keep_vba=True) rather than an empty
    workbook, so the Workbook_Open macro is always present in the file."""
    year = year or date.today().year
    entry_records = data_entry_reference.to_dict("records")
    existing = load_existing_values(bu, year, consolidated)
    current_month = MONTHS[date.today().month - 1]
    completion = compute_completion(entry_records, existing)

    wb = openpyxl.load_workbook(TEMPLATE_PATH, keep_vba=True)
    for sheet_name in list(wb.sheetnames):
        wb.remove(wb[sheet_name])

    write_instructions_sheet(wb, bu, year)
    write_tracking_sheet(wb, completion)
    for month in MONTHS:
        write_month_sheet(wb, month, entry_records, existing, is_current=(month == current_month))
    write_annual_summary_sheet(wb, entry_records)

    out_path = config.DATA_ENTRY_DIR / f"{bu}_data_entry_CSR_{year}.xlsm"
    wb.save(out_path)
    _, filled_current, total = next(c for c in completion if c[0] == current_month)
    print(f"{bu}: {out_path.name} generated ({len(entry_records)} indicators to fill in; "
          f"current month = {current_month}, {filled_current}/{total} already filled in)")


def main():
    """Entry point: generates the data entry file of the current year for the
    BUs given on the command line (all of them by default)."""
    requested = sys.argv[1:] or config.BU_LIST
    year = date.today().year
    data_entry_reference = load_data_entry_reference()
    consolidated = load_consolidated()
    for bu in requested:
        generate_for_bu(bu, data_entry_reference, consolidated, year=year)
    print(f"\nFiles written to {config.DATA_ENTRY_DIR}/ (year {year})")


if __name__ == "__main__":
    main()
