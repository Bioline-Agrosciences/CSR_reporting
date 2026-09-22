"""
Builds and emails the anomaly report sent after every csr_calc_engine.py
run. Three checks, each mapped to a concrete, already-established rule
rather than a vague "does this look weird" judgment call:

  1. Data quality corrections/flags — every row with a "Data quality note"
     (decimal comma auto-corrected, or unparsable text) from the historical
     migration or a later monthly integration.
  2. Missing values — "input" indicators with no value at all, for every
     month up to and including the current one (never flags a future
     month — nobody owes that data yet). Covers the full year, not just
     the current month: a gap from three months ago that's still unfilled
     stays visible until it's fixed.
  3. Large variations — a month-over-month jump bigger than
     config.VARIATION_THRESHOLD, the same rule already used to highlight
     cells red in the data entry file's "Annual summary" tab, applied here
     to the full consolidated result (entered AND calculated indicators),
     across the whole year.

Checks 2 and 3 run over the full year, which can turn up hundreds of rows —
too much to read inline in an email. The email body only carries a per-BU
COUNT summary, rendered as an actual HTML table (not a monospace text dump —
see build_email_html); the full row-by-row detail goes into an Excel
attachment (output_data/Anomaly_report_CSR.xlsx) instead.

Controlled by config.SEND_ERROR_REPORT_EMAIL (set to False to stop sending
it) and config.REPORT_RECIPIENTS (who gets it).

Sending goes through the local Outlook desktop app (COM automation), not
SMTP — no email password is stored anywhere in this project. This only
works on Windows, and ONLY IF OUTLOOK IS ALREADY RUNNING (see
send_report_email's docstring for why it never tries to launch Outlook
itself) — if it isn't, the report is still printed to the console and the
detail workbook still gets written, just not emailed. The import is lazy
(inside send_report_email) so importing this module — and running the
tests — works fine everywhere else.
"""
import sys
from datetime import date
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config

MONTH_ORDER = ["January", "February", "March", "April", "May", "June", "July",
               "August", "September", "October", "November", "December"]


def find_data_quality_notes(raw: pd.DataFrame) -> pd.DataFrame:
    """Every row with a non-empty "Data quality note", across every year
    present in `raw` (a flag from a prior year is still worth a look, so
    this one — unlike find_missing_values — is NOT restricted to the
    current year)."""
    return raw[raw["Data quality note"].notna()][["BU", "Year", "Month", "ID", "Value", "Data quality note"]]


def find_missing_values(raw: pd.DataFrame, input_ids: set, current_year: int, current_month: str) -> pd.DataFrame:
    """Every (BU, Month, ID) combination, for CURRENT_YEAR and months up to
    and including current_month, where an "input" indicator has no value at
    all — either the row is missing entirely, or its Value is empty. Months
    after current_month are never flagged, and so is any other year: a prior
    year's raw rows are excluded up front, otherwise a completed year would
    still get flagged forever as "missing" for the months it never had
    (BU/indicator combinations that started after that year, for instance),
    and a future year's early rows (if any exist ahead of schedule) would be
    compared against due_months that don't apply to them."""
    raw = raw[raw.Year == current_year]
    due_months = MONTH_ORDER[:MONTH_ORDER.index(current_month) + 1]
    bus = sorted(raw["BU"].unique())
    if not bus or not input_ids:
        return pd.DataFrame(columns=["BU", "Month", "ID"])

    expected = pd.MultiIndex.from_product([bus, due_months, sorted(input_ids)],
                                           names=["BU", "Month", "ID"]).to_frame(index=False)
    have_value = raw[raw.Month.isin(due_months) & raw.ID.isin(input_ids) & raw["Value"].notna()]
    merged = expected.merge(have_value[["BU", "Month", "ID"]], on=["BU", "Month", "ID"],
                             how="left", indicator=True)
    missing = merged[merged["_merge"] == "left_only"][["BU", "Month", "ID"]].copy()
    missing["Month"] = pd.Categorical(missing["Month"], categories=MONTH_ORDER, ordered=True)
    return missing.sort_values(["BU", "Month", "ID"]).reset_index(drop=True)


def find_large_variations(consolidated: pd.DataFrame) -> pd.DataFrame:
    """Every (BU, Year, ID) month-over-month jump bigger than
    config.VARIATION_THRESHOLD. Grouped by Year too (not just BU/ID) so
    December of one year is never compared to January of the next as if it
    were a normal month-over-month step — a year boundary is exactly the
    kind of place a real jump (annual reset, new pricing, etc.) is expected
    and shouldn't trigger a false anomaly."""
    rows = []
    for (bu, year, kpi_id), grp in consolidated.groupby(["BU", "Year", "ID"], sort=False):
        grp = grp.sort_values("Month")
        prev_value, prev_month = None, None
        for _, r in grp.iterrows():
            value = r["Value"]
            if pd.notna(value) and prev_value is not None and prev_value != 0:
                variation = abs(value - prev_value) / abs(prev_value)
                if variation > config.VARIATION_THRESHOLD:
                    rows.append({
                        "BU": bu, "Year": year, "ID": kpi_id,
                        "From": prev_month, "From value": prev_value,
                        "To": r["Month"], "To value": value,
                        "Variation": f"{variation:.0%}",
                    })
            if pd.notna(value):
                prev_value, prev_month = value, r["Month"]
    return pd.DataFrame(rows, columns=["BU", "Year", "ID", "From", "From value", "To", "To value", "Variation"])


def build_summary_table(quality_notes: pd.DataFrame, missing_values: pd.DataFrame,
                         large_variations: pd.DataFrame, bus: list) -> pd.DataFrame:
    """One row per BU (every BU passed in `bus`, even ones with zero
    anomalies — a clean BU is worth seeing too), with a count for each of
    the three checks."""
    table = pd.DataFrame({"BU": bus})
    table["Data quality"] = table["BU"].map(quality_notes.groupby("BU").size()).fillna(0).astype(int)
    table["Missing values"] = table["BU"].map(missing_values.groupby("BU").size()).fillna(0).astype(int)
    table["Large variations"] = table["BU"].map(large_variations.groupby("BU").size()).fillna(0).astype(int)
    return table


def build_email_body(summary_table: pd.DataFrame, current_year: int, current_month: str,
                      attachment_name: str | None) -> str:
    """The short PLAIN-TEXT version of the report — used for the console
    print in run_and_send. The actual email uses build_email_html instead,
    so it renders as a real table in Outlook rather than a monospace dump
    (which Outlook doesn't guarantee anyway). `attachment_name` is the
    actual (timestamped) attachment filename, or None if there's nothing to
    attach — never hardcode the name here, it drifts from reality
    otherwise."""
    lines = [f"CSR consolidation report — {current_month} {current_year}", ""]

    total_quality = int(summary_table["Data quality"].sum())
    total_missing = int(summary_table["Missing values"].sum())
    total_variations = int(summary_table["Large variations"].sum())

    if not (total_quality or total_missing or total_variations):
        lines.append("No anomaly found this month.")
        return "\n".join(lines)

    lines.append("Per-BU summary:")
    lines.append(summary_table.to_string(index=False))
    lines.append("")
    lines.append(f"Totals: {total_quality} data quality flag(s), {total_missing} missing value(s), "
                 f"{total_variations} large variation(s) (month-over-month, > "
                 f"{round(config.VARIATION_THRESHOLD * 100)}%).")
    if attachment_name:
        lines.append(f"Full row-by-row detail attached ({attachment_name}).")

    return "\n".join(lines)


_TABLE_STYLE = "border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:13px;"
_TH_STYLE = "border:1px solid #999999;padding:4px 12px;background-color:#1F3864;color:#ffffff;text-align:right;"
_TH_STYLE_FIRST = _TH_STYLE + "text-align:left;"
_TD_STYLE = "border:1px solid #999999;padding:4px 12px;text-align:right;"
_TD_STYLE_FIRST = _TD_STYLE + "text-align:left;"


def _html_table(df: pd.DataFrame) -> str:
    """A small HTML table with styles inline on every cell (rather than a
    <style> block) — Outlook's rendering engine (Word) reliably applies
    inline styles but often strips <style> blocks in the head, so inline is
    the only formatting that survives in every client."""
    header_cells = "".join(
        f'<th style="{_TH_STYLE_FIRST if i == 0 else _TH_STYLE}">{col}</th>'
        for i, col in enumerate(df.columns)
    )
    body_rows = []
    for _, row in df.iterrows():
        cells = "".join(
            f'<td style="{_TD_STYLE_FIRST if i == 0 else _TD_STYLE}">{row[col]}</td>'
            for i, col in enumerate(df.columns)
        )
        body_rows.append(f"<tr>{cells}</tr>")
    return f'<table style="{_TABLE_STYLE}"><tr>{header_cells}</tr>{"".join(body_rows)}</table>'


def build_email_html(summary_table: pd.DataFrame, current_year: int, current_month: str,
                      attachment_name: str | None) -> str:
    """The HTML version of the report sent as the email's actual body: a
    real, styled per-BU table instead of a plain-text dump. The row-by-row
    detail lives in the Excel attachment, not here — see
    write_detail_workbook. `attachment_name` is the actual (timestamped)
    attachment filename, or None if there's nothing to attach."""
    total_quality = int(summary_table["Data quality"].sum())
    total_missing = int(summary_table["Missing values"].sum())
    total_variations = int(summary_table["Large variations"].sum())

    parts = [f'<p style="font-family:Calibri,Arial,sans-serif;font-size:15px;">'
             f'<b>CSR consolidation report — {current_month} {current_year}</b></p>']

    if not (total_quality or total_missing or total_variations):
        parts.append('<p style="font-family:Calibri,Arial,sans-serif;font-size:13px;">'
                      'No anomaly found this month.</p>')
        return "".join(parts)

    parts.append('<p style="font-family:Calibri,Arial,sans-serif;font-size:13px;">Per-BU summary:</p>')
    parts.append(_html_table(summary_table))
    parts.append(
        f'<p style="font-family:Calibri,Arial,sans-serif;font-size:13px;">'
        f'<b>Totals:</b> {total_quality} data quality flag(s), {total_missing} missing value(s), '
        f'{total_variations} large variation(s) (month-over-month, &gt; '
        f'{round(config.VARIATION_THRESHOLD * 100)}%).</p>'
    )
    if attachment_name:
        parts.append(f'<p style="font-family:Calibri,Arial,sans-serif;font-size:13px;">'
                      f'Full row-by-row detail attached ({attachment_name}).</p>')

    return "".join(parts)


def write_detail_workbook(quality_notes: pd.DataFrame, missing_values: pd.DataFrame,
                           large_variations: pd.DataFrame, path) -> None:
    """Writes the full row-by-row detail behind the email's summary counts,
    one sheet per check, so the inbox stays readable while the detail is
    still one click away."""
    with pd.ExcelWriter(path) as writer:
        quality_notes.to_excel(writer, sheet_name="Data quality", index=False)
        missing_values.to_excel(writer, sheet_name="Missing values", index=False)
        large_variations.to_excel(writer, sheet_name="Large variations", index=False)


def send_report_email(subject: str, html_body: str, attachment_path=None) -> bool:
    """Sends `html_body` (plus `attachment_path`, if given) to
    config.REPORT_RECIPIENTS via the local Outlook desktop app, as an HTML
    email (real tables, not a monospace text dump). Returns False without
    sending anything if config.SEND_ERROR_REPORT_EMAIL is False, there are
    no recipients, or Outlook isn't already running (see below). Never
    raises: a failure to send is printed as a warning, never a crash of the
    consolidation run — the report itself is always printed to the console
    regardless.

    Deliberately uses GetActiveObject, NOT Dispatch: GetActiveObject only
    attaches to an Outlook that is ALREADY running, and fails fast if it
    isn't. Dispatch() would instead try to LAUNCH Outlook when it's not
    running — and doing that from a non-interactive process (a scheduled
    task, or this script run from an automation/background context) can
    hang indefinitely with no visible error, waiting on a UI Outlook can't
    actually show. Confirmed the hard way: a run left Outlook-less hung for
    over an hour with ~0% CPU use, never timing out on its own. Simplest
    fix is to never attempt the launch — just require Outlook to already be
    open."""
    if not config.SEND_ERROR_REPORT_EMAIL:
        print("Anomaly report email disabled (config.SEND_ERROR_REPORT_EMAIL = False).")
        return False
    if not config.REPORT_RECIPIENTS:
        print("No report recipients configured (config.REPORT_RECIPIENTS is empty) — email not sent.")
        return False

    try:
        import win32com.client
        outlook = win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        print("Outlook doesn't appear to be running — open Outlook and re-run to get the "
              "anomaly report by email (it was still printed above).")
        return False

    try:
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = "; ".join(config.REPORT_RECIPIENTS)
        mail.Subject = subject
        mail.HTMLBody = html_body
        if attachment_path is not None:
            mail.Attachments.Add(str(attachment_path))
        mail.Send()
        print(f"Anomaly report emailed to {', '.join(config.REPORT_RECIPIENTS)}.")
        return True
    except Exception as e:
        print(f"Could not send the anomaly report email: {e}")
        return False


def run_and_send(raw: pd.DataFrame, consolidated: pd.DataFrame, reference: pd.DataFrame,
                  current_year: int, current_month: str) -> str:
    """Builds the anomaly report from this run's data, prints the plain-text
    version to the console, writes the detail workbook (only if there's
    something to show), and emails the HTML version (a real table, see
    build_email_html) — the single entry point csr_calc_engine.py calls
    after computing the consolidated result. Data quality notes and large
    variations are reported across every year present in `raw`/`consolidated`
    (a data quality flag from a prior year is still worth a look); missing
    values are restricted to `current_year` (see find_missing_values — a
    finished prior year is never flagged as still owing data). Returns the
    plain-text body."""
    input_ids = set(reference[reference.Kind == "input"]["ID"])
    quality_notes = find_data_quality_notes(raw)
    missing_values = find_missing_values(raw, input_ids, current_year, current_month)
    large_variations = find_large_variations(consolidated)

    bus = sorted(raw["BU"].unique())
    summary_table = build_summary_table(quality_notes, missing_values, large_variations, bus)

    has_anomalies = bool(len(quality_notes) or len(missing_values) or len(large_variations))
    attachment_path = None
    if has_anomalies:
        # Timestamped (not overwritten every run) so past reports stay
        # around as a record of what was flagged and when, rather than only
        # ever showing the latest run's snapshot.
        attachment_path = config.OUTPUT_DIR / f"Anomaly_report_CSR_{date.today().isoformat()}.xlsx"
        write_detail_workbook(quality_notes, missing_values, large_variations, attachment_path)

    attachment_name = attachment_path.name if attachment_path else None

    body = build_email_body(summary_table, current_year, current_month, attachment_name)
    print("\n" + body)

    html_body = build_email_html(summary_table, current_year, current_month, attachment_name)
    send_report_email(f"CSR consolidation report — {current_month} {current_year}", html_body, attachment_path)
    return body
