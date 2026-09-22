"""Tests for scripts/consolidation_report.py.

send_report_email's actual Outlook-sending path (config.SEND_ERROR_REPORT_EMAIL
= True with recipients configured) is deliberately NOT exercised here: on
Windows it would send a real email every time the test suite runs, and on
CI (Linux) the win32com import fails anyway — neither is a meaningful test.
Only the "disabled" and "no recipients" short-circuits, which never touch
win32com, are tested.
"""
import pandas as pd
import pytest

import config
import consolidation_report as report


def test_find_data_quality_notes_keeps_only_flagged_rows():
    raw = pd.DataFrame([
        {"BU": "BAF", "Year": 2026, "Month": "January", "ID": "A", "Value": 1.0, "Comment": None,
         "Data quality note": None},
        {"BU": "BAF", "Year": 2026, "Month": "January", "ID": "B", "Value": 2.0, "Comment": None,
         "Data quality note": "Corrected: decimal comma"},
    ])
    result = report.find_data_quality_notes(raw)
    assert list(result["ID"]) == ["B"]


def test_find_data_quality_notes_covers_every_year():
    # Unlike find_missing_values, a data quality flag from a prior year is
    # still worth a look — no Year filter here.
    raw = pd.DataFrame([
        {"BU": "BAF", "Year": 2025, "Month": "January", "ID": "A", "Value": 1.0, "Comment": None,
         "Data quality note": "Corrected: decimal comma"},
    ])
    result = report.find_data_quality_notes(raw)
    assert list(result["ID"]) == ["A"]


def test_find_missing_values_flags_absent_row_and_empty_value_up_to_current_month():
    raw = pd.DataFrame([
        {"BU": "BU1", "Year": 2026, "Month": "January", "ID": "A", "Value": 10.0},
        {"BU": "BU1", "Year": 2026, "Month": "January", "ID": "B", "Value": None},     # empty -> missing
        {"BU": "BU1", "Year": 2026, "Month": "February", "ID": "B", "Value": 5.0},
        # BU1/February/A: no row at all -> missing
        {"BU": "BU1", "Year": 2026, "Month": "March", "ID": "A", "Value": None},       # not due yet, ignored
    ])
    missing = report.find_missing_values(raw, {"A", "B"}, current_year=2026, current_month="February")

    got = set(zip(missing.BU, missing.Month.astype(str), missing.ID))
    assert got == {("BU1", "January", "B"), ("BU1", "February", "A")}


def test_find_missing_values_ignores_other_years():
    # A finished prior year must never show up as still owing data, and a
    # different year's row must not satisfy the current year's requirement.
    raw = pd.DataFrame([
        {"BU": "BU1", "Year": 2025, "Month": "January", "ID": "A", "Value": 10.0},
    ])
    missing = report.find_missing_values(raw, {"A"}, current_year=2026, current_month="January")

    assert missing.empty  # no 2026 rows at all -> no BU to report on, not "everything missing"


def test_find_missing_values_empty_when_nothing_due():
    raw = pd.DataFrame(columns=["BU", "Year", "Month", "ID", "Value"])
    result = report.find_missing_values(raw, set(), current_year=2026, current_month="January")
    assert result.empty


def test_find_large_variations_flags_jump_above_threshold_only():
    consolidated = pd.DataFrame([
        {"BU": "BU1", "Year": 2026, "ID": "X", "Month": "January", "Value": 100.0},
        {"BU": "BU1", "Year": 2026, "ID": "X", "Month": "February", "Value": 150.0},  # +50% -> flagged
        {"BU": "BU1", "Year": 2026, "ID": "X", "Month": "March", "Value": 155.0},     # +3.3% -> not flagged
    ])
    consolidated["Month"] = pd.Categorical(consolidated["Month"], categories=report.MONTH_ORDER, ordered=True)

    result = report.find_large_variations(consolidated)

    assert len(result) == 1
    row = result.iloc[0]
    assert (row["BU"], row["ID"], row["From"], row["To"]) == ("BU1", "X", "January", "February")


def test_find_large_variations_ignores_missing_values_and_zero_baseline():
    consolidated = pd.DataFrame([
        {"BU": "BU1", "Year": 2026, "ID": "X", "Month": "January", "Value": 0.0},
        {"BU": "BU1", "Year": 2026, "ID": "X", "Month": "February", "Value": 40.0},   # baseline was 0 -> skipped
        {"BU": "BU1", "Year": 2026, "ID": "X", "Month": "March", "Value": None},      # missing -> doesn't become the new baseline
        {"BU": "BU1", "Year": 2026, "ID": "X", "Month": "April", "Value": 44.0},      # +10% vs February (40) -> not flagged
    ])
    consolidated["Month"] = pd.Categorical(consolidated["Month"], categories=report.MONTH_ORDER, ordered=True)

    result = report.find_large_variations(consolidated)

    assert result.empty


def test_find_large_variations_never_compares_across_a_year_boundary():
    # December of one year followed by January of the next must never be
    # treated as a month-over-month step — a year boundary is exactly where
    # a real jump (annual reset, new pricing) is expected.
    consolidated = pd.DataFrame([
        {"BU": "BU1", "Year": 2025, "ID": "X", "Month": "December", "Value": 1000.0},
        {"BU": "BU1", "Year": 2026, "ID": "X", "Month": "January", "Value": 1.0},
    ])
    consolidated["Month"] = pd.Categorical(consolidated["Month"], categories=report.MONTH_ORDER, ordered=True)

    result = report.find_large_variations(consolidated)

    assert result.empty


def test_build_summary_table_counts_per_bu_including_zeroes():
    quality_notes = pd.DataFrame([{"BU": "BAF", "Month": "January", "ID": "A", "Value": 1.0,
                                    "Data quality note": "Corrected: decimal comma"}])
    missing = pd.DataFrame([{"BU": "BAF", "Month": "January", "ID": "B"},
                             {"BU": "BAF", "Month": "February", "ID": "C"}])
    variations = pd.DataFrame(columns=["BU", "ID", "From", "From value", "To", "To value", "Variation"])

    table = report.build_summary_table(quality_notes, missing, variations, bus=["BAF", "BFR"])

    baf = table[table.BU == "BAF"].iloc[0]
    assert (baf["Data quality"], baf["Missing values"], baf["Large variations"]) == (1, 2, 0)
    bfr = table[table.BU == "BFR"].iloc[0]
    assert (bfr["Data quality"], bfr["Missing values"], bfr["Large variations"]) == (0, 0, 0)


def test_build_email_body_reports_no_anomaly_when_everything_is_empty():
    table = report.build_summary_table(pd.DataFrame(columns=["BU"]), pd.DataFrame(columns=["BU"]),
                                        pd.DataFrame(columns=["BU"]), bus=["BAF"])
    body = report.build_email_body(table, current_year=2026, current_month="January", attachment_name=None)
    assert "No anomaly found this month." in body
    assert "January" in body


def test_build_email_body_includes_totals_and_attachment_mention():
    quality_notes = pd.DataFrame([{"BU": "BAF", "Data quality note": "x"}])
    missing = pd.DataFrame(columns=["BU"])
    variations = pd.DataFrame(columns=["BU"])
    table = report.build_summary_table(quality_notes, missing, variations, bus=["BAF"])

    body = report.build_email_body(table, current_year=2026, current_month="February",
                                    attachment_name="Anomaly_report_CSR_2026-09-21.xlsx")

    assert "Per-BU summary:" in body
    assert "Totals: 1 data quality flag(s), 0 missing value(s), 0 large variation(s)" in body
    assert "Anomaly_report_CSR_2026-09-21.xlsx" in body
    assert "No anomaly found" not in body


def test_build_email_html_reports_no_anomaly_when_everything_is_empty():
    table = report.build_summary_table(pd.DataFrame(columns=["BU"]), pd.DataFrame(columns=["BU"]),
                                        pd.DataFrame(columns=["BU"]), bus=["BAF"])
    html = report.build_email_html(table, current_year=2026, current_month="January", attachment_name=None)
    assert "No anomaly found this month." in html
    assert "<table" not in html


def test_build_email_html_renders_a_real_table_with_bu_rows():
    quality_notes = pd.DataFrame([{"BU": "BAF", "Data quality note": "x"}])
    missing = pd.DataFrame(columns=["BU"])
    variations = pd.DataFrame(columns=["BU"])
    table = report.build_summary_table(quality_notes, missing, variations, bus=["BAF", "BFR"])

    html = report.build_email_html(table, current_year=2026, current_month="February",
                                    attachment_name="Anomaly_report_CSR_2026-09-21.xlsx")

    assert "<table" in html and "</table>" in html
    assert html.count("<tr>") == 3  # header + BAF + BFR
    assert "<th" in html and "<td" in html
    assert "BAF" in html and "BFR" in html
    assert "Anomaly_report_CSR_2026-09-21.xlsx" in html


def test_write_detail_workbook_roundtrips_all_three_sheets(tmp_path):
    quality_notes = pd.DataFrame([{"BU": "BAF", "Month": "January", "ID": "A", "Value": 1.0,
                                    "Data quality note": "Corrected: decimal comma"}])
    missing = pd.DataFrame([{"BU": "BAF", "Month": "January", "ID": "B"}])
    variations = pd.DataFrame([{"BU": "BAF", "ID": "C", "From": "January", "From value": 1.0,
                                 "To": "February", "To value": 2.0, "Variation": "100%"}])
    path = tmp_path / "Anomaly_report_CSR.xlsx"

    report.write_detail_workbook(quality_notes, missing, variations, path)

    sheets = pd.read_excel(path, sheet_name=None)
    assert set(sheets) == {"Data quality", "Missing values", "Large variations"}
    assert list(sheets["Data quality"]["ID"]) == ["A"]
    assert list(sheets["Missing values"]["ID"]) == ["B"]
    assert list(sheets["Large variations"]["ID"]) == ["C"]


def test_send_report_email_disabled_by_config_does_nothing(monkeypatch):
    monkeypatch.setattr(config, "SEND_ERROR_REPORT_EMAIL", False)
    assert report.send_report_email("subject", "body") is False


def test_send_report_email_skips_when_no_recipients(monkeypatch):
    monkeypatch.setattr(config, "SEND_ERROR_REPORT_EMAIL", True)
    monkeypatch.setattr(config, "REPORT_RECIPIENTS", [])
    assert report.send_report_email("subject", "body") is False
