"""Tests for scripts/integrate_data_entry.py.

read_data_entry_file() relies on the same column layout as
generate_data_entry_file.write_month_sheet() (via DISPLAY_COLS, imported
directly from that module) — the round-trip test below builds the file with
the real write_month_sheet rather than a hand-rolled workbook, so any drift
between the two scripts breaks this test instead of production re-integration.
"""
import openpyxl
import pandas as pd
import pytest

import config
import generate_data_entry_file as gen
import integrate_data_entry as integ


def test_none_stays_none_without_note():
    assert integ.parse_raw_value(None) == (None, None)


def test_empty_string_is_treated_as_blank_not_a_quality_issue():
    # Unlike extract_reference_and_data.parse_raw_value: a cell just cleared
    # by the referent (whitespace) is not an anomaly to flag.
    assert integ.parse_raw_value("   ") == (None, None)


def test_number_passes_through_unchanged():
    assert integ.parse_raw_value(56.7) == (56.7, None)


def test_french_decimal_comma_is_converted_and_flagged():
    value, note = integ.parse_raw_value("209,3")
    assert value == pytest.approx(209.3)
    assert "decimal comma" in note


def test_unparsable_text_is_flagged():
    value, note = integ.parse_raw_value("28,9kWh")
    assert value is None
    assert "NOT PARSABLE" in note


def _sample_record():
    return {
        "ID": "Wat.1", "Topic": "Water", "KPI": "Water consumption", "Unit": "m3",
        "Definition": "definition", "Calculation detail": "calc detail", "Source of data": "source",
    }


def test_read_data_entry_file_matches_write_month_sheet_column_layout(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DATA_ENTRY_DIR", tmp_path)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    gen.write_month_sheet(wb, "January", [_sample_record()],
                           {("January", "Wat.1"): (123.4, "hello")}, is_current=True)
    wb.save(tmp_path / "BU1_data_entry_CSR.xlsm")

    df = integ.read_data_entry_file("BU1")

    row = df[(df.Month == "January") & (df.ID == "Wat.1")].iloc[0]
    assert row["Value"] == pytest.approx(123.4)
    assert row["Comment"] == "hello"
    assert pd.isna(row["Data quality note"])


def test_read_data_entry_file_raises_clear_error_when_file_missing(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DATA_ENTRY_DIR", tmp_path)

    with pytest.raises(FileNotFoundError):
        integ.read_data_entry_file("UnknownBU")
