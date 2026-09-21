"""Tests for scripts/generate_data_entry_file.py.

The most important test here is the write/read "round-trip" for a monthly
tab (test_write_then_load_existing_values_roundtrip): it's exactly the kind
of regression we hit when adding the "Calculation detail" column — the
"Value of the month"/"Comment" columns had shifted without
load_existing_values() being updated accordingly.
"""
import openpyxl
import pandas as pd

import config
import generate_data_entry_file as gen


def test_is_filled():
    assert gen._is_filled(None) is False
    assert gen._is_filled(float("nan")) is False
    assert gen._is_filled(0) is True
    assert gen._is_filled(5.5) is True
    assert gen._is_filled("some text") is True


def test_compute_completion_counts_filled_indicators_per_month():
    entry_records = [{"ID": "A"}, {"ID": "B"}]
    existing = {
        ("January", "A"): (10, None),
        ("January", "B"): (None, None),  # present but empty -> not counted
    }

    completion = gen.compute_completion(entry_records, existing)

    assert len(completion) == 12
    assert ("January", 1, 2) in completion
    assert ("February", 0, 2) in completion  # no data for this month


def test_load_data_entry_reference_keeps_only_input_kind_owned_by_csr_referent(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "INPUT_DIR", tmp_path)
    reference = pd.DataFrame([
        {"ID": "A", "Kind": "input", "Responsible": "CSR referent"},
        {"ID": "B", "Kind": "calculated", "Responsible": "CSR referent"},  # excluded: calculated
        {"ID": "C", "Kind": "input", "Responsible": "CSO"},                # excluded: not the referent
        {"ID": "D", "Kind": "input", "Responsible": "CSR referent"},
    ])
    with pd.ExcelWriter(tmp_path / "Indicator_reference_CSR.xlsx") as writer:
        reference.to_excel(writer, sheet_name="Reference", index=False)

    result = gen.load_data_entry_reference()

    assert list(result["ID"]) == ["A", "D"]


def test_context_col_widths_covers_every_display_col():
    # If a column is added/removed from DISPLAY_COLS without updating
    # CONTEXT_COL_WIDTHS, write_month_sheet raises a KeyError at runtime —
    # this test catches it beforehand.
    assert set(gen.CONTEXT_COL_WIDTHS) == set(gen.DISPLAY_COLS)


def _sample_record(id_="Wat.1"):
    return {
        "ID": id_, "Topic": "Water", "KPI": "Water consumption", "Unit": "m3",
        "Definition": "definition", "Calculation detail": "calc detail",
        "Source of data": "source",
    }


def test_write_then_load_existing_values_roundtrip(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DATA_ENTRY_DIR", tmp_path)

    existing_at_write_time = {("January", "Wat.1"): (123.4, "hello")}
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    gen.write_month_sheet(wb, "January", [_sample_record()], existing_at_write_time, is_current=True)
    wb.save(tmp_path / "BU1_data_entry_CSR.xlsm")

    empty_consolidated = pd.DataFrame(columns=["BU", "Month", "ID", "Value", "Comment", "Data quality note"])
    reloaded = gen.load_existing_values("BU1", empty_consolidated)

    assert reloaded[("January", "Wat.1")] == (123.4, "hello")


def test_load_existing_values_prefers_data_entry_file_over_consolidated(tmp_path, monkeypatch):
    # In case of a conflict between the already-integrated history and the
    # current data entry file (not yet integrated), the latter must win
    # (it's the most recent value) — see the function's docstring.
    monkeypatch.setattr(config, "DATA_ENTRY_DIR", tmp_path)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    gen.write_month_sheet(wb, "January", [_sample_record()], {("January", "Wat.1"): (999, "new")},
                           is_current=True)
    wb.save(tmp_path / "BU1_data_entry_CSR.xlsm")

    consolidated = pd.DataFrame([
        {"BU": "BU1", "Month": "January", "ID": "Wat.1", "Value": 1, "Comment": "old",
         "Data quality note": None},
    ])
    reloaded = gen.load_existing_values("BU1", consolidated)

    assert reloaded[("January", "Wat.1")] == (999, "new")
