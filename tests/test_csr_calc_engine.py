"""Tests for the calculation engine (scripts/csr_calc_engine.py): the only
real business logic executed by the pipeline (the "calculated" indicators'
formulas). Everything here is testable with no Excel file at all:
compute_bu_month takes a plain dict, run() plain pandas DataFrames."""
import openpyxl
import pandas as pd
import pytest

import config
import csr_calc_engine as engine


def test_z_treats_none_and_nan_as_zero():
    assert engine.z(None) == 0
    assert engine.z(float("nan")) == 0
    assert engine.z(5) == 5
    assert engine.z(0) == 0


@pytest.mark.parametrize("kpi_id,values,expected", [
    ("Ref.1", {f"Ref.{i}": i for i in range(2, 10)}, sum(range(2, 10))),
    # A detail row left blank (None, not absent from the dict) counts as 0
    # (like SUM() in Excel), not an error.
    ("Ref.1", {"Ref.2": 5, **{f"Ref.{i}": None for i in range(3, 10)}}, 5),
])
def test_formula_matches_expected_value(kpi_id, values, expected):
    v = lambda i: values[i]
    assert engine.FORMULAS[kpi_id](v) == pytest.approx(expected)


def test_formulas_only_contains_additive_indicators():
    # 23/09/2026 decision: no ratios/divisions here (can't be summed to a
    # group total), and no Env.1 (never available to this pipeline). If this
    # ever fails, someone added a division back — see the module docstring.
    assert set(engine.FORMULAS) == {"Ene.9", "Ref.1"}


def test_ene9_combines_direct_and_weighted_sources():
    values = {
        "Ene.1": 10, "Ene.2": 20, "Ene.3": 0, "Ene.4": 0, "Ene.5": 0,
        "Ene.6": 2, "Ene.6.1": 3,   # 2*3 = 6
        "Ene.7": 4, "Ene.7.1": 5,   # 4*5 = 20
        "Ene.8": 1,                 # 1*5 (Ene.7.1) = 5
    }
    v = lambda i: values[i]
    assert engine.FORMULAS["Ene.9"](v) == pytest.approx(10 + 20 + 6 + 20 + 5)


def test_compute_bu_month_computes_both_sums_independently():
    raw = {
        "Ene.1": 100, "Ene.2": 10, "Ene.3": 10, "Ene.4": 10, "Ene.5": 0,
        "Ene.6": 0, "Ene.6.1": 0, "Ene.7": 0, "Ene.7.1": 0, "Ene.8": 0,
        "Ref.2": 1, "Ref.3": 1, "Ref.4": 0, "Ref.5": 0, "Ref.6": 0,
        "Ref.7": 0, "Ref.8": 0, "Ref.9": 0,
    }
    result = engine.compute_bu_month(raw)

    assert result["Ene.9"] == pytest.approx(130)
    assert result["Ref.1"] == pytest.approx(2)
    # Original entries stay unchanged in the result.
    assert result["Ene.1"] == 100


def test_compute_bu_month_leaves_indicator_absent_when_input_missing():
    # Ene.9 needs Ene.1..Ene.8: missing entirely (not just empty) makes the
    # whole sum fail rather than silently treat it as 0 — must stay absent,
    # not crash, not put in a bogus value.
    raw = {"Ref.2": 5}
    result = engine.compute_bu_month(raw)
    assert "Ene.9" not in result
    assert result["Ref.2"] == 5


def _ref_rows(bu, year, month, ref2_value):
    """A full Ref.2..Ref.9 set of raw rows (Ref.1 needs every one of them
    present, even at 0, to compute at all — see compute_bu_month: a
    reference entirely missing from raw_values fails the whole sum, it's
    not the same as an empty/0 value)."""
    return [{"BU": bu, "Year": year, "Month": month, "ID": f"Ref.{i}",
              "Value": ref2_value if i == 2 else 0.0} for i in range(2, 10)]


def test_run_merges_computed_and_input_values_with_reference_metadata():
    reference = pd.DataFrame([
        {"ID": "Ref.2", "Topic": "Refrigerants", "KPI": "R449A leak", "Unit": "Kg", "Kind": "input"},
        {"ID": "Ref.1", "Topic": "Refrigerants", "KPI": "Total refrigerants leaks", "Unit": "Kg",
         "Kind": "calculated"},
    ])
    raw = pd.DataFrame(_ref_rows("BAF", 2026, "January", 3.0))

    out = engine.run(reference, raw)

    ref1 = out[out.ID == "Ref.1"].iloc[0]
    assert ref1["Value"] == pytest.approx(3.0)
    assert ref1["Kind"] == "calculated"
    assert ref1["Topic"] == "Refrigerants"
    assert ref1["Year"] == 2026
    assert set(out.ID) == {"Ref.2", "Ref.1"}


def test_run_keeps_two_years_separate():
    # Same BU/month/ID in two different years must not be summed or merged
    # into one row — a year boundary is a hard partition, not just another
    # grouping key.
    reference = pd.DataFrame([
        {"ID": "Ref.2", "Topic": "Refrigerants", "KPI": "R449A leak", "Unit": "Kg", "Kind": "input"},
        {"ID": "Ref.1", "Topic": "Refrigerants", "KPI": "Total refrigerants leaks", "Unit": "Kg",
         "Kind": "calculated"},
    ])
    raw = pd.DataFrame(_ref_rows("BAF", 2026, "January", 3.0) + _ref_rows("BAF", 2027, "January", 7.0))

    out = engine.run(reference, raw)

    ref1_by_year = out[out.ID == "Ref.1"].set_index("Year")["Value"]
    assert ref1_by_year[2026] == pytest.approx(3.0)
    assert ref1_by_year[2027] == pytest.approx(7.0)


def test_run_ignores_computed_id_absent_from_reference():
    # Ref.1 is computable (every Ref.2..Ref.9 is entered, even if 0) but
    # deliberately absent from the reference list (indicator removed):
    # run() must not crash, just silently skip it.
    reference = pd.DataFrame([
        {"ID": "Ref.2", "Topic": "Refrigerants", "KPI": "R449A leak", "Unit": "Kg", "Kind": "input"},
    ])
    raw = pd.DataFrame(
        [{"BU": "BAF", "Year": 2026, "Month": "January", "ID": f"Ref.{i}",
          "Value": 3.0 if i == 2 else 0.0} for i in range(2, 10)]
    )
    out = engine.run(reference, raw)
    assert list(out.ID) == ["Ref.2"]


# ---------------------------------------------------------------------------
# build_completion_table — the per-BU/per-month % completion matrix, scoped
# to CSR-referent-owned indicators only (never Safety/Finance/HR-owned ones,
# even though those are also Kind == "input"), and to current_year only.
# ---------------------------------------------------------------------------

def test_build_completion_table_only_counts_csr_referent_owned_indicators():
    reference = pd.DataFrame([
        {"ID": "Wat.1", "Kind": "input", "Responsible": "CSR referent"},
        {"ID": "Ene.1", "Kind": "input", "Responsible": "CSR referent"},
        {"ID": "Fin.1", "Kind": "input", "Responsible": "Finance - Nikola Terzic"},  # excluded
        {"ID": "Saf.1", "Kind": "input", "Responsible": "H&S manager"},              # excluded
        {"ID": "Wat.2", "Kind": "calculated", "Responsible": "CSR referent"},        # excluded: calculated
    ])
    raw = pd.DataFrame([
        {"BU": "BAF", "Year": 2026, "Month": "January", "ID": "Wat.1", "Value": 100.0},
        # Ene.1 not filled for BAF/January -> should count as missing (1/2 = 50%).
        {"BU": "BAF", "Year": 2026, "Month": "January", "ID": "Fin.1", "Value": 999.0},  # Finance-owned, ignored
        {"BU": "BAF", "Year": 2026, "Month": "January", "ID": "Saf.1", "Value": 999.0},  # H&S-owned, ignored
    ])

    table = engine.build_completion_table(raw, reference, current_year=2026)

    jan = table[table.BU == "BAF"].iloc[0]
    assert jan["January"] == "50%"


def test_build_completion_table_covers_all_12_months_future_included():
    reference = pd.DataFrame([{"ID": "Wat.1", "Kind": "input", "Responsible": "CSR referent"}])
    raw = pd.DataFrame([{"BU": "BAF", "Year": 2026, "Month": "January", "ID": "Wat.1", "Value": 100.0}])

    table = engine.build_completion_table(raw, reference, current_year=2026)

    row = table[table.BU == "BAF"].iloc[0]
    assert row["January"] == "100%"
    assert row["December"] == "0%"  # never due, never filled — still shown, not omitted
    assert set(engine.MONTH_ORDER).issubset(set(table.columns))


def test_build_completion_table_has_one_row_per_bu_present_in_raw_data():
    reference = pd.DataFrame([{"ID": "Wat.1", "Kind": "input", "Responsible": "CSR referent"}])
    raw = pd.DataFrame([
        {"BU": "BAF", "Year": 2026, "Month": "January", "ID": "Wat.1", "Value": 100.0},
        {"BU": "BFR", "Year": 2026, "Month": "January", "ID": "Wat.1", "Value": None},
    ])

    table = engine.build_completion_table(raw, reference, current_year=2026)

    assert set(table.BU) == {"BAF", "BFR"}
    assert table[table.BU == "BFR"].iloc[0]["January"] == "0%"


def test_build_completion_table_ignores_other_years():
    # A BU that only has 2025 data shouldn't pollute the 2026 completion
    # matrix, and a 2025-only value must not count as "filled" for 2026.
    reference = pd.DataFrame([{"ID": "Wat.1", "Kind": "input", "Responsible": "CSR referent"}])
    raw = pd.DataFrame([
        {"BU": "BAF", "Year": 2025, "Month": "January", "ID": "Wat.1", "Value": 100.0},
        {"BU": "BFR", "Year": 2026, "Month": "January", "ID": "Wat.1", "Value": 100.0},
    ])

    table = engine.build_completion_table(raw, reference, current_year=2026)

    assert set(table.BU) == {"BFR"}


# ---------------------------------------------------------------------------
# validate_against_originals — the one-off check against the original raw
# Excel files. Unlike the rest of the engine, this one does real file I/O
# and only communicates through print(), so the tests below build a tiny
# fake "raw file" per case and read back stdout (capsys) rather than a
# return value.
# ---------------------------------------------------------------------------

def _write_fixture_workbook(path, sheets):
    """sheets: {month_name: {indicator_id: value}}. "February" must always
    be one of the keys: validate_against_originals unconditionally reads
    that sheet's column B to build its ID -> row lookup, whichever month is
    actually being compared."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for month, id_to_value in sheets.items():
        ws = wb.create_sheet(month)
        for row, (kpi_id, value) in enumerate(id_to_value.items(), start=2):
            ws.cell(row=row, column=2, value=kpi_id)
            ws.cell(row=row, column=10, value=value)
    wb.save(path)


@pytest.fixture
def single_bu(monkeypatch, tmp_path):
    """Restricts BU_LIST/RAW_FILES to one fake BU pointing at tmp_path, so a
    test only has to build one small workbook instead of the real six."""
    monkeypatch.setattr(config, "RAW_FILES", {"TestBU": "testbu.xlsx"})
    monkeypatch.setattr(config, "RAW_DATA_DIR", tmp_path)
    monkeypatch.setattr(engine, "BU_LIST", ["TestBU"])
    return tmp_path / "testbu.xlsx"


def test_validate_skips_silently_when_raw_files_are_absent(capsys, single_bu):
    # single_bu points at a file that is never written: raw_data/ is empty,
    # which is the normal case for a routine monthly run.
    engine.validate_against_originals(pd.DataFrame())
    out = capsys.readouterr().out
    assert "Validation skipped" in out


def test_validate_reports_no_mismatch_when_recalculation_agrees(capsys, single_bu):
    _write_fixture_workbook(single_bu, {"February": {"Ene.9": 2.0}})
    consolidated = pd.DataFrame([
        {"BU": "TestBU", "Year": engine.HISTORICAL_VALIDATION_YEAR, "Month": "February", "ID": "Ene.9", "Value": 2.0},
    ])
    engine.validate_against_originals(consolidated)
    out = capsys.readouterr().out
    assert "1 values compared" in out
    assert "0 UNEXPLAINED discrepancy" in out


def test_validate_flags_an_undocumented_discrepancy(capsys, single_bu):
    _write_fixture_workbook(single_bu, {"February": {"Ene.9": 99.0}})
    consolidated = pd.DataFrame([
        {"BU": "TestBU", "Year": engine.HISTORICAL_VALIDATION_YEAR, "Month": "February", "ID": "Ene.9", "Value": 2.0},
    ])
    engine.validate_against_originals(consolidated)
    out = capsys.readouterr().out
    assert "1 UNEXPLAINED discrepancy" in out
    assert "MISMATCH" in out


def test_validate_reports_a_known_correction_separately_from_real_mismatches(
    capsys, single_bu, monkeypatch,
):
    # Same disagreement as the test above, but this time it is a documented,
    # expected correction (KNOWN_CORRECTIONS) — it must be counted on its
    # own line, not raised as an unexplained mismatch.
    monkeypatch.setitem(
        engine.KNOWN_CORRECTIONS, ("TestBU", "February", "Ene.9"),
        "test fixture: documented discrepancy",
    )
    _write_fixture_workbook(single_bu, {"February": {"Ene.9": 99.0}})
    consolidated = pd.DataFrame([
        {"BU": "TestBU", "Year": engine.HISTORICAL_VALIDATION_YEAR, "Month": "February", "ID": "Ene.9", "Value": 2.0},
    ])
    engine.validate_against_originals(consolidated)
    out = capsys.readouterr().out
    assert "1 expected correction" in out
    assert "0 UNEXPLAINED discrepancy" in out


def test_validate_never_compares_the_january_sheet(capsys, single_bu):
    # January never had the calculated rows in the original files: even if
    # a January sheet exists with a wildly different value, it must be
    # skipped rather than reported as a mismatch.
    _write_fixture_workbook(single_bu, {
        "February": {"Ene.9": 2.0},
        "January": {"Ene.9": 999.0},
    })
    consolidated = pd.DataFrame([
        {"BU": "TestBU", "Year": engine.HISTORICAL_VALIDATION_YEAR, "Month": "February", "ID": "Ene.9", "Value": 2.0},
        {"BU": "TestBU", "Year": engine.HISTORICAL_VALIDATION_YEAR, "Month": "January", "ID": "Ene.9", "Value": 2.0},
    ])
    engine.validate_against_originals(consolidated)
    out = capsys.readouterr().out
    assert "1 values compared" in out
    assert "0 UNEXPLAINED discrepancy" in out


def test_validate_skips_a_non_numeric_original_value(capsys, single_bu):
    # A cell still showing an Excel error (#REF!, #DIV/0!) is read by
    # openpyxl as a plain string: nothing sensible to compare it against.
    _write_fixture_workbook(single_bu, {"February": {"Ene.9": "#REF!"}})
    consolidated = pd.DataFrame([
        {"BU": "TestBU", "Year": engine.HISTORICAL_VALIDATION_YEAR, "Month": "February", "ID": "Ene.9", "Value": 130.0},
    ])
    engine.validate_against_originals(consolidated)
    out = capsys.readouterr().out
    assert "0 values compared" in out


def test_validate_skips_when_no_recalculated_value_exists_for_that_id_month(capsys, single_bu):
    # The original file has the ID/month, but the recalculation has nothing
    # for it (e.g. missing raw input): skip rather than crash on an empty
    # mine_series.
    _write_fixture_workbook(single_bu, {"February": {"Ene.9": 2.0}})
    consolidated = pd.DataFrame(columns=["BU", "Year", "Month", "ID", "Value"])
    engine.validate_against_originals(consolidated)
    out = capsys.readouterr().out
    assert "0 values compared" in out