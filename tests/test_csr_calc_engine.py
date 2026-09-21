"""Tests for the calculation engine (scripts/csr_calc_engine.py): the only
real business logic executed by the pipeline (the "calculated" indicators'
formulas). Everything here is testable with no Excel file at all:
compute_bu_month takes a plain dict, run() plain pandas DataFrames."""
import pandas as pd
import pytest

import csr_calc_engine as engine


def test_z_treats_none_and_nan_as_zero():
    assert engine.z(None) == 0
    assert engine.z(float("nan")) == 0
    assert engine.z(5) == 5
    assert engine.z(0) == 0


def test_hours_worked_uses_saf42_when_present():
    v = lambda i: {"Saf.4.2": 500}[i]
    assert engine._hours_worked(v) == 500


def test_hours_worked_falls_back_to_default_formula_when_saf42_missing_value():
    # Reproduces =IF(Saf.4.2="", Env.1*Saf.4.1*8, Saf.4.2): Saf.4.2 empty (None).
    d = {"Saf.4.2": None, "Env.1": 10, "Saf.4.1": 2}
    v = lambda i: d[i]
    assert engine._hours_worked(v) == 10 * 2 * 8


@pytest.mark.parametrize("kpi_id,values,expected", [
    ("Wat.2", {"Wat.1": 100, "Env.1": 50}, 2.0),
    ("Was.3", {"Was.2": 20, "Was.1": 200}, 0.1),
    ("Was.4", {"Was.1": 30, "Env.1": 60}, 0.5),
    ("Ref.1", {f"Ref.{i}": i for i in range(2, 10)}, sum(range(2, 10))),
    # A detail row left blank (None, not absent from the dict) counts as 0
    # (like SUM() in Excel), not an error.
    ("Ref.1", {"Ref.2": 5, **{f"Ref.{i}": None for i in range(3, 10)}}, 5),
])
def test_formula_matches_expected_value(kpi_id, values, expected):
    v = lambda i: values[i]
    assert engine.FORMULAS[kpi_id](v) == pytest.approx(expected)


def test_ene9_combines_direct_and_weighted_sources():
    values = {
        "Ene.1": 10, "Ene.2": 20, "Ene.3": 0, "Ene.4": 0, "Ene.5": 0,
        "Ene.6": 2, "Ene.6.1": 3,   # 2*3 = 6
        "Ene.7": 4, "Ene.7.1": 5,   # 4*5 = 20
        "Ene.8": 1,                 # 1*5 (Ene.7.1) = 5
    }
    v = lambda i: values[i]
    assert engine.FORMULAS["Ene.9"](v) == pytest.approx(10 + 20 + 6 + 20 + 5)


def test_saf6_and_saf7_use_hours_worked_denominator():
    values = {"Saf.2": 1, "Saf.3": 1, "Saf.5": 2, "Saf.4.2": 1000}
    v = lambda i: values[i]
    assert engine.FORMULAS["Saf.6"](v) == pytest.approx((1 + 1) / 1000 * 1_000_000)
    assert engine.FORMULAS["Saf.7"](v) == pytest.approx(2 / 1000 * 1_000)


def test_compute_bu_month_resolves_dependencies_regardless_of_order():
    # Ene.10/Ene.11 depend on the result of Ene.9, computed in the same pass.
    raw = {
        "Ene.1": 100, "Ene.2": 10, "Ene.3": 10, "Ene.4": 10, "Ene.5": 0,
        "Ene.6": 0, "Ene.6.1": 0, "Ene.7": 0, "Ene.7.1": 0, "Ene.8": 0,
        "Env.1": 50, "Wat.1": 25,
        "Was.1": 40, "Was.2": 4,
        "Saf.2": 1, "Saf.3": 1, "Saf.5": 2, "Saf.4.2": 1000,
        "Ref.2": 1, "Ref.3": 1, "Ref.4": 0, "Ref.5": 0, "Ref.6": 0,
        "Ref.7": 0, "Ref.8": 0, "Ref.9": 0,
    }
    result = engine.compute_bu_month(raw)

    assert result["Ene.9"] == pytest.approx(130)
    assert result["Ene.10"] == pytest.approx((10 + 10 + 10) / 130)
    assert result["Ene.11"] == pytest.approx(130 / 50)
    assert result["Wat.2"] == pytest.approx(25 / 50)
    assert result["Ref.1"] == pytest.approx(2)
    # Original entries stay unchanged in the result.
    assert result["Ene.1"] == 100


def test_compute_bu_month_leaves_indicator_absent_when_input_missing():
    # No data for Wat.1/Env.1: Wat.2 must stay absent, not crash, not put in
    # a bogus value.
    raw = {"Ene.1": 5}
    result = engine.compute_bu_month(raw)
    assert "Wat.2" not in result
    assert result["Ene.1"] == 5


def test_run_merges_computed_and_input_values_with_reference_metadata():
    reference = pd.DataFrame([
        {"ID": "Wat.1", "Topic": "Water", "KPI": "Water consumption", "Unit": "m3", "Kind": "input"},
        {"ID": "Env.1", "Topic": "General info", "KPI": "Sales", "Unit": "kEUR", "Kind": "input"},
        {"ID": "Wat.2", "Topic": "Water", "KPI": "Water intensity", "Unit": "m3/kEUR", "Kind": "calculated"},
    ])
    raw = pd.DataFrame([
        {"BU": "BAF", "Month": "January", "ID": "Wat.1", "Value": 100.0},
        {"BU": "BAF", "Month": "January", "ID": "Env.1", "Value": 50.0},
    ])

    out = engine.run(reference, raw)

    wat2 = out[out.ID == "Wat.2"].iloc[0]
    assert wat2["Value"] == pytest.approx(2.0)
    assert wat2["Kind"] == "calculated"
    assert wat2["Topic"] == "Water"
    assert set(out.ID) == {"Wat.1", "Env.1", "Wat.2"}


def test_run_ignores_computed_id_absent_from_reference():
    # Wat.2 is computable (Wat.1 and Env.1 are entered) but deliberately
    # absent from the reference list (indicator removed): run() must not
    # crash, just silently skip it, along with Env.1 likewise absent.
    reference = pd.DataFrame([
        {"ID": "Wat.1", "Topic": "Water", "KPI": "Water consumption", "Unit": "m3", "Kind": "input"},
    ])
    raw = pd.DataFrame([
        {"BU": "BAF", "Month": "January", "ID": "Wat.1", "Value": 100.0},
        {"BU": "BAF", "Month": "January", "ID": "Env.1", "Value": 50.0},
    ])
    out = engine.run(reference, raw)
    assert list(out.ID) == ["Wat.1"]
