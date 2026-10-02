"""Tests for scripts/csr_parameters.py: the (Parameter, BU, Year) lookup of
the emission/conversion factors in CSR_parameters.xlsx, its fallback to an
earlier year, and the checks surfaced in the anomaly report."""
import pandas as pd
import pytest

import csr_parameters as params
import migrate_2026_10_create_parameters_file as migration


def _parameters(*rows):
    return pd.DataFrame([{"Parameter": p, "BU": bu, "Year": y, "Value": v, "Unit": "", "Source": ""}
                         for p, bu, y, v in rows], columns=params.PARAMETER_COLS)


def test_resolve_parameter_uses_the_exact_year_when_present():
    index = params.build_parameter_index(_parameters(("F", "BFR", 2026, 1.0), ("F", "BFR", 2027, 2.0)))
    assert params.resolve_parameter(index, "F", "BFR", 2026) == (1.0, 2026)
    assert params.resolve_parameter(index, "F", "BFR", 2027) == (2.0, 2027)


def test_resolve_parameter_falls_back_to_the_most_recent_earlier_year():
    index = params.build_parameter_index(_parameters(("F", "BFR", 2025, 1.0), ("F", "BFR", 2026, 2.0)))
    assert params.resolve_parameter(index, "F", "BFR", 2028) == (2.0, 2026)


def test_resolve_parameter_never_uses_a_later_year():
    index = params.build_parameter_index(_parameters(("F", "BFR", 2027, 2.0)))
    with pytest.raises(params.MissingParameterError):
        params.resolve_parameter(index, "F", "BFR", 2026)


def test_resolve_parameter_is_per_bu():
    index = params.build_parameter_index(_parameters(("F", "BFR", 2026, 1.0)))
    with pytest.raises(params.MissingParameterError):
        params.resolve_parameter(index, "F", "BAF", 2026)


@pytest.mark.parametrize("bad,message", [
    (_parameters(("F", "BFR", 2026, 1.0), ("F", "BFR", 2026, 2.0)), "more than one row"),
    (_parameters(("F", "BFR", 2026, None)), "empty Value"),
    (pd.DataFrame({"Parameter": ["F"], "BU": ["BFR"], "Value": [1.0]}), "missing column"),
])
def test_validate_parameters_rejects_an_ambiguous_or_incomplete_file(bad, message):
    with pytest.raises(ValueError, match=message):
        params.validate_parameters(bad)


def test_find_parameter_issues_reports_fallbacks_and_missing_values():
    every_factor_2026 = [(name, "BFR", 2026, 1.0) for name in set(params.PARAMETER_FOR_ID.values())
                         if name != "LPG_EMISSION_FACTOR"]
    raw = pd.DataFrame([{"BU": "BFR", "Year": 2026}, {"BU": "BFR", "Year": 2027}])

    issues = params.find_parameter_issues(_parameters(*every_factor_2026), raw)

    by_year = issues.groupby("Year")
    # 2026: only LPG_EMISSION_FACTOR, missing everywhere.
    assert list(by_year.get_group(2026)["Parameter"]) == ["LPG_EMISSION_FACTOR"]
    # 2027: every 2026 factor carried over, plus LPG still missing.
    issues_2027 = by_year.get_group(2027).set_index("Parameter")["Issue"]
    assert len(issues_2027) == len(set(params.PARAMETER_FOR_ID.values()))
    assert issues_2027["LPG_EMISSION_FACTOR"].startswith("Missing")
    assert issues_2027["ELECTRICITY_EMISSION_FACTOR"] == "No 2027 value, 2026 value used instead"


def test_find_parameter_issues_is_empty_when_every_factor_has_its_year():
    rows = [(name, "BFR", 2026, 1.0) for name in set(params.PARAMETER_FOR_ID.values())]
    raw = pd.DataFrame([{"BU": "BFR", "Year": 2026}])
    assert params.find_parameter_issues(_parameters(*rows), raw).empty


def test_migration_covers_every_parameter_for_every_bu_with_the_former_config_values():
    df = migration.build_parameters()
    assert set(df.Parameter) == set(params.PARAMETER_FOR_ID.values())
    assert len(df) == len(set(params.PARAMETER_FOR_ID.values())) * len(migration.BUS)
    index = params.build_parameter_index(df)
    assert params.resolve_parameter(index, "ELECTRICITY_EMISSION_FACTOR", "BFR", 2026)[0] == pytest.approx(0.000035)
    assert params.resolve_parameter(index, "LPG_CONVERSION_FACTOR", "BAF", 2026)[0] == pytest.approx(13.8)
