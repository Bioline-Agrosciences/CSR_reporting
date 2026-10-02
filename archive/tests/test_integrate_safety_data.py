"""
Tests for integrate_safety_data.py — written before wiring the script into
the real monthly run, against small synthetic workbooks (no dependency on
the real Working Hours / BlueKanGo files) so they run anywhere, fast.

Usage
-----
    uv run pytest tests/test_integrate_safety_data.py -v
    (or just `uv run pytest` from the project root — picked up automatically)
"""
import sys
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

# tests/ is a sibling of scripts/, not nested inside it (see Project layout
# in README.md) — so the path to add is ../scripts, not just the parent.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "scripts"))
import integrate_safety_data as isd


# ---------------------------------------------------------------------------
# read_working_hours
# ---------------------------------------------------------------------------
def _write_working_hours(path, rows):
    """rows: list of (bu, {month: value_or_None}) — only January/February
    used in these tests, matching isd.MONTHS ordering."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Working hours (Bioline Employees)", "January", "February"])
    for bu, months in rows:
        ws.append([bu, months.get("January"), months.get("February")])
    wb.save(path)


def test_read_working_hours_skips_total_row(tmp_path):
    path = tmp_path / "hours.xlsx"
    _write_working_hours(path, [
        ("BAF", {"January": 1000, "February": 900}),
        ("Total", {"January": 1000, "February": 900}),
    ])
    df = isd.read_working_hours(path)
    assert set(df["BU"]) == {"BAF"}
    assert len(df) == 2  # BAF x (January, February), Total excluded


def test_read_working_hours_skips_undeclared_months(tmp_path):
    path = tmp_path / "hours.xlsx"
    _write_working_hours(path, [("BAF", {"January": 1000, "February": None})])
    df = isd.read_working_hours(path)
    assert list(df["Month"]) == ["January"]  # February not yet declared -> left out, not 0


def test_read_working_hours_produces_saf_4_2_rows(tmp_path):
    path = tmp_path / "hours.xlsx"
    _write_working_hours(path, [("BAF", {"January": 1000})])
    df = isd.read_working_hours(path)
    row = df.iloc[0]
    assert row["ID"] == "Saf.4.2"
    assert row["Value"] == 1000.0


def test_read_working_hours_tags_target_year(tmp_path):
    """Added 22/09/2026: every row is tagged with target_year, so the same
    (BU, Month) doesn't collide across two different years' data once it
    reaches Raw_data_CSR.xlsx."""
    path = tmp_path / "hours.xlsx"
    _write_working_hours(path, [("BAF", {"January": 1000})])
    df = isd.read_working_hours(path, target_year=2027)
    assert df.iloc[0]["Year"] == 2027


def test_read_working_hours_defaults_to_current_year(tmp_path):
    path = tmp_path / "hours.xlsx"
    _write_working_hours(path, [("BAF", {"January": 1000})])
    df = isd.read_working_hours(path)  # no target_year given
    assert df.iloc[0]["Year"] == date.today().year


# ---------------------------------------------------------------------------
# read_accidents
# ---------------------------------------------------------------------------
def _write_accidents(path, rows):
    """rows: list of dicts with keys bu, date, interim, arret, days_lost."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Accidents du travail Bioline"])
    ws.append(["N°", "Business Unit", "Date de l'accident",
               "Intérimaire / Agency / non Bioline employees", "Arrêt de travail",
               "Durée totale de l'arrêt (calcul)"])
    for i, r in enumerate(rows, start=1):
        ws.append([i, r["bu"], r["date"], r.get("interim", "Non"), r["arret"], r.get("days_lost")])
    wb.save(path)


def test_read_accidents_splits_lti_and_nlti(tmp_path, capsys):
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline Africa", "date": datetime(2026, 8, 7), "arret": "OUI", "days_lost": 2},
        {"bu": "Bioline Africa", "date": datetime(2026, 8, 10), "arret": "NON"},
    ])
    df = isd.read_accidents(target_year=2026, path=path)
    values = dict(zip(df["ID"], df["Value"]))
    assert values["Saf.3"] == 1     # 1 LTI
    assert values["Saf.2"] == 1     # 1 NLTI
    assert values["Saf.5"] == 2     # 2 days lost


def test_read_accidents_excludes_interim_workers(tmp_path):
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline UK", "date": datetime(2026, 6, 3), "interim": "Oui", "arret": "NON"},
    ])
    df = isd.read_accidents(target_year=2026, path=path)
    assert df.empty  # the only row is an agency worker -> excluded


def test_read_accidents_excludes_other_years(tmp_path):
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline France", "date": datetime(2025, 11, 27), "arret": "OUI", "days_lost": 5},
    ])
    df = isd.read_accidents(target_year=2026, path=path)
    assert df.empty


def test_read_accidents_warns_on_unmapped_bu(tmp_path, capsys):
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline Atlantis", "date": datetime(2026, 3, 1), "arret": "NON"},
    ])
    df = isd.read_accidents(target_year=2026, path=path)
    assert df.empty
    assert "Bioline Atlantis" in capsys.readouterr().out


def test_read_accidents_tags_target_year(tmp_path):
    """Added 22/09/2026: Saf.2/Saf.3/Saf.5 rows carry the Year they were
    aggregated for, matching the Year on the Saf.4.2 rows from
    read_working_hours — both flow into the same upsert key."""
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline Africa", "date": datetime(2026, 8, 7), "arret": "NON"},
    ])
    df = isd.read_accidents(target_year=2026, path=path)
    assert set(df["Year"]) == {2026}


@pytest.mark.parametrize("label,code", [
    ("Bioline Africa", "BAF"), ("Bioline US", "BUS"), ("Bioline UK", "BUK"),
    ("Bioline Iberia", "BIB"), ("Bioline France", "BFR"),
])
def test_bu_name_map_known_labels(label, code):
    assert isd.BU_NAME_MAP[label] == code


# ---------------------------------------------------------------------------
# compute_days_without_accident (Saf.1)
# ---------------------------------------------------------------------------
def _working_hours_df(pairs, year=2026):
    return pd.DataFrame([
        {"BU": bu, "Year": year, "Month": month, "ID": "Saf.4.2", "Value": 1000,
         "Comment": None, "Data quality note": None}
        for bu, month in pairs
    ])


def test_days_without_accident_counts_from_last_accident(tmp_path):
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline Africa", "date": datetime(2026, 8, 7), "arret": "NON"},
    ])
    events = isd._read_accident_events(path)
    wh = _working_hours_df([("BAF", "August")])
    out = isd.compute_days_without_accident(events, wh, target_year=2026)
    # last day of August 2026 - Aug 7 2026 = 24 days
    assert out.iloc[0]["ID"] == "Saf.1"
    assert out.iloc[0]["Value"] == 24
    assert out.iloc[0]["Year"] == 2026


def test_days_without_accident_year_comes_from_working_hours_not_fallback(tmp_path):
    """Added 22/09/2026: the Saf.1 row's Year should reflect the (BU, Month)
    pair's own reporting year — taken from working_hours_df's "Year" column
    — not the target_year fallback, which only matters for callers that
    don't pass a "Year" column at all (kept for backward compatibility)."""
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline Africa", "date": datetime(2027, 8, 7), "arret": "NON"},
    ])
    events = isd._read_accident_events(path)
    wh = _working_hours_df([("BAF", "August")], year=2027)
    # target_year deliberately mismatched (2026) — the working_hours_df's own
    # Year (2027) must win.
    out = isd.compute_days_without_accident(events, wh, target_year=2026)
    assert out.iloc[0]["Year"] == 2027


def test_days_without_accident_any_accident_type_resets_counter(tmp_path):
    """A non-lost-time accident also resets the counter — Aurélie said
    "aucun accident", not "aucun accident avec arrêt"."""
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline Africa", "date": datetime(2026, 8, 20), "arret": "NON"},
    ])
    events = isd._read_accident_events(path)
    wh = _working_hours_df([("BAF", "August")])
    out = isd.compute_days_without_accident(events, wh, target_year=2026)
    assert out.iloc[0]["Value"] == 11  # Aug 31 - Aug 20


def test_days_without_accident_looks_back_beyond_target_year(tmp_path):
    """Unlike Saf.2/Saf.3/Saf.5, Saf.1 is not restricted to the target
    year — a late-previous-year accident still counts as the most recent
    one for an early month of the new year."""
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline France", "date": datetime(2025, 12, 15), "arret": "OUI", "days_lost": 3},
    ])
    events = isd._read_accident_events(path)
    wh = _working_hours_df([("BFR", "January")])
    out = isd.compute_days_without_accident(events, wh, target_year=2026)
    assert out.iloc[0]["Value"] == 47  # Jan 31 2026 - Dec 15 2025


def test_days_without_accident_skipped_when_no_accident_on_record(tmp_path):
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [])
    events = isd._read_accident_events(path)
    wh = _working_hours_df([("BAF", "August")])
    out = isd.compute_days_without_accident(events, wh, target_year=2026)
    assert out.empty  # no accident on record at all -> no guessed start date


def test_days_without_accident_ignores_interim_workers(tmp_path):
    path = tmp_path / "accidents.xlsx"
    _write_accidents(path, [
        {"bu": "Bioline Africa", "date": datetime(2026, 8, 30), "interim": "Oui", "arret": "NON"},
    ])
    events = isd._read_accident_events(path)
    assert events.empty  # agency worker -> excluded even from the event list Saf.1 uses


# ---------------------------------------------------------------------------
# fill_zero_accident_months
# ---------------------------------------------------------------------------
def test_fill_zero_accident_months_adds_explicit_zeros():
    working_hours_df = _working_hours_df([("BAF", "January")])
    accidents_df = pd.DataFrame(columns=["BU", "Year", "Month", "ID", "Value", "Comment", "Data quality note"])
    out = isd.fill_zero_accident_months(accidents_df, working_hours_df)
    values = dict(zip(out["ID"], out["Value"]))
    assert values == {"Saf.2": 0, "Saf.3": 0, "Saf.5": 0}
    assert set(out["Year"]) == {2026}


def test_fill_zero_accident_months_does_not_overwrite_real_values():
    working_hours_df = _working_hours_df([("BAF", "January")])
    accidents_df = pd.DataFrame([
        {"BU": "BAF", "Year": 2026, "Month": "January", "ID": "Saf.3", "Value": 1,
         "Comment": None, "Data quality note": None},
    ])
    out = isd.fill_zero_accident_months(accidents_df, working_hours_df)
    values = dict(zip(out["ID"], out["Value"]))
    assert values["Saf.3"] == 1     # real value kept, not overwritten by the 0-filler
    assert values["Saf.2"] == 0     # untouched ID still filled with 0


def test_fill_zero_accident_months_does_not_cross_years():
    """Added 22/09/2026: a real Saf.3 value for BAF/January/2027 must not be
    treated as "already covered" for BAF/January/2026 — each year gets its
    own explicit 0-filling."""
    working_hours_df = pd.concat([
        _working_hours_df([("BAF", "January")], year=2026),
        _working_hours_df([("BAF", "January")], year=2027),
    ], ignore_index=True)
    accidents_df = pd.DataFrame([
        {"BU": "BAF", "Year": 2027, "Month": "January", "ID": "Saf.3", "Value": 1,
         "Comment": None, "Data quality note": None},
    ])
    out = isd.fill_zero_accident_months(accidents_df, working_hours_df)
    by_year = {(r["Year"], r["ID"]): r["Value"] for r in out.to_dict("records")}
    assert by_year[(2027, "Saf.3")] == 1   # real value, untouched
    assert by_year[(2026, "Saf.3")] == 0   # 2026 gets its own explicit 0, not skipped
