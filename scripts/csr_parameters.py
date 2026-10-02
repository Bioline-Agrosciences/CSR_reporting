"""
Emission and conversion factors used by csr_calc_engine.py's
CONTEXTUAL_VALUES (Ene.6.1, Ene.7.1, Ene.12-15) — read from
input_data/CSR_parameters.xlsx instead of being hardcoded in config.py
(decision of 02/10/2026), for three reasons:

  - ONE source for both the local pipeline and the Fabric notebook
    (nb_consolidate_csr_data reads this same file through the OneLake
    shortcut) — no more two copies of every factor to keep in sync by hand.
  - Editable in Excel without touching the code, with a "Source" column
    tracing where each factor comes from (useful for an audit).
  - Keyed by (Parameter, BU, Year): grid emission factors are republished
    every year, and updating one must NOT silently recompute the previous
    years' already-published emissions with the new value — which is
    exactly what a single per-BU constant in config.py did.

A (BU, Year) with no row of its own for a parameter falls back to the most
recent EARLIER year — a new year's factor is often only published months
into that year, and a carried-over factor beats an empty Carb.X in January.
Never a LATER year: that would apply a factor to a period it wasn't
published for. find_parameter_issues() surfaces every fallback (and every
truly missing value) in the anomaly report, so a carried-over factor is
never silent.

Sheet "Parameters", one row per (Parameter, BU, Year):

    Parameter | BU | Year | Value | Unit | Source

Created once by migrate_2026_10_create_parameters_file.py, from the values
that used to live in config.py. To add a new year's factors: add rows with
the new Year — never edit the previous year's rows in place.

Dependencies: pandas, openpyxl.
"""
import sys
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import config

PARAMETERS_PATH = config.INPUT_DIR / "CSR_parameters.xlsx"
PARAMETERS_SHEET = "Parameters"
PARAMETER_COLS = ["Parameter", "BU", "Year", "Value", "Unit", "Source"]

# Indicator ID (as it appears in the reference list and the results) ->
# parameter name in CSR_parameters.xlsx.
PARAMETER_FOR_ID = {
    "Ene.6.1": "LPG_CONVERSION_FACTOR",
    "Ene.7.1": "FUEL_CONVERSION_FACTOR",
    "Ene.12": "ELECTRICITY_EMISSION_FACTOR",
    "Ene.13": "NATURAL_GAS_EMISSION_FACTOR",
    "Ene.14": "FUEL_EMISSION_FACTOR",
    "Ene.15": "LPG_EMISSION_FACTOR",
}


class MissingParameterError(LookupError):
    """No value for this parameter/BU, neither for the requested year nor
    for any earlier one."""


def validate_parameters(parameters: pd.DataFrame) -> None:
    """Fails loudly on a malformed file rather than computing emissions from
    an ambiguous or empty factor: required columns, no empty Value, and at
    most one row per (Parameter, BU, Year)."""
    missing_cols = [c for c in ["Parameter", "BU", "Year", "Value"] if c not in parameters.columns]
    if missing_cols:
        raise ValueError(f"{PARAMETERS_PATH.name}: missing column(s) {missing_cols}")
    empty = parameters[parameters["Value"].isna()]
    if len(empty):
        raise ValueError(f"{PARAMETERS_PATH.name}: empty Value on row(s):\n{empty.to_string(index=False)}")
    dupes = parameters[parameters.duplicated(["Parameter", "BU", "Year"], keep=False)]
    if len(dupes):
        raise ValueError(f"{PARAMETERS_PATH.name}: more than one row for the same (Parameter, BU, Year):\n"
                         f"{dupes.to_string(index=False)}")


def load_parameters(path: Path = None) -> pd.DataFrame:
    path = path or PARAMETERS_PATH
    if not Path(path).exists():
        raise FileNotFoundError(f"Parameters file not found: {path} — "
                                f"run `uv run scripts/migrate_2026_10_create_parameters_file.py` once to create it.")
    parameters = pd.read_excel(path, sheet_name=PARAMETERS_SHEET)
    validate_parameters(parameters)
    return parameters


def build_parameter_index(parameters: pd.DataFrame | None) -> dict:
    """{(Parameter, BU): [(Year, Value), ...] sorted by Year} — built once
    per run, so resolve_parameter doesn't re-filter the DataFrame for every
    (BU, Year, Month) group."""
    index = {}
    if parameters is None:
        return index
    for name, bu, year, value in parameters[["Parameter", "BU", "Year", "Value"]].itertuples(index=False):
        index.setdefault((name, bu), []).append((int(year), float(value)))
    for rows in index.values():
        rows.sort()
    return index


def resolve_parameter(index: dict, name: str, bu: str, year: int) -> tuple[float, int]:
    """(value, year it was taken from): the row for `year` itself, otherwise
    the most recent earlier year (see module docstring). Raises
    MissingParameterError if there is none."""
    candidates = [row for row in index.get((name, bu), []) if row[0] <= year]
    if not candidates:
        raise MissingParameterError(f"{name} for {bu}: no value for {year} or any earlier year")
    year_used, value = candidates[-1]
    return value, year_used


def find_parameter_issues(parameters: pd.DataFrame | None, raw: pd.DataFrame) -> pd.DataFrame:
    """For every (BU, Year) present in the raw data and every parameter: a
    row if the value was carried over from an earlier year, or is missing
    altogether (the indicators depending on it — e.g. Carb.1 for
    ELECTRICITY_EMISSION_FACTOR, Ene.9 for the conversion factors — are then
    left out of the results rather than computed with a 0)."""
    index = build_parameter_index(parameters)
    rows = []
    for bu, year in sorted(set(zip(raw["BU"], raw["Year"].astype(int)))):
        for name in sorted(set(PARAMETER_FOR_ID.values())):
            try:
                _, year_used = resolve_parameter(index, name, bu, year)
            except MissingParameterError:
                rows.append({"BU": bu, "Year": year, "Parameter": name,
                             "Issue": "Missing: dependent indicators not computed"})
                continue
            if year_used != year:
                rows.append({"BU": bu, "Year": year, "Parameter": name,
                             "Issue": f"No {year} value, {year_used} value used instead"})
    return pd.DataFrame(rows, columns=["BU", "Year", "Parameter", "Issue"])
