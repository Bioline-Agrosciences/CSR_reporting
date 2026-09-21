"""Tests for scripts/extract_reference_and_data.py: only the pure function
parse_raw_value (cleaning malformed "Monthly value" cells). build_reference/
build_raw_long need the actual original raw Excel workbooks (raw_data/),
not available in CI — not tested here."""
import pytest

import extract_reference_and_data as extract


def test_none_stays_none_without_note():
    assert extract.parse_raw_value(None) == (None, None)


@pytest.mark.parametrize("raw", [42, 42.0, 0, -3.5])
def test_number_passes_through_unchanged(raw):
    value, note = extract.parse_raw_value(raw)
    assert value == float(raw)
    assert note is None


def test_french_decimal_comma_is_converted_and_flagged():
    value, note = extract.parse_raw_value("209,3")
    assert value == pytest.approx(209.3)
    assert "decimal comma" in note


def test_french_decimal_with_surrounding_whitespace():
    value, note = extract.parse_raw_value("  54,41 ")
    assert value == pytest.approx(54.41)
    assert "decimal comma" in note


def test_plain_numeric_text_is_converted_and_flagged():
    value, note = extract.parse_raw_value("56.7")
    assert value == pytest.approx(56.7)
    assert "numeric text" in note


def test_unit_mixed_into_number_is_unparsable():
    value, note = extract.parse_raw_value("28,9kWh")
    assert value is None
    assert "NOT PARSABLE" in note
    assert "28,9kWh" in note


def test_free_text_is_unparsable():
    value, note = extract.parse_raw_value("about 30")
    assert value is None
    assert "NOT PARSABLE" in note


def test_unexpected_type_is_flagged():
    value, note = extract.parse_raw_value(["not", "a", "scalar"])
    assert value is None
    assert "Unexpected type" in note
