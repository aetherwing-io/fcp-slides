"""Tests for lib/units.py — EMU unit conversion."""

import pytest

from fcp_slides.lib.units import parse_length, format_length, EMU_PER_INCH, EMU_PER_CM


class TestParseLength:
    def test_inches(self):
        assert parse_length("2in") == 2 * EMU_PER_INCH

    def test_centimeters(self):
        assert parse_length("5cm") == 5 * EMU_PER_CM

    def test_points(self):
        assert parse_length("72pt") == 72 * 12700

    def test_pixels(self):
        assert parse_length("100px") == 100 * 9525

    def test_bare_number_is_emu(self):
        assert parse_length("914400") == 914400

    def test_decimal_inches(self):
        assert parse_length("2.5in") == int(2.5 * EMU_PER_INCH)

    def test_case_insensitive(self):
        assert parse_length("2IN") == 2 * EMU_PER_INCH

    def test_whitespace(self):
        assert parse_length("  2in  ") == 2 * EMU_PER_INCH

    def test_invalid_raises(self):
        with pytest.raises(ValueError):
            parse_length("abc")

    def test_empty_raises(self):
        with pytest.raises(ValueError):
            parse_length("")


class TestFormatLength:
    def test_whole_inches(self):
        assert format_length(2 * EMU_PER_INCH) == "2in"

    def test_whole_cm(self):
        assert format_length(5 * EMU_PER_CM) == "5cm"

    def test_fractional_inches(self):
        result = format_length(int(2.5 * EMU_PER_INCH))
        assert "2.5" in result
