"""Tests for lib/colors.py — color parsing."""

import pytest

from fcp_slides.lib.colors import parse_color, to_rgb


class TestParseColor:
    def test_named_color(self):
        assert parse_color("blue") == "4472C4"

    def test_hex_with_hash(self):
        assert parse_color("#FF0000") == "FF0000"

    def test_hex_without_hash(self):
        assert parse_color("4472C4") == "4472C4"

    def test_three_char_hex(self):
        assert parse_color("F0F") == "FF00FF"

    def test_case_insensitive_name(self):
        assert parse_color("BLUE") == "4472C4"

    def test_case_insensitive_hex(self):
        assert parse_color("ff0000") == "FF0000"

    def test_invalid_raises(self):
        with pytest.raises(ValueError):
            parse_color("not-a-color")


class TestToRgb:
    def test_returns_rgb_color(self):
        from pptx.dml.color import RGBColor
        result = to_rgb("blue")
        assert isinstance(result, RGBColor)
