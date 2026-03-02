"""Named color palette and hex parsing for presentation formatting."""

from __future__ import annotations

from pptx.util import Pt
from pptx.dml.color import RGBColor

# Standard Office/PowerPoint color palette
NAMED_COLORS: dict[str, str] = {
    "blue": "4472C4",
    "orange": "ED7D31",
    "gray": "A5A5A5",
    "gold": "FFC000",
    "lt-blue": "5B9BD5",
    "green": "70AD47",
    "red": "FF0000",
    "dk-green": "00B050",
    "white": "FFFFFF",
    "black": "000000",
    "yellow": "FFFF00",
    "purple": "7030A0",
    "teal": "00B0F0",
    "dk-blue": "002060",
    "dk-red": "C00000",
    "lt-gray": "D9D9D9",
    "dk-gray": "404040",
}


def parse_color(color_str: str) -> str:
    """Parse a color string to a 6-char hex value (no #).

    Accepts:
      - Named colors: "blue", "red"
      - Hex with #: "#4472C4"
      - Hex without #: "4472C4"
      - 3-char shorthand: "F0F" → "FF00FF"
    """
    name = color_str.lower().strip()
    if name in NAMED_COLORS:
        return NAMED_COLORS[name]

    hex_str = color_str.lstrip("#").strip()

    if len(hex_str) == 6 and all(c in "0123456789ABCDEFabcdef" for c in hex_str):
        return hex_str.upper()

    if len(hex_str) == 3 and all(c in "0123456789ABCDEFabcdef" for c in hex_str):
        return "".join(c + c for c in hex_str).upper()

    raise ValueError(f"Invalid color: {color_str!r}")


def to_rgb(color_str: str) -> RGBColor:
    """Parse a color string directly to an RGBColor object."""
    hex_str = parse_color(color_str)
    return RGBColor.from_string(hex_str)
