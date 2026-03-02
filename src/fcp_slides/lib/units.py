"""EMU unit conversion for human-friendly length specifications.

PowerPoint uses English Metric Units (EMUs) internally.
This module converts from human-friendly syntax like 2in, 5cm, 72pt.
"""

from __future__ import annotations

import re

EMU_PER_INCH = 914400
EMU_PER_CM = 360000
EMU_PER_PT = 12700
EMU_PER_PX = 9525  # at 96 DPI

_LENGTH_RE = re.compile(
    r"^(-?\d+(?:\.\d+)?)\s*(in|cm|pt|px|emu)?$", re.IGNORECASE
)


def parse_length(s: str) -> int:
    """Parse a length string to EMUs.

    Accepts:
      - "2in"    → 1828800 EMU
      - "5cm"    → 1800000 EMU
      - "72pt"   → 914400 EMU
      - "100px"  → 952500 EMU
      - "914400" → 914400 EMU (bare number = EMU)
      - "2.5in"  → 2286000 EMU

    Raises ValueError if the string cannot be parsed.
    """
    s = s.strip()
    m = _LENGTH_RE.match(s)
    if not m:
        raise ValueError(f"Invalid length: {s!r}")

    value = float(m.group(1))
    unit = (m.group(2) or "emu").lower()

    if unit == "in":
        return int(value * EMU_PER_INCH)
    elif unit == "cm":
        return int(value * EMU_PER_CM)
    elif unit == "pt":
        return int(value * EMU_PER_PT)
    elif unit == "px":
        return int(value * EMU_PER_PX)
    elif unit == "emu":
        return int(value)

    raise ValueError(f"Unknown unit: {unit!r}")


def format_length(emu: int) -> str:
    """Format EMUs as a human-friendly string.

    Returns the most natural unit representation.
    """
    if emu % EMU_PER_INCH == 0:
        val = emu // EMU_PER_INCH
        return f"{val}in"
    if emu % EMU_PER_CM == 0:
        val = emu // EMU_PER_CM
        return f"{val}cm"

    # Default to inches with decimal
    inches = emu / EMU_PER_INCH
    if inches == int(inches):
        return f"{int(inches)}in"
    return f"{inches:.2f}in"
