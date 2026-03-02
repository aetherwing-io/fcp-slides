"""Shape type mapping for auto-shape creation.

Maps friendly names to MSO_AUTO_SHAPE_TYPE enum values.
"""

from __future__ import annotations

from pptx.enum.shapes import MSO_SHAPE

# Friendly name → MSO_SHAPE constant
SHAPE_TYPES: dict[str, int] = {
    "rectangle": MSO_SHAPE.RECTANGLE,
    "rect": MSO_SHAPE.RECTANGLE,
    "rounded-rect": MSO_SHAPE.ROUNDED_RECTANGLE,
    "oval": MSO_SHAPE.OVAL,
    "circle": MSO_SHAPE.OVAL,
    "diamond": MSO_SHAPE.DIAMOND,
    "triangle": MSO_SHAPE.ISOSCELES_TRIANGLE,
    "right-triangle": MSO_SHAPE.RIGHT_TRIANGLE,
    "pentagon": MSO_SHAPE.PENTAGON,
    "hexagon": MSO_SHAPE.HEXAGON,
    "octagon": MSO_SHAPE.OCTAGON,
    "star-5": MSO_SHAPE.STAR_5_POINT,
    "star-4": MSO_SHAPE.STAR_4_POINT,
    "arrow-right": MSO_SHAPE.RIGHT_ARROW,
    "arrow-left": MSO_SHAPE.LEFT_ARROW,
    "arrow-up": MSO_SHAPE.UP_ARROW,
    "arrow-down": MSO_SHAPE.DOWN_ARROW,
    "arrow-left-right": MSO_SHAPE.LEFT_RIGHT_ARROW,
    "arrow-up-down": MSO_SHAPE.UP_DOWN_ARROW,
    "chevron": MSO_SHAPE.CHEVRON,
    "cross": MSO_SHAPE.CROSS,
    "heart": MSO_SHAPE.HEART,
    "cloud": MSO_SHAPE.CLOUD,
    "callout-rect": MSO_SHAPE.RECTANGULAR_CALLOUT,
    "callout-rounded": MSO_SHAPE.ROUNDED_RECTANGULAR_CALLOUT,
    "callout-oval": MSO_SHAPE.OVAL_CALLOUT,
    "callout-cloud": MSO_SHAPE.CLOUD_CALLOUT,
    "brace-left": MSO_SHAPE.LEFT_BRACE,
    "brace-right": MSO_SHAPE.RIGHT_BRACE,
    "bracket-left": MSO_SHAPE.LEFT_BRACKET,
    "bracket-right": MSO_SHAPE.RIGHT_BRACKET,
    "cube": MSO_SHAPE.CUBE,
    "cylinder": MSO_SHAPE.CAN,
    "donut": MSO_SHAPE.DONUT,
    "no-symbol": MSO_SHAPE.NO_SYMBOL,
    "lightning": MSO_SHAPE.LIGHTNING_BOLT,
    "sun": MSO_SHAPE.SUN,
    "moon": MSO_SHAPE.MOON,
    "smiley": MSO_SHAPE.SMILEY_FACE,
    "plus": MSO_SHAPE.MATH_PLUS,
    "minus": MSO_SHAPE.MATH_MINUS,
    "flowchart-process": MSO_SHAPE.FLOWCHART_PROCESS,
    "flowchart-decision": MSO_SHAPE.FLOWCHART_DECISION,
    "flowchart-terminator": MSO_SHAPE.FLOWCHART_TERMINATOR,
    "flowchart-data": MSO_SHAPE.FLOWCHART_DATA,
}


def resolve_shape_type(name: str) -> int | None:
    """Resolve a friendly shape name to MSO_SHAPE constant.

    Returns None if the name is not recognized.
    """
    return SHAPE_TYPES.get(name.lower())


def list_shape_types() -> list[str]:
    """Return sorted list of available shape type names."""
    return sorted(set(SHAPE_TYPES.keys()))
