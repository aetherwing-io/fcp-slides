"""Chart type mapping for presentation charts.

Maps friendly names to python-pptx chart type constants.
"""

from __future__ import annotations

from pptx.enum.chart import XL_CHART_TYPE

# Friendly name → XL_CHART_TYPE constant
CHART_TYPES: dict[str, int] = {
    "column": XL_CHART_TYPE.COLUMN_CLUSTERED,
    "column-stacked": XL_CHART_TYPE.COLUMN_STACKED,
    "column-100": XL_CHART_TYPE.COLUMN_STACKED_100,
    "bar": XL_CHART_TYPE.BAR_CLUSTERED,
    "bar-stacked": XL_CHART_TYPE.BAR_STACKED,
    "bar-100": XL_CHART_TYPE.BAR_STACKED_100,
    "line": XL_CHART_TYPE.LINE,
    "line-markers": XL_CHART_TYPE.LINE_MARKERS,
    "line-stacked": XL_CHART_TYPE.LINE_STACKED,
    "pie": XL_CHART_TYPE.PIE,
    "pie-exploded": XL_CHART_TYPE.PIE_EXPLODED,
    "doughnut": XL_CHART_TYPE.DOUGHNUT,
    "area": XL_CHART_TYPE.AREA,
    "area-stacked": XL_CHART_TYPE.AREA_STACKED,
    "scatter": XL_CHART_TYPE.XY_SCATTER,
    "scatter-lines": XL_CHART_TYPE.XY_SCATTER_LINES,
    "radar": XL_CHART_TYPE.RADAR,
    "radar-filled": XL_CHART_TYPE.RADAR_FILLED,
}


def resolve_chart_type(name: str) -> int | None:
    """Resolve a friendly chart type name to XL_CHART_TYPE constant."""
    return CHART_TYPES.get(name.lower())


def list_chart_types() -> list[str]:
    """Return sorted list of available chart type names."""
    return sorted(CHART_TYPES.keys())
