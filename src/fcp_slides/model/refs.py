"""Shape reference types for slides index.

ShapeRef points into python-pptx shape objects via slide index
and shape_id for O(1) lookups after index rebuild.
"""

from __future__ import annotations

from dataclasses import dataclass


@dataclass
class ShapeRef:
    """Reference to a shape within a presentation."""

    label: str
    slide_idx: int  # 0-based index into prs.slides
    shape_id: int  # shape.shape_id in python-pptx
    shape_type: str  # "textbox", "rectangle", "table", "chart", "picture", etc.
    placeholder_idx: int | None = None  # placeholder index if applicable
