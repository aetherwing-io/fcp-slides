"""Slide layout name resolution.

Maps friendly names to slide layout indices in standard PowerPoint templates.
Falls back to fuzzy matching against the actual layout names in the presentation.
"""

from __future__ import annotations

from difflib import SequenceMatcher

# Friendly name → common layout name in standard templates
LAYOUT_ALIASES: dict[str, str] = {
    "title": "Title Slide",
    "title-slide": "Title Slide",
    "title-content": "Title and Content",
    "section": "Section Header",
    "section-header": "Section Header",
    "two-content": "Two Content",
    "comparison": "Comparison",
    "title-only": "Title Only",
    "blank": "Blank",
    "content-caption": "Content with Caption",
    "picture-caption": "Picture with Caption",
}


def resolve_layout(name: str, prs) -> "SlideLayout | None":
    """Resolve a layout name to a SlideLayout object.

    Tries in order:
      1. Friendly alias (e.g., "title" → "Title Slide")
      2. Exact match against actual layout names
      3. Case-insensitive match
      4. Fuzzy match (>0.6 similarity)

    Returns None if no match found.
    """
    # Try alias first
    alias = LAYOUT_ALIASES.get(name.lower())
    target = alias or name

    layouts = prs.slide_layouts

    # Exact match
    for layout in layouts:
        if layout.name == target:
            return layout

    # Case-insensitive match
    target_lower = target.lower()
    for layout in layouts:
        if layout.name.lower() == target_lower:
            return layout

    # Fuzzy match
    best_match = None
    best_score = 0.0
    for layout in layouts:
        score = SequenceMatcher(None, target_lower, layout.name.lower()).ratio()
        if score > best_score:
            best_score = score
            best_match = layout

    if best_match and best_score > 0.6:
        return best_match

    return None


def list_layouts(prs) -> list[str]:
    """Return list of available layout names in the presentation."""
    return [layout.name for layout in prs.slide_layouts]
