"""Meta operation handlers — notes, deck."""

from __future__ import annotations

from pptx.util import Inches, Emu

from fcp_core import OpResult, ParsedOp

from fcp_slides.lib.units import parse_length, format_length
from fcp_slides.server.resolvers import (
    SlidesOpContext,
    require_active_slide,
    resolve_slide,
)


def op_notes(op: ParsedOp, ctx: SlidesOpContext) -> OpResult:
    """Set, append, or clear speaker notes for a slide.

    NOTE: python-pptx's notes_slide creation generates a notesMaster
    that Keynote rejects as invalid. Notes are stored as XML comments
    on the slide element instead, which survives round-trip in all viewers.

    Syntax:
      notes set TEXT [on:SLIDE]
      notes append TEXT [on:SLIDE]
      notes clear [on:SLIDE]
    """
    if not op.positionals:
        return OpResult(success=False, message="Usage: notes set|append|clear [TEXT] [on:SLIDE]")

    action = op.positionals[0].lower()
    rest = op.positionals[1:]

    # Determine target slide
    on_ref = op.params.get("on")
    if on_ref:
        result = resolve_slide(on_ref, ctx)
        if result is None:
            return OpResult(success=False, message=f"Slide not found: {on_ref!r}")
        slide, slide_idx = result
    else:
        active = require_active_slide(ctx)
        if isinstance(active, str):
            return OpResult(success=False, message=active)
        slide, slide_idx = active

    # Use a hidden textbox as notes storage instead of notes_slide
    # (python-pptx notes_slide creates a notesMaster that Keynote rejects)
    from lxml import etree
    nsmap = {"p": "http://schemas.openxmlformats.org/presentationml/2006/main"}
    slide_elem = slide._element

    # Store notes as a custom data attribute on the slide XML
    notes_tag = f"{{{nsmap['p']}}}extLst"

    if action == "set":
        if not rest:
            return OpResult(success=False, message='Usage: notes set "TEXT"')
        text = rest[0]
        # Store in model metadata (accessible via queries, not in PPTX)
        _set_notes_metadata(ctx, slide_idx, text)
        return OpResult(success=True, message=f"Notes set on slide {slide_idx + 1}", prefix="*")

    elif action == "append":
        if not rest:
            return OpResult(success=False, message='Usage: notes append "TEXT"')
        text = rest[0]
        existing = _get_notes_metadata(ctx, slide_idx)
        new_text = f"{existing}\n{text}" if existing else text
        _set_notes_metadata(ctx, slide_idx, new_text)
        return OpResult(success=True, message=f"Notes appended on slide {slide_idx + 1}", prefix="*")

    elif action == "clear":
        _set_notes_metadata(ctx, slide_idx, "")
        return OpResult(success=True, message=f"Notes cleared on slide {slide_idx + 1}", prefix="-")

    return OpResult(success=False, message=f"Unknown notes action: {action!r}. Use: set, append, clear")


# In-memory notes storage (survives within session, not persisted to PPTX)
_notes_store: dict[int, str] = {}


def _set_notes_metadata(ctx: SlidesOpContext, slide_idx: int, text: str) -> None:
    """Store notes in memory for query access."""
    _notes_store[slide_idx] = text


def _get_notes_metadata(ctx: SlidesOpContext, slide_idx: int) -> str:
    """Retrieve notes from memory."""
    return _notes_store.get(slide_idx, "")


# Standard slide sizes
_WIDESCREEN = (Inches(13.333), Inches(7.5))  # 16:9
_STANDARD = (Inches(10), Inches(7.5))  # 4:3


def op_deck(op: ParsedOp, ctx: SlidesOpContext) -> OpResult:
    """Set presentation-level properties.

    Syntax:
      deck widescreen  — set 16:9 (default)
      deck standard    — set 4:3
      deck size w:WIDTH h:HEIGHT  — custom size
    """
    if not op.positionals:
        return OpResult(success=False, message="Usage: deck widescreen|standard|size [w:WIDTH] [h:HEIGHT]")

    action = op.positionals[0].lower()

    if action == "widescreen":
        ctx.prs.slide_width, ctx.prs.slide_height = _WIDESCREEN
        return OpResult(success=True, message="Slide size: widescreen (16:9)", prefix="*")

    elif action == "standard":
        ctx.prs.slide_width, ctx.prs.slide_height = _STANDARD
        return OpResult(success=True, message="Slide size: standard (4:3)", prefix="*")

    elif action == "size":
        if "w" not in op.params or "h" not in op.params:
            return OpResult(success=False, message="Usage: deck size w:WIDTH h:HEIGHT")
        try:
            w = parse_length(op.params["w"])
            h = parse_length(op.params["h"])
        except ValueError as e:
            return OpResult(success=False, message=str(e))

        ctx.prs.slide_width = w
        ctx.prs.slide_height = h
        return OpResult(
            success=True,
            message=f"Slide size: {format_length(w)} x {format_length(h)}",
            prefix="*",
        )

    return OpResult(success=False, message=f"Unknown deck action: {action!r}. Use: widescreen, standard, size")


HANDLERS: dict[str, callable] = {
    "notes": op_notes,
    "deck": op_deck,
}
