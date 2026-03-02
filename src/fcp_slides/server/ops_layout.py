"""Layout operation handlers — align, distribute, z-order."""

from __future__ import annotations

from pptx.util import Emu

from fcp_core import OpResult, ParsedOp

from fcp_slides.server.resolvers import (
    SlidesOpContext,
    require_active_slide,
    resolve_shape_on_slide,
)


def op_align(op: ParsedOp, ctx: SlidesOpContext) -> OpResult:
    """Align shapes relative to each other or the slide.

    Syntax: align left|right|center|top|bottom|middle SHAPE [SHAPE...]
    """
    if len(op.positionals) < 2:
        return OpResult(success=False, message="Usage: align left|right|center|top|bottom|middle SHAPE [SHAPE...]")

    result = require_active_slide(ctx)
    if isinstance(result, str):
        return OpResult(success=False, message=result)
    slide, slide_idx = result

    direction = op.positionals[0].lower()
    shape_refs = op.positionals[1:]

    shapes = []
    for ref in shape_refs:
        shape = resolve_shape_on_slide(ref, slide, slide_idx, ctx)
        if shape is None:
            return OpResult(success=False, message=f"Shape not found: {ref!r}")
        shapes.append(shape)

    if not shapes:
        return OpResult(success=False, message="No shapes to align")

    valid_directions = {"left", "right", "center", "top", "bottom", "middle"}
    if direction not in valid_directions:
        return OpResult(success=False, message=f"Unknown direction: {direction!r}. Use: {', '.join(sorted(valid_directions))}")

    if len(shapes) == 1:
        # Align to slide
        slide_width = ctx.prs.slide_width or Emu(9144000)
        slide_height = ctx.prs.slide_height or Emu(6858000)
        shape = shapes[0]

        if direction == "left":
            shape.left = 0
        elif direction == "right":
            shape.left = slide_width - (shape.width or 0)
        elif direction == "center":
            shape.left = (slide_width - (shape.width or 0)) // 2
        elif direction == "top":
            shape.top = 0
        elif direction == "bottom":
            shape.top = slide_height - (shape.height or 0)
        elif direction == "middle":
            shape.top = (slide_height - (shape.height or 0)) // 2
    else:
        # Align to each other (use first shape as reference)
        ref_shape = shapes[0]

        for shape in shapes[1:]:
            if direction == "left":
                shape.left = ref_shape.left
            elif direction == "right":
                shape.left = (ref_shape.left or 0) + (ref_shape.width or 0) - (shape.width or 0)
            elif direction == "center":
                ref_center = (ref_shape.left or 0) + (ref_shape.width or 0) // 2
                shape.left = ref_center - (shape.width or 0) // 2
            elif direction == "top":
                shape.top = ref_shape.top
            elif direction == "bottom":
                shape.top = (ref_shape.top or 0) + (ref_shape.height or 0) - (shape.height or 0)
            elif direction == "middle":
                ref_middle = (ref_shape.top or 0) + (ref_shape.height or 0) // 2
                shape.top = ref_middle - (shape.height or 0) // 2

    return OpResult(
        success=True,
        message=f"Aligned {len(shapes)} shape(s) {direction}",
        prefix="*",
    )


def op_distribute(op: ParsedOp, ctx: SlidesOpContext) -> OpResult:
    """Distribute shapes evenly.

    Syntax: distribute horizontal|vertical SHAPE SHAPE [SHAPE...]
    """
    if len(op.positionals) < 3:
        return OpResult(success=False, message="Usage: distribute horizontal|vertical SHAPE SHAPE [SHAPE...]")

    result = require_active_slide(ctx)
    if isinstance(result, str):
        return OpResult(success=False, message=result)
    slide, slide_idx = result

    axis = op.positionals[0].lower()
    shape_refs = op.positionals[1:]

    if axis not in ("horizontal", "vertical"):
        return OpResult(success=False, message=f"Unknown axis: {axis!r}. Use: horizontal, vertical")

    shapes = []
    for ref in shape_refs:
        shape = resolve_shape_on_slide(ref, slide, slide_idx, ctx)
        if shape is None:
            return OpResult(success=False, message=f"Shape not found: {ref!r}")
        shapes.append(shape)

    if len(shapes) < 2:
        return OpResult(success=False, message="Need at least 2 shapes to distribute")

    if axis == "horizontal":
        # Sort by left position
        shapes.sort(key=lambda s: s.left or 0)
        first_left = shapes[0].left or 0
        last_right = (shapes[-1].left or 0) + (shapes[-1].width or 0)
        total_width = sum(s.width or 0 for s in shapes)
        gap = (last_right - first_left - total_width) / max(1, len(shapes) - 1)

        current_left = first_left
        for shape in shapes:
            shape.left = int(current_left)
            current_left += (shape.width or 0) + gap
    else:
        # Sort by top position
        shapes.sort(key=lambda s: s.top or 0)
        first_top = shapes[0].top or 0
        last_bottom = (shapes[-1].top or 0) + (shapes[-1].height or 0)
        total_height = sum(s.height or 0 for s in shapes)
        gap = (last_bottom - first_top - total_height) / max(1, len(shapes) - 1)

        current_top = first_top
        for shape in shapes:
            shape.top = int(current_top)
            current_top += (shape.height or 0) + gap

    return OpResult(
        success=True,
        message=f"Distributed {len(shapes)} shapes {axis}",
        prefix="*",
    )


def op_z_order(op: ParsedOp, ctx: SlidesOpContext) -> OpResult:
    """Change the z-order of a shape.

    Syntax: z-order front|back|forward|backward SHAPE
    """
    if len(op.positionals) < 2:
        return OpResult(success=False, message="Usage: z-order front|back|forward|backward SHAPE")

    result = require_active_slide(ctx)
    if isinstance(result, str):
        return OpResult(success=False, message=result)
    slide, slide_idx = result

    position = op.positionals[0].lower()
    shape_ref = op.positionals[1]

    if position not in ("front", "back", "forward", "backward"):
        return OpResult(success=False, message=f"Unknown z-order: {position!r}. Use: front, back, forward, backward")

    shape = resolve_shape_on_slide(shape_ref, slide, slide_idx, ctx)
    if shape is None:
        return OpResult(success=False, message=f"Shape not found: {shape_ref!r}")

    # Manipulate z-order via XML
    sp = shape._element
    parent = sp.getparent()
    siblings = list(parent)

    if position == "front":
        parent.remove(sp)
        parent.append(sp)
    elif position == "back":
        parent.remove(sp)
        parent.insert(0, sp)
    elif position == "forward":
        idx = siblings.index(sp)
        if idx < len(siblings) - 1:
            parent.remove(sp)
            parent.insert(idx + 1, sp)
    elif position == "backward":
        idx = siblings.index(sp)
        if idx > 0:
            parent.remove(sp)
            parent.insert(idx - 1, sp)

    return OpResult(success=True, message=f"Shape '{shape_ref}' moved {position}", prefix="*")


HANDLERS: dict[str, callable] = {
    "align": op_align,
    "distribute": op_distribute,
    "z-order": op_z_order,
}
