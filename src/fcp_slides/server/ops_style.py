"""Style operation handlers — style (shape visual), text-style (font formatting)."""

from __future__ import annotations

from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from fcp_core import OpResult, ParsedOp

from fcp_slides.lib.colors import parse_color, to_rgb
from fcp_slides.lib.units import parse_length
from fcp_slides.server.resolvers import (
    SlidesOpContext,
    require_active_slide,
    resolve_shape_on_slide,
)


def op_style(op: ParsedOp, ctx: SlidesOpContext) -> OpResult:
    """Apply visual styling to a shape.

    Syntax: style SHAPE [fill:#HEX] [outline:#HEX] [outline-width:PT]
            [shadow] [opacity:N] [rotation:N]
    """
    if not op.positionals:
        return OpResult(success=False, message="Usage: style SHAPE [fill:#HEX] [outline:#HEX] [shadow] [rotation:N]")

    result = require_active_slide(ctx)
    if isinstance(result, str):
        return OpResult(success=False, message=result)
    slide, slide_idx = result

    shape_ref = op.positionals[0]
    shape = resolve_shape_on_slide(shape_ref, slide, slide_idx, ctx)
    if shape is None:
        return OpResult(success=False, message=f"Shape not found: {shape_ref!r}")

    params = op.params
    flags = {p.lower() for p in op.positionals[1:]}
    changes: list[str] = []

    # Fill color
    if "fill" in params:
        fill = shape.fill
        fill.solid()
        fill.fore_color.rgb = to_rgb(params["fill"])
        changes.append(f"fill:{params['fill']}")

    # Outline
    if "outline" in params:
        line = shape.line
        line.color.rgb = to_rgb(params["outline"])
        changes.append(f"outline:{params['outline']}")

    if "outline-width" in params:
        line = shape.line
        line.width = Pt(float(params["outline-width"]))
        changes.append(f"outline-width:{params['outline-width']}")

    # Shadow
    if "shadow" in flags or "shadow" in params:
        # python-pptx has limited shadow support — set via XML
        shadow_elem = shape._element.find(
            ".//{http://schemas.openxmlformats.org/drawingml/2006/main}effectLst"
        )
        if shadow_elem is None:
            from lxml import etree
            nsmap = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
            sp_pr = shape._element.find(".//a:spPr", nsmap)
            if sp_pr is not None:
                effect_lst = etree.SubElement(sp_pr, f"{{{nsmap['a']}}}effectLst")
                outer_shdw = etree.SubElement(effect_lst, f"{{{nsmap['a']}}}outerShdw")
                outer_shdw.set("blurRad", "50800")
                outer_shdw.set("dist", "38100")
                outer_shdw.set("dir", "2700000")
                srgb = etree.SubElement(outer_shdw, f"{{{nsmap['a']}}}srgbClr")
                srgb.set("val", "000000")
                alpha = etree.SubElement(srgb, f"{{{nsmap['a']}}}alpha")
                alpha.set("val", "40000")
        changes.append("shadow")

    # Rotation
    if "rotation" in params:
        try:
            shape.rotation = float(params["rotation"])
            changes.append(f"rotation:{params['rotation']}")
        except (ValueError, AttributeError):
            pass

    if not changes:
        return OpResult(success=False, message="No style properties specified")

    return OpResult(
        success=True,
        message=f"Shape '{shape_ref}' styled: {', '.join(changes)}",
        prefix="*",
    )


_ALIGN_MAP = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}


def op_text_style(op: ParsedOp, ctx: SlidesOpContext) -> OpResult:
    """Apply text formatting to all text in a shape.

    Syntax: text-style SHAPE [font:NAME] [size:N] [color:#HEX]
            [bold] [italic] [underline] [align:left|center|right] [spacing:N]
    """
    if not op.positionals:
        return OpResult(success=False, message="Usage: text-style SHAPE [font:NAME] [size:N] [color:#HEX] [bold] [italic]")

    result = require_active_slide(ctx)
    if isinstance(result, str):
        return OpResult(success=False, message=result)
    slide, slide_idx = result

    shape_ref = op.positionals[0]
    shape = resolve_shape_on_slide(shape_ref, slide, slide_idx, ctx)
    if shape is None:
        return OpResult(success=False, message=f"Shape not found: {shape_ref!r}")

    if not shape.has_text_frame:
        return OpResult(success=False, message=f"Shape '{shape_ref}' has no text frame")

    params = op.params
    flags = {p.lower() for p in op.positionals[1:]}
    changes: list[str] = []

    tf = shape.text_frame

    for para in tf.paragraphs:
        # Paragraph-level: alignment
        if "align" in params:
            align = _ALIGN_MAP.get(params["align"].lower())
            if align:
                para.alignment = align

        # Run-level: font properties
        for run in para.runs:
            if "font" in params:
                run.font.name = params["font"]
            if "size" in params:
                run.font.size = Pt(float(params["size"]))
            if "color" in params:
                run.font.color.rgb = to_rgb(params["color"])
            if "bold" in flags:
                run.font.bold = True
            if "italic" in flags:
                run.font.italic = True
            if "underline" in flags:
                run.font.underline = True

        # Line spacing
        if "line-spacing" in params or "spacing" in params:
            spacing = params.get("line-spacing") or params.get("spacing")
            try:
                para.line_spacing = Pt(float(spacing))
            except (ValueError, AttributeError):
                pass

    # Build change summary
    for flag in ("bold", "italic", "underline"):
        if flag in flags:
            changes.append(flag)
    for k in ("font", "size", "color", "align", "spacing", "line-spacing"):
        if k in params:
            changes.append(f"{k}:{params[k]}")

    if not changes:
        return OpResult(success=False, message="No text style properties specified")

    return OpResult(
        success=True,
        message=f"Text styled on '{shape_ref}': {', '.join(changes)}",
        prefix="*",
    )


HANDLERS: dict[str, callable] = {
    "style": op_style,
    "text-style": op_text_style,
}
