"""Verb registry for fcp-slides — defines all verb specs."""

from __future__ import annotations

from fcp_core import VerbSpec

VERBS: list[VerbSpec] = [
    # -- Slides --
    VerbSpec(
        verb="slide",
        syntax="slide add|remove|rename|copy|move|hide|unhide|activate [NAME] [layout:LAYOUT] [at:N] [label:NAME]",
        category="slides",
        params=["layout", "at", "label", "after", "before"],
        description="Manage slides: add, remove, rename, copy, move, hide/unhide, or activate.",
    ),
    # -- Shapes --
    VerbSpec(
        verb="shape",
        syntax="shape add|remove|move|resize|duplicate TYPE [label:NAME] [x:POS] [y:POS] [w:SIZE] [h:SIZE]",
        category="shapes",
        params=["label", "x", "y", "w", "h", "cx", "cy", "width", "height", "on"],
        description="Add, remove, move, resize, or duplicate shapes on the active slide.",
    ),
    VerbSpec(
        verb="textbox",
        syntax='textbox TEXT [x:POS] [y:POS] [w:SIZE] [h:SIZE] [label:NAME]',
        category="shapes",
        params=["x", "y", "w", "h", "label", "font", "size", "color", "bold", "align"],
        description="Add a textbox with content (convenience shorthand for shape add + text set).",
    ),
    VerbSpec(
        verb="connector",
        syntax="connector FROM TO [type:straight|elbow|curved]",
        category="shapes",
        params=["type"],
        description="Connect two shapes with a connector line.",
    ),
    # -- Text --
    VerbSpec(
        verb="text",
        syntax="text set|append|clear SHAPE TEXT",
        category="text",
        params=["on"],
        description="Set, append to, or clear text content of a shape.",
    ),
    VerbSpec(
        verb="placeholder",
        syntax='placeholder set title|subtitle|body TEXT [on:SLIDE]',
        category="text",
        params=["on"],
        description="Set text in a slide placeholder by type name.",
    ),
    VerbSpec(
        verb="bullet",
        syntax="bullet SHAPE TEXT [level:N]",
        category="text",
        params=["level", "on"],
        description="Add a bullet item to a shape at a given indent level.",
    ),
    # -- Tables --
    VerbSpec(
        verb="table",
        syntax="table add|set|style|row|header|merge|remove ROWS COLS [label:NAME] [x:POS] [y:POS] [w:SIZE] [h:SIZE]",
        category="tables",
        params=["label", "x", "y", "w", "h"],
        description="Create and manipulate tables on the active slide.",
    ),
    # -- Charts --
    VerbSpec(
        verb="chart",
        syntax='chart add|data|series|axis|remove TYPE [label:NAME] [title:TITLE] [x:POS] [y:POS] [w:SIZE] [h:SIZE]',
        category="charts",
        params=["label", "title", "x", "y", "w", "h", "categories", "series", "values"],
        description="Create and configure charts on the active slide.",
    ),
    # -- Images --
    VerbSpec(
        verb="image",
        syntax="image add|placeholder|remove PATH [x:POS] [y:POS] [w:SIZE] [h:SIZE] [label:NAME]",
        category="images",
        params=["x", "y", "w", "h", "label", "on"],
        description="Add, set placeholder, or remove images.",
    ),
    # -- Layout --
    VerbSpec(
        verb="align",
        syntax="align left|right|center|top|bottom|middle SHAPE [SHAPE...]",
        category="layout",
        params=[],
        description="Align shapes relative to each other or the slide.",
    ),
    VerbSpec(
        verb="distribute",
        syntax="distribute horizontal|vertical SHAPE [SHAPE...]",
        category="layout",
        params=[],
        description="Distribute shapes evenly across horizontal or vertical axis.",
    ),
    VerbSpec(
        verb="z-order",
        syntax="z-order front|back|forward|backward SHAPE",
        category="layout",
        params=[],
        description="Change the z-order (layering) of a shape.",
    ),
    # -- Style --
    VerbSpec(
        verb="style",
        syntax="style SHAPE [fill:#HEX] [outline:#HEX] [outline-width:PT] [shadow] [opacity:N] [rotation:N]",
        category="style",
        params=["fill", "outline", "outline-width", "shadow", "opacity", "rotation"],
        description="Apply visual styling to a shape (fill, outline, shadow, opacity, rotation).",
    ),
    VerbSpec(
        verb="text-style",
        syntax="text-style SHAPE [font:NAME] [size:N] [color:#HEX] [bold] [italic] [underline] [align:left|center|right] [spacing:N]",
        category="style",
        params=["font", "size", "color", "align", "valign", "spacing", "line-spacing"],
        description="Apply text formatting to all text in a shape.",
    ),
    # -- Meta --
    VerbSpec(
        verb="notes",
        syntax="notes set|append|clear TEXT [on:SLIDE]",
        category="meta",
        params=["on"],
        description="Set, append to, or clear speaker notes for a slide.",
    ),
    VerbSpec(
        verb="deck",
        syntax="deck size|widescreen|standard [w:WIDTH] [h:HEIGHT]",
        category="meta",
        params=["w", "h"],
        description="Set presentation-level properties (slide size).",
    ),
]

VERB_MAP: dict[str, VerbSpec] = {v.verb: v for v in VERBS}
