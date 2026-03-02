"""Extra reference card sections for slides help."""

from __future__ import annotations

EXTRA_SECTIONS: dict[str, str] = {
    "Slide References": (
        "Slides can be referenced by:\n"
        "  - 1-based number: 1, 2, 3\n"
        "  - Label: s1, s2 (auto-assigned), or custom label\n"
        "  - Keywords: active, last\n"
    ),
    "Shape References": (
        "Shapes can be referenced by:\n"
        "  - Label: label assigned via label: param or auto-generated\n"
        "  - Shape name: PowerPoint shape name\n"
        "  - 1-based index on the slide\n"
        "\n"
        "Auto-labels: s1_title, s1_subtitle, s1_body (placeholders)\n"
        "             s1_text_box1, s1_rectangle1 (shapes)\n"
    ),
    "Units": (
        "Positions and sizes accept human-friendly units:\n"
        "  2in    — inches (default)\n"
        "  5cm    — centimeters\n"
        "  72pt   — points\n"
        "  100px  — pixels (96 DPI)\n"
        "  914400 — EMUs (raw)\n"
    ),
    "Layouts": (
        "Standard layout names (may vary by template):\n"
        "  title          — Title Slide\n"
        "  title-content  — Title and Content\n"
        "  section        — Section Header\n"
        "  two-content    — Two Content\n"
        "  comparison     — Comparison\n"
        "  title-only     — Title Only\n"
        "  blank          — Blank\n"
    ),
    "Shape Types": (
        "Common shape types for 'shape add TYPE':\n"
        "  rectangle, rounded-rect, oval, diamond, triangle\n"
        "  arrow-right, arrow-left, arrow-up, arrow-down\n"
        "  star-5, star-4, chevron, cross, heart, cloud\n"
        "  callout-rect, callout-rounded, callout-oval\n"
        "  flowchart-process, flowchart-decision, flowchart-terminator\n"
    ),
    "Chart Types": (
        "Chart types for 'chart add TYPE':\n"
        "  column, column-stacked, column-100\n"
        "  bar, bar-stacked, bar-100\n"
        "  line, line-markers, line-stacked\n"
        "  pie, pie-exploded, doughnut\n"
        "  area, area-stacked\n"
        "  scatter, scatter-lines\n"
        "  radar, radar-filled\n"
    ),
    "Colors": (
        "Named colors: blue, orange, gray, gold, lt-blue, green, red,\n"
        "  dk-green, white, black, yellow, purple, teal, dk-blue, dk-red,\n"
        "  lt-gray, dk-gray\n"
        "Hex: #4472C4, 4472C4, F0F (3-char shorthand)\n"
    ),
    "Placeholders": (
        "Standard placeholder names for 'placeholder set TYPE':\n"
        "  title     — slide title\n"
        "  subtitle  — slide subtitle (title slide only)\n"
        "  body      — main content area\n"
    ),
    "Response Prefixes": (
        "+ created  * modified  ~ connection  - removed  ! error/meta  @ bulk\n"
    ),
    "Example Workflow": (
        '  slides_session(\'new "Q4 Report"\')\n'
        "  slides([\n"
        "    'slide add layout:title',\n"
        '    \'placeholder set title "Q4 Revenue Report"\',\n'
        '    \'placeholder set subtitle "Prepared for Board"\',\n'
        "  ])\n"
        "  slides([\n"
        "    'slide add layout:blank',\n"
        '    \'textbox "Key Metrics" x:1in y:0.5in w:8in h:1in\',\n'
        "    'table add 4 3 label:metrics x:1in y:2in w:8in h:4in',\n"
        '    \'table header metrics "Metric" "Q3" "Q4"\',\n'
        '    \'table row metrics 1 "Revenue" "$1.3M" "$1.8M"\',\n'
        "  ])\n"
        "  slides_session('save as:./q4_report.pptx')\n"
    ),
}
