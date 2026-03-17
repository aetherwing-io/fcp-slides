# fcp-slides

MCP server for semantic presentation operations.

## What It Does

fcp-slides lets LLMs create and edit PowerPoint presentations by describing slide intent -- layouts, shapes, tables, charts, text styling -- and renders it into standard `.pptx` files. Instead of writing python-pptx code, the LLM works with operations like `slide add layout:title`, `placeholder set title "Q4 Report"`, and `table add 5 4 label:metrics`. Built on the [FCP](https://github.com/os-tack/fcp) framework, powered by python-pptx for serialization.

## Quick Example

```
slides_session('new "Q4 Review"')

slides([
  'slide add layout:title',
  'placeholder set title "Q4 Business Review"',
  'placeholder set subtitle "Prepared by Finance Team"',
  'slide add layout:blank',
  'table add 5 4 label:metrics x:1in y:1.5in w:8in h:4in',
  'table header metrics "Region" "Revenue" "Costs" "Margin"',
  'table set metrics 2 1 "North"',
  'table set metrics 2 2 "$1.25M"',
  'chart add bar label:revenue title:"Revenue by Region" x:1in y:1.5in w:8in h:4.5in',
  'notes set "Discuss regional performance trends"',
])

slides_session('save as:./q4_review.pptx')
```

### Available MCP Tools

| Tool | Purpose |
|------|---------|
| `slides(ops)` | Batch mutations -- slides, shapes, text, tables, charts, images, styling |
| `slides_query(q)` | Inspect the presentation -- slides, shapes, placeholders, layout info |
| `slides_session(action)` | Lifecycle -- new, open, save, checkpoint, undo, redo |
| `slides_help()` | Full reference card |

## Installation

Requires Python >= 3.11.

```bash
pip install fcp-slides
```

### MCP Client Configuration

```json
{
  "mcpServers": {
    "slides": {
      "command": "uv",
      "args": ["run", "python", "-m", "fcp_slides"]
    }
  }
}
```

## Architecture

3-layer architecture:

```
MCP Server (Intent Layer)
  Parses op strings, dispatches to verb handlers
        |
Semantic Model
  Thin wrapper around python-pptx Presentation
  Slide/shape index, shape refs, undo/redo via byte snapshots
        |
Serialization (python-pptx)
  Semantic model -> .pptx binary output
```

Key features:

- **Slide layouts** -- Title, blank, content, section header, and custom layouts
- **Shapes** -- Rectangles, ellipses, arrows, textboxes, connectors
- **Tables** -- Create, populate, style with headers and merged cells
- **Charts** -- Bar, line, pie, scatter, area, doughnut, and more
- **Images** -- Add from file or set as slide background
- **Text styling** -- Font, size, color, bold, italic, alignment, spacing
- **Layout ops** -- Align, distribute, z-order for shape arrangement
- **Undo/redo** -- Full presentation snapshots with event sourcing
- **Keynote-compatible** -- Standard PPTX output opens in Keynote, Google Slides, LibreOffice

## Development

```bash
uv sync
uv run pytest       # 70 tests
uv run ruff check   # linting
uv run pyright      # type checking
```

## License

MIT
