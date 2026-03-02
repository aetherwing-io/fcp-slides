# fcp-slides

## Project Overview
MCP server that lets LLMs create and edit PowerPoint presentations through a semantic verb DSL.
Uses python-pptx as the native library (Tier 2 architecture).

## Architecture
- `src/fcp_slides/model/` — Thin wrapper around python-pptx Presentation, slide/shape index, shape refs
- `src/fcp_slides/server/` — Verb handlers (ops_*.py), queries, verb registry, resolvers
- `src/fcp_slides/lib/` — Unit conversion, color palette, shape types, layout names, chart types
- `src/fcp_slides/adapter.py` — FcpDomainAdapter bridging fcp-core to python-pptx
- `src/fcp_slides/main.py` — Server entry point

## Key Patterns
- Each `ops_*.py` exports a `HANDLERS` dict mapping verb names to handler functions
- The adapter merges all HANDLERS at import time for dispatch
- `queries.py` exports `QUERY_HANDLERS` for query dispatch
- Undo/redo: byte snapshots via `prs.save(BytesIO)` / `Presentation(BytesIO)`
- Batch atomicity: pre-batch snapshot, rollback on any op failure
- Positions/sizes use EMU internally, lib/units.py converts from `2in`/`5cm`
- Shapes and slides referenced by labels (auto-assigned or user-specified)

## python-pptx Gotchas
- Slide removal requires XML manipulation (`_sldIdLst`)
- Slide copy is partial — uses XML deepcopy
- Slide hide/unhide via XML attribute (`show="0"`)
- Chart data is set at creation; modifications use `chart.replace_data()`
- Positions use EMUs (914400/inch)

## Commands
- `uv run pytest` — Run tests (70 tests)
- `uv run python -c "from fcp_slides.main import main"` — Verify import

## Verb Categories
| Category | Verbs |
|----------|-------|
| Slides | slide (add/remove/rename/copy/move/hide/unhide/activate) |
| Shapes | shape (add/remove/move/resize/duplicate), textbox, connector |
| Text | text (set/append/clear), placeholder (set), bullet |
| Tables | table (add/set/style/row/header/merge/remove) |
| Charts | chart (add/data/series/axis/remove) |
| Images | image (add/placeholder/remove) |
| Layout | align, distribute, z-order |
| Style | style (fill/outline/shadow/rotation), text-style (font/size/color/bold) |
| Meta | notes (set/append/clear), deck (widescreen/standard/size) |
