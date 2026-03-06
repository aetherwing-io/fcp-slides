"""fcp-slides — Presentation File Context Protocol MCP server."""

from __future__ import annotations

from fcp_core.server import create_fcp_server

from fcp_slides.adapter import SlidesAdapter
from fcp_slides.server.reference_card import EXTRA_SECTIONS
from fcp_slides.server.verb_registry import VERBS

adapter = SlidesAdapter()

mcp = create_fcp_server(
    domain="slides",
    adapter=adapter,
    verbs=VERBS,
    extra_sections=EXTRA_SECTIONS,
    extensions=["pptx", "ppt", "odp"],
    name="fcp-slides",
    instructions="FCP Slides server for creating and editing presentation files (pptx). Use slides_session to create a new presentation or open an existing file, slides to add slides, shapes, text, and images, slides_query to inspect slide contents and layout, and slides_help for the full verb reference. Start every interaction with slides_session.",
)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
