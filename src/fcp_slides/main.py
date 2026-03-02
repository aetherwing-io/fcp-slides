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
    name="slides-fcp",
    instructions="Presentation File Context Protocol. Call slides_help for the reference card.",
)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
