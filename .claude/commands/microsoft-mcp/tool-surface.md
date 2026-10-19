---
description: Verify the registered MCP tool surface (guards the 109-tool inventory)
allowed-tools: ["Bash(uv run:*)"]
---

Run the tool-surface regression guard — verifies every mail-port tool stays registered. Update `tests/test_tool_surface.py` **and** the tool inventory in `CLAUDE.md` when adding or removing tools.

!`uv run pytest tests/test_tool_surface.py -q`
