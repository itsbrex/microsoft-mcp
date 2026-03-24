"""
Inbox triage example for the integrated Microsoft MCP code-mode surface.

This example assumes a client object named ``mcp`` that exposes the public
code-mode APIs:

- search_tools(query)
- list_tools()
- tools_info(tool_names)
- get_required_keys_for_tool(tool_name)
- call_tool_chain(code, timeout?)

The example is intentionally small and focused on the integrated workflow:
discovery first, interface inspection second, multi-step execution last.
"""

from __future__ import annotations

import asyncio
from typing import Any


async def run_inbox_triage(mcp: Any) -> dict[str, Any]:
    """Fetch ranked inbox summaries, hydrate only the top items, and return a compact report."""

    discovery = await mcp.search_tools("inbox triage selective hydration")
    active_tools = await mcp.list_tools()
    tool_info = await mcp.tools_info(
        ["list_inbox_items", "get_inbox_item_detail", "search_emails"]
    )
    required_keys = await mcp.get_required_keys_for_tool("list_inbox_items")

    report = await mcp.call_tool_chain(
        """
summary = microsoft.list_inbox_items({"limit": 20})
top_items = summary["items"][:3]
details = [
    microsoft.get_inbox_item_detail({"item_id": item["id"], "kind": item["kind"]})
    for item in top_items
]

return {
    "titles": [item["title"] for item in top_items],
    "actions": [
        detail["action_hints"][0] if detail["action_hints"] else "review"
        for detail in details
    ],
    "scores": [item["score"] for item in top_items],
}
""",
        timeout=30,
    )

    return {
        "discovery": discovery,
        "active_tools": active_tools,
        "tool_info": tool_info,
        "required_keys": required_keys,
        "report": report,
    }


async def _demo() -> None:
    raise RuntimeError(
        "Replace the placeholder client with the integrated Microsoft MCP code-mode client."
    )


if __name__ == "__main__":
    asyncio.run(_demo())
