# Code Mode Orchestration for Inbox Triage

This document explains how to use the integrated code-mode surface in `microsoft-mcp` for multi-step inbox workflows.

By default, `microsoft-mcp` exposes only the code-mode tools publicly. `call_tool_chain` still runs against the internal auth-aware Microsoft business-tool registry.

If you wrap `microsoft-mcp` behind an external UTCP code-mode bridge, set the wrapped server to `MICROSOFT_MCP_TOOL_MODE=hybrid` so the bridge can discover the business tools.

## Core Principle

The server still handles response shaping. The code-mode layer handles orchestration over the live Microsoft tool registry:

- discovery
- interface inspection
- selective hydration
- local ranking or summarization

Do not use code mode to compensate for raw Graph payload size. Use it after the server has already produced compact summaries.

## Public APIs

The integrated surface exposes these operations:

- `search_tools(query)` - find relevant Microsoft tools for a task
- `list_tools()` - list the active, auth-aware tool set
- `tools_info(tool_names)` - return tool metadata and generated interfaces
- `get_required_keys_for_tool(tool_name)` - inspect required config or secrets
- `call_tool_chain(code, timeout?)` - execute sandboxed multi-step code against the active tool set
- `utcp_codemode_usage` - prompt that teaches the discovery-first workflow

## Recommended Inbox Triage Flow

```
search_tools("inbox triage")
    -> identify the smallest useful tool set

list_tools()
    -> confirm which tools are active for the current auth mode

tools_info(["list_inbox_items", "get_inbox_item_detail", "search_emails"])
    -> inspect the generated interfaces

call_tool_chain(...)
    -> fetch summaries, hydrate only the top items, and return a compact report
```

### Step 1: list_inbox_items

Use `list_inbox_items` to get normalized summaries ranked by urgency. Fields present on every item:

| Field | Type | Description |
|---|---|---|
| `id` | str | Opaque item identifier |
| `kind` | str | `email`, `invite_message`, or `event` |
| `source_tool` | str | Which Graph API tool produced the item |
| `title` | str | Subject line or event title |
| `snippet` | str | Short preview of the body |
| `participants` | list[str] | Sender or organizer + key recipients |
| `when` | str or absent | ISO timestamp if known |
| `state` | str | `unread`, `read`, `flagged`, etc. |
| `score` | float | Urgency score |
| `reason` | str | Human-readable explanation of score |
| `action_hints` | list[str] | Suggested next actions |
| `web_url` | str | Deep link to the item |

### Step 2: Optional Narrowing

Use `search_emails`, `list_invite_messages`, or `unified_search` when the user has a specific keyword, sender, subject, or needs to isolate Outlook meeting notices stored in the mailbox.
Use `list_tools` and `tools_info` when you need to verify whether a tool is active or what arguments it expects.

### Step 3: Hydrate Only the Top Items

Only call `get_inbox_item_detail` for the items selected by the triage logic. Hydrating everything defeats the point of code mode.

### Step 4: Compile the Report in Code Mode

The report should be compact and decision-oriented:

- item title and kind
- one-sentence summary
- suggested action
- any deadline or time-sensitive signal

### Step 5: Apply Cleanup Actions in Bulk

Once triage decides what to keep, archive, or tag, use the inbox-management tools instead of iterating manual single-message workflows:

- `bulk_manage_emails` for one action across many messages
- `mark_email_read` for a single message read/unread toggle
- `set_email_categories` for replacing category labels on one message
- `ensure_master_categories` when Outlook needs those labels registered in the mailbox master category list with colors
- `move_email` or `archive_email` for filing a message after review
- `delete_email` for cleanup after confirmation
- `delete_invite_message` for removing meeting notices that should not stay in the inbox
- `rsvp_to_event` for event accepts, tentative holds, or declines without emailing the organizer unless `send_response=True`
- `rsvp_to_invite_message` for responding from the invite message itself while still defaulting to `send_response=False`

## Example

```python
async def triage_inbox(mcp):
    # Discovery first
    matches = await mcp.search_tools("inbox triage and selective hydration")
    print(matches)

    # Inspect the active contract
    info = await mcp.tools_info([
        "list_inbox_items",
        "get_inbox_item_detail",
        "search_emails",
    ])
    print(info)

    result = await mcp.call_tool_chain(
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
}
"""
    )

    return result
```

## Cleanup Example

```python
async def clean_up_after_triage(mcp, ids_to_archive, ids_to_mark_read):
    await mcp.bulk_manage_emails({
        "email_ids": ids_to_archive,
        "action": "archive",
    })

    await mcp.bulk_manage_emails({
        "email_ids": ids_to_mark_read,
        "action": "mark_read",
    })
```

## What Code Mode Is Not

- It is not a substitute for server-side response shaping.
- It is not required for single-item lookups.
- It is not required for single-step listing.

## Related Files

- [`README.md`](/Users/hack/github/microsoft-mcp/README.md)
- [`IMPLEMENTATION.md`](/Users/hack/github/microsoft-mcp/IMPLEMENTATION.md)
- [`examples/code-mode/inbox_triage.py`](/Users/hack/github/microsoft-mcp/examples/code-mode/inbox_triage.py)
