# Code Mode Orchestration for Inbox Triage

This document explains when and how to use Code Mode with this MCP server to orchestrate
multi-step inbox triage workflows efficiently.

## Core Principle

The server handles **payload shaping** — trimming raw Microsoft Graph responses so they fit
within token budgets. Code Mode handles **orchestration** — batching follow-up calls over
only the items that actually need full hydration.

These are separate concerns. Do not conflate them.

## When to Use Code Mode

Use Code Mode when you need to:

- Fetch a list of summaries, then hydrate only the selected subset
- Apply ranking or filtering logic that the server does not expose as a parameter
- Batch multiple `get_inbox_item_detail` calls and reduce them to a single compact report
- Chain `list_inbox_items` → `unified_search`/`search_emails` → `get_inbox_item_detail`
  without passing intermediate full payloads to the model

Do **not** use Code Mode as a workaround for raw Graph payload size. The server already applies
`response_shaping.py` to every response. If token usage is still too high, check `BudgetHints`
parameters on the individual tools, not Code Mode.

## Recommended Inbox Triage Flow

```
list_inbox_items(limit=20)
    -> ranked InboxItem summaries (id, kind, title, snippet, score, reason)

[optional] unified_search(query="...") or search_emails(query="...")
    -> narrowed summaries for keyword/sender searches

get_inbox_item_detail(item_id=..., kind=...)   [call for top 2-3 items only]
    -> full body, participants, action_hints

[Code Mode] compute and return triage report
```

### Step 1: list_inbox_items

Returns normalized `InboxItem` summaries ranked by urgency. Fields present on every item:

| Field | Type | Description |
|---|---|---|
| `id` | str | Opaque item identifier (pass to `get_inbox_item_detail`) |
| `kind` | str | `"email"` or `"event"` |
| `source_tool` | str | Which Graph API tool produced this item |
| `title` | str | Subject line or event title |
| `snippet` | str | Short preview of the body |
| `participants` | list[str] | Sender or organizer + key recipients |
| `when` | str or absent | ISO timestamp (absent if unknown) |
| `state` | str | `"unread"`, `"read"`, `"flagged"`, etc. |
| `score` | float | Urgency score (higher = more urgent) |
| `reason` | str | Human-readable explanation of score |
| `action_hints` | list[str] | Suggested next actions |
| `web_url` | str | Deep link to item in Outlook/Teams |

### Step 2: unified_search or search_emails (optional)

Use `unified_search` or `search_emails` when the user has a specific keyword, sender, or subject
to narrow results before hydrating. Avoid calling `get_inbox_item_detail` on items that have not
passed a relevance filter.

### Step 3: get_inbox_item_detail

Only call this for items the triage logic has selected. Hydrating 10+ items is expensive.
The typical pattern is top 3 by score, unless the user specifies otherwise.

### Step 4: Compile the report in Code Mode

Code Mode receives the detail payloads and reduces them to a compact triage report that fits
in a single assistant message. The report should include:

- Item title and kind
- One-sentence summary of what needs to happen
- Suggested action (reply, accept, delegate, archive)
- Any deadlines or time-sensitive signals

## MCP Server Registration

To use this server from a Code Mode script, register it in your MCP configuration:

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "/path/to/uv",
      "args": [
        "run", "--python", "3.13",
        "--project", "/path/to/microsoft-mcp",
        "microsoft-mcp"
      ],
      "env": {
        "MICROSOFT_MCP_AUTH_METHOD": "msal",
        "MICROSOFT_MCP_ACCOUNT_ID": "your-email@example.com",
        "MICROSOFT_MCP_CLIENT_ID": "d3590ed6-52b3-4102-aeff-aad2292ab01c"
      }
    }
  }
}
```

## What Code Mode Is Not

- It is **not** a substitute for server-side response shaping. If raw Graph payloads are too
  large, configure `BudgetHints` or `ResponseProfile` parameters on the tool call, not in
  Code Mode.
- It is **not** required for simple single-item lookups. Use `get_inbox_item_detail` directly
  when you already know the item ID.
- It is **not** required for single-step listing. `list_inbox_items` is already optimized for
  direct assistant use.

## Full Example

See [`examples/code-mode/inbox_triage.ts`](../examples/code-mode/inbox_triage.ts) for a
complete TypeScript script that:

1. Registers this MCP server
2. Calls `list_inbox_items` for ranked summaries
3. Optionally narrows with `unified_search` or `search_emails`
4. Hydrates the top 3 items with `get_inbox_item_detail`
5. Returns a compact triage report
