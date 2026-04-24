---
description: Weekly tool-surface audit pass — bugs, response_profile drift, cache integrity, ranker signals
---

Run a periodic regression scan of the microsoft-mcp tool surface against the invariants set by the 2026-04-23 audit (`docs/superpowers/plans/2026-04-23-tool-surface-audit.md`).

## Step 1: Run the regression suite

```
uv run pytest tests/test_tool_surface_contract.py -v
```

All tests must pass. If any fail, investigate the FAILED test's docstring — it ties back to a specific audit ID (B1, B3, A1, etc.) so you can map the regression to the original finding.

## Step 2: Run `/techdebt src/microsoft_mcp/`

This invokes the existing techdebt skill across the production module. Flag:

1. New list/search tools without `response_profile`.
2. New raw httpx calls (anywhere outside `graph.py`).
3. New silent `except Exception: pass` blocks.
4. Tool functions with public names (no leading underscore) but lacking `@mcp.tool`.
5. Any `InboxItem(...)` construction that skips `mentioned`/`flagged`/`is_newsletter`/`starts_in_minutes`.
6. New `raise Exception(...)` calls in auth modules without `from e`.

## Step 3: Spot-check the inbox ranker live

If you have a working MSAL session:

```
uv run python -c "from microsoft_mcp import tools; print(tools.list_inbox_items.fn(limit=5))"
```

Sanity: items should NOT all score `10.0` (that's the unread-only fingerprint of a regressed ranker — see B3).

## Step 4: Report

Report `pass/fail` against each of these five categories. Suggest fixes; do not apply them automatically. If you find a regression, open a fix branch off `master` rather than patching ad-hoc.
