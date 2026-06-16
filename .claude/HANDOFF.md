# Session Handoff

**Branch**: master
**When**: 2026-06-12 20:30:35 UTC
**Session**: c8a00a23-ebe4-4b0e-93bc-efed50f345f1

## Summary
Already correct. Investigated, verified:

**Server expose only 5 code-mode tools by default. Old tools registered internal but disabled — by design.**

How it work:

1. **Default mode `codemode_only`** — `_resolve_tool_mode()` at `tools.py:130` reads `MICROSOFT_MCP_TOOL_MODE`, defaults `codemode_only`. `hybrid` = opt-in to expose everything.
2. **Gate at `tools.py:5494`** — `_configure_public_tool_mode()` runs as last line of module, after all registration. Disables every tool except 5 in `CODE_MODE_TOOL_NAMES` (`tools.py:120`): `search_tools`, `list_tools`, `tools_info`, `get_required_keys_for_tool`, `call_tool_chain`.
3. **Disabled ≠ deleted.** Business tools stay in FastMCP registry.

## Modified Files
- CLAUDE.md
- README.md
- env.example
- plans/ADD_CODE_MODE_PLAN.md
- plans/code-mode-revamp/code-mode-mcp-server.md
- pyproject.toml
- src/microsoft_mcp/auth_msal.py
- src/microsoft_mcp/inbox_models.py
- src/microsoft_mcp/inbox_ranking.py
- src/microsoft_mcp/response_shaping.py
- src/microsoft_mcp/server.py
- src/microsoft_mcp/tools.py
- src/microsoft_mcp/utcp_bridge_config.py
- tests/test_auth_msal.py
- tests/test_code_mode_tools.py
- tests/test_draft_tools.py
- tests/test_graph_401_retry.py
- tests/test_inbox_ranking.py
- tests/test_refresh_all_accounts.py
- tests/test_tool_surface_contract.py
