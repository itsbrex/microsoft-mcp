---
description: Live auth test — Graph /me call per saved MSAL account
allowed-tools: ["Bash(uv run:*)"]
argument-hint: "[--json]"
---

Test authentication by making a live Microsoft Graph `/me` call for each saved MSAL account. Network calls, but no token refresh and no interactive flows.

If no accounts are found in the default store, prefix with `MICROSOFT_MCP_TOKENS_DIR=~/.config/outlook-creds/tokens`.

Arguments passed: `$ARGUMENTS`

!`uv run microsoft-mcp auth test $ARGUMENTS`
