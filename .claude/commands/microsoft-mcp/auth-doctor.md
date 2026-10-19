---
description: Diagnose MSAL auth health — perms, duplicate stores, expiry (no network)
allowed-tools: ["Bash(uv run:*)"]
argument-hint: "[--json]"
---

Diagnose MSAL token storage health: file permissions, duplicate account stores, expiry drift. Read-only, no network calls.

If no accounts are found in the default store (`~/.config/microsoft-mcp/tokens/`), re-run with the outlook-creds compat store:
`MICROSOFT_MCP_TOKENS_DIR=~/.config/outlook-creds/tokens`

Arguments passed: `$ARGUMENTS`

!`uv run microsoft-mcp auth doctor $ARGUMENTS`
