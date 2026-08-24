---
description: Read-only MSAL token health (no network) — per-account validity + expiry
allowed-tools: ["Bash(uv run:*)"]
argument-hint: "[--json]"
---

Show read-only token health for all saved MSAL accounts (does not refresh, makes no network calls). Decodes each saved token's expiry and reports per-account validity.

Arguments passed: `$ARGUMENTS`

!`uv run microsoft-mcp auth status $ARGUMENTS`
