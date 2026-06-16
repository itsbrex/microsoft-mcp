---
description: Refresh MSAL tokens (all accounts, single account, or force re-auth)
allowed-tools: ["Bash(uv run:*)"]
argument-hint: "[email] [--force]"
---

Refresh saved MSAL access tokens. Arguments determine scope:

- no arg → refresh all saved accounts (mirrors `refresh_all_accounts()`)
- `<email>` → refresh only that account
- `--force <email>` → clear the saved tokens for `<email>` and re-trigger the MSAL device-code flow

Required env: `MICROSOFT_MCP_AUTH_METHOD=msal`, `MICROSOFT_MCP_CLIENT_ID`. The device-code flow blocks until the user completes the browser flow.

Arguments passed: `$ARGUMENTS`

!`uv run auth_refresh.py $ARGUMENTS`
