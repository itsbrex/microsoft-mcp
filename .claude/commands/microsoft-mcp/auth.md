---
description: Run the authentication flow (Azure SDK or MSAL device-code)
allowed-tools: ["Bash(uv run:*)"]
---

Authenticate the active Microsoft account. Flow depends on `MICROSOFT_MCP_AUTH_METHOD`:

- Unset / `azure` → `InteractiveBrowserCredential` (opens a browser)
- `msal` → device-code flow (copy/paste code into the browser)

Required env: `MICROSOFT_MCP_CLIENT_ID`. For MSAL also set `MICROSOFT_MCP_ACCOUNT_ID` so tokens land in `~/.config/microsoft-mcp/tokens/<account>_access_token.json`.

!`uv run authenticate.py`
