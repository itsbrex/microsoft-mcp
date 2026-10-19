---
description: Run the MCP server locally (MSAL) for a smoke test
allowed-tools: ["Bash(uv run:*)", "Bash(timeout:*)"]
argument-hint: "[account-email]"
---

# Run MCP server (smoke test)

Start the server with MSAL auth to verify it boots and registers tools. Logging goes to stderr (stdout is reserved for JSON-RPC).

Account: `$ARGUMENTS` (defaults to `MICROSOFT_MCP_ACCOUNT_ID` or "default").

## Steps

1. Boot with a short timeout so it does not hang the session:
   ```bash
   MICROSOFT_MCP_AUTH_METHOD=msal \
   MICROSOFT_MCP_CLIENT_ID=d3590ed6-52b3-4102-aeff-aad2292ab01c \
   MICROSOFT_MCP_ACCOUNT_ID="$ARGUMENTS" \
   MICROSOFT_MCP_REFRESH_ON_STARTUP=0 \
   timeout 10 uv run microsoft-mcp 2>&1 | tail -30
   ```
2. Confirm startup log lines show auth init + tool registration, no tracebacks.
3. `timeout` exiting 124 after clean startup logs = pass.
