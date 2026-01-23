# Device Code Flow Authentication

## Overview

The email is **not sent** to initiate device code flow. The user's identity is determined by which Microsoft account they sign into at the device login page.

## Device Code Flow Request

From `src/microsoft_mcp/auth_msal.py:364-367`:

```python
flow = app.initiate_device_flow(scopes=DEFAULT_SCOPES)
```

Only the **scopes** are sent. The request to Microsoft's `/devicecode` endpoint contains:

| Parameter | Value | Description |
|-----------|-------|-------------|
| `client_id` | `d3590ed6-52b3-4102-aeff-aad2292ab01c` | Microsoft Office client ID |
| `scope` | `https://graph.microsoft.com/.default offline_access` | Permissions requested |

That's it. No email, no username.

## Where the Email Comes From

The email is determined **by which Microsoft account the user signs into** at `microsoft.com/devicelogin`. Microsoft returns it in the token response:

```python
# Line 424-428 in auth_msal.py
id_token_claims = result.get("id_token_claims", {})
email = id_token_claims.get("preferred_username") or id_token_claims.get("email")
```

## What `MICROSOFT_MCP_ACCOUNT_ID` Does

The `account_identifier` (from `MICROSOFT_MCP_ACCOUNT_ID` env var) is only used for **naming the token files**:

```
~/.config/microsoft-mcp/tokens/
├── user@example.com_access_token.json
├── user@example.com_refresh_only.txt
└── user@example.com_access_only.txt
```

It's a local identifier to support multiple accounts - it doesn't affect the authentication request itself.

## Summary

| Parameter | Sent to Microsoft? | Purpose |
|-----------|-------------------|---------|
| `client_id` | Yes | Identifies the app |
| `scope` | Yes | Permissions requested |
| Email/account_id | **No** | Only for local token file naming |

The user chooses which account to authenticate by signing in at the device login page.
