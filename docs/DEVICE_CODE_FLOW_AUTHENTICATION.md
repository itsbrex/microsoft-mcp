# Device Code Flow Authentication

## Overview

The email address is **not sent** to initiate device code flow. Microsoft determines the signed-in identity from the account that completes the device prompt, or from a cached MSAL account if silent auth succeeds first.

## Device Code Flow Request

In `src/microsoft_mcp/auth_msal.py`, the device prompt starts here:

```python
flow = app.initiate_device_flow(scopes=DEFAULT_SCOPES)
```

The request body sent to Microsoft's `/devicecode` endpoint contains only the public client ID and scopes:

| Parameter | Value | Description |
|-----------|-------|-------------|
| `client_id` | `d3590ed6-52b3-4102-aeff-aad2292ab01c` (or custom) | Public client ID |
| `scope` | `https://graph.microsoft.com/.default offline_access` | Permissions requested |

The email address is not part of that request.

## What Determines the Authority

The device-code request runs under the MSAL app authority configured before the request is sent. In `microsoft-mcp`, authority resolution works like this:

1. Use `MICROSOFT_MCP_TENANT_ID` if it is explicitly set.
2. Otherwise, if `MICROSOFT_MCP_ACCOUNT_ID` matches an `outlook-creds` profile, reuse that profile's stored authority and tenant.
3. Otherwise, fall back to `common`.

That means `client_id` and `scope` are the only values posted to `/devicecode`, but the tenant-specific authority still matters because it controls which Microsoft identity endpoint handles the request.

## Where the Email Comes From

On a fresh login, the user signs in at `https://login.microsoft.com/device`. Microsoft returns the email in the token response:

```python
id_token_claims = result.get("id_token_claims", {})
email = id_token_claims.get("preferred_username") or id_token_claims.get("email")
```

## What `MICROSOFT_MCP_ACCOUNT_ID` Does

`MICROSOFT_MCP_ACCOUNT_ID` is a local selector. It is **not** sent to the device-code endpoint, but it affects three local behaviors:

### 1. Token file naming

```
~/.config/microsoft-mcp/tokens/
├── user@example.com_access_token.json
├── user@example.com_refresh_only.txt
└── user@example.com_access_only.txt
```

### 2. Cached account selection

Before showing a fresh device code, `microsoft-mcp` tries to reuse cached MSAL accounts. If `MICROSOFT_MCP_ACCOUNT_ID` is set, it narrows the lookup:

```python
accounts = app.get_accounts(username=self.account_identifier)
```

That lets the auth layer prefer silent auth for one specific account.

### 3. Authority resolution from `outlook-creds`

If `MICROSOFT_MCP_TENANT_ID` is unset, `microsoft-mcp` checks for:

```text
~/config/outlook-creds/tokens/<normalized-account>/account_info.json
```

When that file exists, it reuses the stored:

- `authority`
- `realm` / tenant ID
- `aud` / client ID when no explicit client ID was provided

This is the behavior that allows fresh device-code login to use a tenant-specific authority instead of always defaulting to `common`.

## Summary

| Parameter | Sent to Microsoft? | Purpose |
|-----------|-------------------|---------|
| `client_id` | Yes | Identifies the public client |
| `scope` | Yes | Permissions requested |
| `MICROSOFT_MCP_ACCOUNT_ID` | No | Local token naming, cached-account selection, optional authority lookup |
| Authority / tenant | Indirectly | Determined by MSAL app configuration before the request is sent |

The user still chooses which account to authenticate by signing in at the device page, but the configured authority can restrict that choice to a specific tenant.
