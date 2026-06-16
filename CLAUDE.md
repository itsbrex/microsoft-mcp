# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Microsoft MCP is a Model Context Protocol server that provides AI assistants with access to Microsoft 365 services (Outlook, Calendar, OneDrive, Contacts, Teams) via the Microsoft Graph API. It uses delegated access authentication, allowing the application to act on behalf of signed-in users.

## Common Commands

```bash
# Install dependencies
uv sync

# Run the MCP server
uv run microsoft-mcp

# Run authentication (required first time)
MICROSOFT_MCP_CLIENT_ID="your-app-id" uv run authenticate.py

# Run all tests
uv run pytest tests/ -v

# Run specific test file
uv run pytest tests/test_auth.py -v

# Type checking
uv run pyright

# Format code
uvx ruff format .

# Lint and fix
uvx ruff check --fix --unsafe-fixes .
```

## Architecture

### Core Modules (`src/microsoft_mcp/`)

- **`auth.py`** - Azure SDK authentication with `AzureAuthentication` class
  - Uses `InteractiveBrowserCredential` for delegated access
  - Persists `AuthenticationRecord` to `~/.ms-graph-mcp-azure-auth-record.json`
  - Azure SDK handles all token caching/refresh automatically

- **`auth_msal.py`** - MSAL device code flow authentication with `MSALRefreshTokenAuth` class
  - Alternative for CLI/headless environments without browser access
  - File-based token storage, compatible with outlook-creds tokens
  - Uses Microsoft Office client ID by default (works out of box)

- **`auth_base.py`** - Protocol definition for authentication providers
  - `AuthProvider` protocol defining `get_token()`, `exists_valid_token()`, `authenticate()`, `clear_cache()`

- **`graph.py`** - HTTP client for Microsoft Graph API
  - Uses `httpx` with retry logic (429 rate limiting, 5xx errors)
  - Handles pagination via `@odata.nextLink`
  - Supports large file uploads via chunked sessions
  - Global auth instance via `set_auth_instance()`/`get_auth_instance()`

- **`tools.py`** - FastMCP tool implementations (30+ tools)
  - Account Management (8 tools): `list_accounts`, `set_active_account`, `get_active_account`, `authenticate_new_account`, `refresh_all_accounts`, `refresh_account`, `force_reauthenticate_account`, `verify_account_tokens`
  - Email (9 tools): list, get, send, reply, move, delete, attachments
  - Calendar (7 tools): events, availability, responses
  - Contacts (6 tools): CRUD + search
  - Files (6 tools): OneDrive operations
  - Teams (6 tools): chat and channel messages
  - Signatures (2 read-only tools): `list_signatures`, `get_signature` — the assistant can inspect local plain-text signatures but cannot mutate them.
  - Search: unified search across all services with KQL support
  - Initializes global auth instance based on `MICROSOFT_MCP_AUTH_METHOD`

- **`signatures.py`** - Local plain-text signature store
  - Files live at `~/.config/microsoft-mcp/signatures/<account-slug>-<name>.txt` (override via `MICROSOFT_MCP_SIGNATURES_DIR`). Optional `.html` siblings used verbatim for HTML drafts; otherwise `.txt` is auto-converted (`\n` → `<br>\n`, wrapped in `<div class="signature">`).
  - Microsoft Graph does not expose signature settings (no `/me/signature`, not in `mailboxSettings`, no "apply default signature" flag), so this module owns the workaround: `create_email_draft` and `update_email_draft` call `apply_signature` and append to the body before POST/PATCH.
  - Account slug resolution: `MICROSOFT_MCP_SIGNATURE_ACCOUNT` if set, else slugify `MICROSOFT_MCP_ACCOUNT_ID` (lowercase, `@`/`.` → `-`, strip non-`[a-z0-9-]`, collapse repeated `-`).
  - Injection on draft tools: pass `signature="name"` to apply, `signature="none"` to suppress an env default. With no arg, `MICROSOFT_MCP_REPLY_SIGNATURE` is used for reply/reply_all (falling back to `MICROSOFT_MCP_DEFAULT_SIGNATURE`); `MICROSOFT_MCP_DEFAULT_SIGNATURE` is used for new drafts. `update_email_draft` only applies a signature when `body` is supplied.
  - Missing signature files do **not** fail the draft; the tool result includes a `signature_warning` field and the draft is created/updated without a signature.

- **`signatures_cli.py`** - CLI for managing the local store
  - Exposed two ways: standalone `microsoft-mcp-signatures <cmd>` console script and `microsoft-mcp signatures <cmd>` subcommand on the main entry point. `server.main()` dispatches `argv[0] == "signatures"` to the CLI before any Graph imports.
  - Subcommands: `list`, `show`, `set`, `edit`, `rm`, `path`, `dir`. `set` accepts `--from-file`, `--stdin`, or `--editor` ($VISUAL > $EDITOR > vi).

- **`server.py`** - MCP server entry point, validates `MICROSOFT_MCP_CLIENT_ID`

### Key Patterns

**Authentication Flow**: `tools.py` creates auth instance based on `MICROSOFT_MCP_AUTH_METHOD` → sets on `graph` module → all Graph API calls use global instance

**Dual Auth Support**: Azure SDK (browser) or MSAL (device code) via `MICROSOFT_MCP_AUTH_METHOD=azure|msal`

**Multi-Account Support** (MSAL only): Install accounts during setup OR add them at runtime via `authenticate_new_account(email)` (triggers a device-code flow on the server's stderr). Switch the active account via `set_active_account(email)`. Inspect with `list_accounts()` / `get_active_account()`. Tokens stored per-account as `{email}_access_token.json`. Refresh tokens for every saved account at once via `refresh_all_accounts()` (mirrors `outlook auth refresh`). The server auto-refreshes all saved MSAL tokens on startup (opt-out via `MICROSOFT_MCP_REFRESH_ON_STARTUP=0`).

**401 auto-recovery (MSAL):** when Microsoft Graph returns 401 (e.g., after a long-idle session or upstream token revocation), `graph.request` calls `auth.force_refresh()` and replays the request once. If the second attempt also fails, the original 401 surfaces. Azure auth path is unchanged (its SDK manages refresh internally).

**Dependency Injection**: Graph module uses global `_global_auth` instance; tests mock via `set_auth_instance()`

**Error Handling**: All tools log errors with `exc_info=True` and re-raise; HTTP retries use exponential backoff

### Environment Variables

**Azure SDK Auth (default):**
- `MICROSOFT_MCP_CLIENT_ID` (required) - Azure AD application ID
- `MICROSOFT_MCP_TENANT_ID` (optional) - defaults to "common"
- `MICROSOFT_MCP_REDIRECT_URI` (optional) - for non-localhost deployments
- `AZURE_CRED_CACHE_FILE` (optional) - custom AuthenticationRecord path
- `AZURE_TOKEN_CACHE_FILE` (optional) - custom token cache path

**MSAL Auth:**
- `MICROSOFT_MCP_AUTH_METHOD=msal` - enable MSAL device code flow
- `MICROSOFT_MCP_CLIENT_ID` (required) - Microsoft Office client ID: `d3590ed6-52b3-4102-aeff-aad2292ab01c`
- `MICROSOFT_MCP_TENANT_ID` (optional) - defaults to "common"
- `MICROSOFT_MCP_TOKENS_DIR` (optional) - token storage directory (defaults to `~/.config/microsoft-mcp/tokens/`)
- `MICROSOFT_MCP_ACCOUNT_ID` (optional) - account identifier for token file naming (defaults to "default", typically set to user's email)
- `MICROSOFT_MCP_REFRESH_ON_STARTUP` (optional) - defaults to "1" (on for MSAL). Set to "0" to skip the refresh-all-accounts pass at server startup.

**Response Shaping:**
- `MICROSOFT_MCP_RESPONSE_PROFILE` (optional) - `legacy` (default) or `assistant`. Controls response shaping for list/search tools. Individual tool calls can override via `response_profile` parameter.

**Signatures (local plain-text store):**
- `MICROSOFT_MCP_SIGNATURES_DIR` (optional) - signature directory (default `~/.config/microsoft-mcp/signatures/`).
- `MICROSOFT_MCP_SIGNATURE_ACCOUNT` (optional) - account slug used in filenames; defaults to a slug derived from `MICROSOFT_MCP_ACCOUNT_ID`.
- `MICROSOFT_MCP_DEFAULT_SIGNATURE` (optional) - signature name appended to new drafts when `create_email_draft` is called without an explicit `signature`. Used as a fallback for replies when `MICROSOFT_MCP_REPLY_SIGNATURE` is unset.
- `MICROSOFT_MCP_REPLY_SIGNATURE` (optional) - signature name for reply/reply_all drafts.
- `MICROSOFT_MCP_SIGNATURE_RFC3676` (optional) - `1` to use the RFC 3676 `-- ` sig delimiter; default is a blank line.

## MCP Configuration Format

For manual MCP server configuration (Cursor, Claude Desktop, etc.), use the following formats:

### MSAL Auth (Recommended)

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "/path/to/uv",
      "args": ["run", "--python", "3.13", "--project", "/path/to/microsoft-mcp", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_AUTH_METHOD": "msal",
        "MICROSOFT_MCP_ACCOUNT_ID": "your-email@example.com",
        "MICROSOFT_MCP_CLIENT_ID": "d3590ed6-52b3-4102-aeff-aad2292ab01c"
      }
    }
  }
}
```

**Important:**
- Use `microsoft-mcp` entry point (preferred); `python -m microsoft_mcp.server` and `python src/microsoft_mcp/server.py` also work
- Use full path to `uv` executable (find with `which uv`)
- `MICROSOFT_MCP_CLIENT_ID` must be set to `d3590ed6-52b3-4102-aeff-aad2292ab01c` (Microsoft Office client ID)
- `MICROSOFT_MCP_ACCOUNT_ID` identifies which account's tokens to use

### Multiple Accounts

For multiple Microsoft accounts, create separate server entries:

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "/path/to/uv",
      "args": ["run", "--python", "3.13", "--project", "/path/to/microsoft-mcp", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_AUTH_METHOD": "msal",
        "MICROSOFT_MCP_ACCOUNT_ID": "work@company.com",
        "MICROSOFT_MCP_CLIENT_ID": "d3590ed6-52b3-4102-aeff-aad2292ab01c"
      }
    },
    "microsoft-mcp-personal_outlook_com": {
      "command": "/path/to/uv",
      "args": ["run", "--python", "3.13", "--project", "/path/to/microsoft-mcp", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_AUTH_METHOD": "msal",
        "MICROSOFT_MCP_ACCOUNT_ID": "personal@outlook.com",
        "MICROSOFT_MCP_CLIENT_ID": "d3590ed6-52b3-4102-aeff-aad2292ab01c"
      }
    }
  }
}
```

### Azure SDK Auth

```json
{
  "mcpServers": {
    "microsoft-mcp": {
      "command": "/path/to/uv",
      "args": ["run", "--python", "3.13", "--project", "/path/to/microsoft-mcp", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_CLIENT_ID": "your-azure-app-id-uuid"
      }
    }
  }
}
```

## Testing

Tests use `unittest.mock` for mocking Azure auth and HTTP responses. The `conftest.py` provides shared fixtures. Focus is on unit testing logic and integration between modules since FastMCP decorators make direct function testing complex.

**Single-account test fixture policy.** Tests that write MSAL token files (`{email}_access_token.json`, `{email}_refresh_only.txt`) or exercise `refresh_all_accounts` MUST use only one account identifier per test, canonicalized to `broach@cresa.com` (declared as `TEST_EMAIL` at the top of each affected test module). Multi-account fixtures on disk caused real-world auth issues during `refresh_all_accounts`, so the supported pattern — and the only pattern exercised in tests — is single-account. Mismatch/drift tests still write only one token file; they vary the JWT `upn` claim inside that file to simulate misconfiguration, not the filename.

## Development Guidelines

- Keep `IMPLEMENTATION.md` updated with any architectural changes
- Use virtual environment in `.venv` for all Python execution
- Run `black` or `ruff format` on edited files (a PostToolUse hook in `.claude/settings.json` does this automatically on `Write|Edit|MultiEdit`)
- Logging goes to stderr only (MCP protocol uses stdout for JSON-RPC)

## Claude Code Setup

This repo ships a shared `.claude/` so every collaborator gets the same tooling:

- `.claude/settings.json` — permissions allowlist, PostToolUse ruff hook, status line, PostCompact reminder. Committed.
- `.claude/settings.local.json` — per-user overrides. Gitignored.
- `.claude/commands/` — `/test`, `/lint`, `/format`, `/run`, `/auth`, `/auth-refresh`, `/auth-verify`, `/commit-push-pr`, `/techdebt`, `/bridge-regen`.
- `.claude/agents/` — `test-writer` (haiku), `code-simplifier` (haiku), `doc-sync` (haiku), `graph-reviewer` (sonnet).
- `.claude/scripts/` — hook and statusline helpers (bash + jq).

Run `/weekly-audit` monthly (or after a sweep of new tools) to guard against regression of the 2026-04-23 audit findings.

## Self-correcting learning loop

When Claude makes a repeatable mistake, end the correction with:

> "Update your CLAUDE.md so you don't make that mistake again."

For PR reviews, use the `@claude` GitHub Action to let reviewers pin corrections into this file from inline comments, e.g.:

```
nit: don't construct `httpx.AsyncClient` directly, use graph.request
@claude add to CLAUDE.md: all Graph calls must go through microsoft_mcp.graph.request so retry + pagination + auth headers are applied.
```

## Known gotchas

- **MSAL disables Teams tools.** See commit `7dae88f` — MSAL uses the Microsoft Office public client ID (`d3590ed6-…`) which lacks Teams delegated permissions, so Teams tools are unregistered under `MICROSOFT_MCP_AUTH_METHOD=msal`. If you need Teams, register your own Azure AD app with the required Teams delegated permissions and switch to `MICROSOFT_MCP_AUTH_METHOD=azure` with your app's client ID.
- **`server.py` invocation forms (all three work).** `microsoft-mcp` (console script, preferred), `python -m microsoft_mcp.server`, and `python src/microsoft_mcp/server.py` are all valid entry points. See `tests/test_server_entry.py` for regression coverage.
