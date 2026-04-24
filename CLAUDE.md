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
  - Account Management (3 tools): `list_accounts`, `set_active_account`, `get_active_account`
  - Email (9 tools): list, get, send, reply, move, delete, attachments
  - Calendar (7 tools): events, availability, responses
  - Contacts (6 tools): CRUD + search
  - Files (6 tools): OneDrive operations
  - Teams (6 tools): chat and channel messages
  - Search: unified search across all services with KQL support
  - Initializes global auth instance based on `MICROSOFT_MCP_AUTH_METHOD`

- **`server.py`** - MCP server entry point, validates `MICROSOFT_MCP_CLIENT_ID`

### Key Patterns

**Authentication Flow**: `tools.py` creates auth instance based on `MICROSOFT_MCP_AUTH_METHOD` → sets on `graph` module → all Graph API calls use global instance

**Dual Auth Support**: Azure SDK (browser) or MSAL (device code) via `MICROSOFT_MCP_AUTH_METHOD=azure|msal`

**Multi-Account Support** (MSAL only): Install multiple accounts during setup, switch at runtime via `set_active_account()`. Tokens stored per-account as `{email}_access_token.json`.

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

**Response Shaping:**
- `MICROSOFT_MCP_RESPONSE_PROFILE` (optional) - `legacy` (default) or `assistant`. Controls response shaping for list/search tools. Individual tool calls can override via `response_profile` parameter.

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

## Development Guidelines

- Keep `IMPLEMENTATION.md` updated with any architectural changes
- Use virtual environment in `.venv` for all Python execution
- Run `black` or `ruff format` on edited files (a PostToolUse hook in `.claude/settings.json` does this automatically on `Write|Edit|MultiEdit`)
- Logging goes to stderr only (MCP protocol uses stdout for JSON-RPC)

## Claude Code Setup

This repo ships a shared `.claude/` so every collaborator gets the same tooling:

- `.claude/settings.json` — permissions allowlist, PostToolUse ruff hook, status line, PostCompact reminder. Committed.
- `.claude/settings.local.json` — per-user overrides. Gitignored.
- `.claude/commands/` — `/test`, `/lint`, `/format`, `/run`, `/auth`, `/commit-push-pr`, `/techdebt`, `/bridge-regen`.
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
