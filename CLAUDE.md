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
- `MICROSOFT_MCP_CLIENT_ID` (optional) - defaults to Microsoft Office client ID
- `MICROSOFT_MCP_TENANT_ID` (optional) - defaults to "common"
- `MICROSOFT_MCP_TOKENS_DIR` (optional) - token storage directory (defaults to `~/.config/microsoft-mcp/tokens/`)

## Testing

Tests use `unittest.mock` for mocking Azure auth and HTTP responses. The `conftest.py` provides shared fixtures. Focus is on unit testing logic and integration between modules since FastMCP decorators make direct function testing complex.

## Development Guidelines

- Keep `IMPLEMENTATION.md` updated with any architectural changes
- Use virtual environment in `.venv` for all Python execution
- Run `black` or `ruff format` on edited files
- Logging goes to both stderr and `mcp.log` in the working directory
