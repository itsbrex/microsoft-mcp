# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Project Overview

Microsoft MCP is a Model Context Protocol server that provides AI assistants with access to Microsoft 365 services (Outlook, Calendar, OneDrive, Contacts, Teams) via the Microsoft Graph API. It uses delegated access authentication, allowing the application to act on behalf of signed-in users.

## Common Commands

```bash
# Install dependencies
uv sync

# Run the MCP server
uv run microsoft-mcp

# Run authentication (required first time, MSAL example)
MICROSOFT_MCP_AUTH_METHOD=msal \
MICROSOFT_MCP_ACCOUNT_ID="your-email@example.com" \
MICROSOFT_MCP_CLIENT_ID="d3590ed6-52b3-4102-aeff-aad2292ab01c" \
uv run authenticate.py

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
  - Supports per-account token files and cached-account selection via `MICROSOFT_MCP_ACCOUNT_ID`
  - Reuses tenant-specific authority metadata from `outlook-creds` when tenant is unset and account metadata exists

- **`auth_base.py`** - Protocol definition for authentication providers
  - `AuthProvider` protocol defining `get_token()`, `exists_valid_token()`, `authenticate()`, `clear_cache()`

- **`graph.py`** - HTTP client for Microsoft Graph API
  - Uses `httpx` with retry logic (429 rate limiting, 5xx errors)
  - Handles pagination via `@odata.nextLink`
  - Supports large file uploads via chunked sessions
  - Global auth instance via `set_auth_instance()`/`get_auth_instance()`
  - Falls back to the auth method configured in the environment after loading `.env`

- **`tools.py`** - FastMCP tool implementations (109 registered tools)
  - Account Management (9 tools): `list_accounts`, `set_active_account`, `get_active_account`, `get_user_details`, `authenticate_new_account`, `refresh_all_accounts`, `refresh_account`, `force_reauthenticate_account`, `verify_account_tokens`
  - Auth (2 tools): `is_logged_in`, `login`
  - Code-Mode / Tool Registry (5 tools): `list_tools`, `search_tools`, `tools_info`, `get_required_keys_for_tool`, `call_tool_chain`
  - Mail Folders (5 tools): `list_mail_folders`, `get_mail_folder`, `create_mail_folder`, `rename_mail_folder`, `delete_mail_folder`
  - Master Categories (6 tools): `list_master_categories`, `get_master_category`, `create_master_category`, `update_master_category`, `delete_master_category`, `ensure_master_categories`
  - Email (13 tools): `list_emails`, `get_email`, `create_email_draft`, `update_email_draft`, `mark_email_read`, `set_email_categories`, `move_email`, `archive_email`, `delete_email`, `bulk_manage_emails`, `list_invite_messages`, `delete_invite_message`, `get_mailtips`
  - Reply/Forward Drafts (4 tools): `reply_email_draft`, `reply_all_email_draft`, `forward_email_draft`, `send_email_draft` — `send_email_draft` is the only tool that sends to the wire; all others create drafts
  - Signatures (2 read-only tools): `list_signatures`, `get_signature`
  - Attachments (4 tools): `get_attachment`, `add_email_attachment`, `list_attachments`, `download_attachments`
  - Inbox Rules (9 tools): `list_inbox_rules`, `get_inbox_rule`, `create_inbox_rule`, `update_inbox_rule`, `delete_inbox_rule`, `toggle_inbox_rule`, `reorder_inbox_rules`, `export_inbox_rules`, `import_inbox_rules`
  - Focused Inbox Overrides (4 tools): `list_focused_overrides`, `create_focused_override`, `update_focused_override`, `delete_focused_override`
  - Calendar (5 tools): `list_events`, `get_event`, `rsvp_to_event`, `rsvp_to_invite_message`, `check_availability`
  - Contacts (2 tools): `list_contacts`, `get_contact`
  - Files (2 tools): `list_files`, `get_file`
  - Search (5 tools): `unified_search`, `search_files`, `search_emails`, `search_events`, `search_contacts`
  - Teams (6 tools): chat and channel messages (disabled under MSAL)
  - Assistant-Native Inbox (2 tools): `list_inbox_items`, `get_inbox_item_detail`
  - Microsoft To-Do (12 tools): `list_todo_lists`, `create_todo_list`, `list_tasks`, `create_task`, `update_task`, `complete_task`, `delete_task`, `list_checklist_items`, `add_checklist_item`, `update_checklist_item`, `delete_checklist_item`, `create_task_from_email`
  - Email Templates (5 tools): `list_email_templates`, `render_email_template`, `find_template_variables`, `get_template_placeholders`, `substitute_template_variables`
  - Signature Parser + Phone (2 tools): `parse_email_signature`, `normalize_phone_number`
  - Intel Reports (4 tools): `generate_morning_briefing`, `get_priority_signals`, `get_contact_intelligence`, `get_end_of_day_recap`
  - Bounce Scanning (1 tool): `scan_bounces`
  - Initializes global auth instance based on `MICROSOFT_MCP_AUTH_METHOD`

- **`signatures.py`** - Local plain-text signature store
  - Files live at `~/.config/microsoft-mcp/signatures/<account-slug>-<name>.txt` (override via `MICROSOFT_MCP_SIGNATURES_DIR`). Optional `.html` siblings used verbatim for HTML drafts; otherwise `.txt` is auto-converted.
  - Draft tools (`create_email_draft`, `update_email_draft`) call `apply_signature` and append to the body before POST/PATCH.

- **`signatures_cli.py`** - CLI for managing signatures
  - Exposed as `microsoft-mcp-signatures <cmd>` or `microsoft-mcp signatures <cmd>`. Subcommands: `list`, `show`, `set`, `edit`, `rm`, `path`, `dir`.

- **`auth_cli.py`** - CLI for refreshing/inspecting MSAL tokens
  - Exposed as `microsoft-mcp-auth <cmd>` or `microsoft-mcp auth <cmd>`. Subcommands: `refresh`, `verify`, `status`, `list`, `test`, `doctor`.

- **`rules.py`** + **`rules_cli.py`** - Outlook inbox message rules
  - `rules.py`: pure helpers for building Graph rule payloads from snake_case kwargs or YAML templates.
  - `rules_cli.py`: CLI exposed as `microsoft-mcp-rules <cmd>` or `microsoft-mcp rules <cmd>`. Subcommands: `list`, `get`, `create`, `delete`, `toggle`, `export`, `import`.

- **`todo.py`** - Microsoft To-Do payload builders
  - Pure module. `parse_due_date()`, `build_task_payload()`, `build_linked_resource()`. Accepts injected `today` for deterministic testing.

- **`templates_engine.py`** + **`templates_data/`** - YAML email/calendar template system
  - YAML-based templates with placeholders and conditional sections. User templates shadow built-ins.

- **`signature_parser.py`** - Email signature and OOO contact/job-change extraction
  - `parse_signature_block()`, `parse_email_body()`, `normalize_phone_e164()`.

- **`code_mode.py`** - Code-mode execution environment
  - Enables `MICROSOFT_MCP_TOOL_MODE=codemode_only` restricting to code-mode tools only.

- **`utcp_bridge_config.py`** - UTCP bridge configuration
  - Exposed as `microsoft-mcp-utcp-config` console script.

- **`inbox_models.py`** - Inbox data models for assistant-native inbox tools

- **`inbox_ranking.py`** - Inbox ranking and prioritization scoring

- **`response_shaping.py`** - Response profile shaping (`legacy` or `assistant` mode)

- **`search_cache.py`** - In-memory search result caching

- **`intel/`** - Intelligence report package: collectors → analyzers → engine
  - `intel/collectors/`: `email.py`, `calendar.py`, `contacts.py`, `threads.py` — each accepts injected `request` and `now`.
  - `intel/analyzers/`: `priority.py`, `relationships.py`, `schedule.py`.
  - `intel/engine.py`: `generate_briefing()`, `generate_signals()`, `generate_contact_report()`, `generate_recap()`.
  - `intel_cli.py`: CLI exposed as `microsoft-mcp-intel <cmd>` or `microsoft-mcp intel <cmd>`.

- **`bounces.py`** + **`bounces_cli.py`** - NDR/bounce classifier and folder scanner
  - `bounces_cli.py`: CLI exposed as `microsoft-mcp-bounces <cmd>` or `microsoft-mcp bounces <cmd>`.

- **`server.py`** - MCP server entry point
  - Loads `.env` before importing auth-sensitive modules
  - Validates `MICROSOFT_MCP_CLIENT_ID`
  - Dispatches `argv[0] in {"signatures", "auth", "rules", "intel", "bounces"}` to the respective CLI before importing the Graph stack

### Key Patterns

**Authentication Flow**: `tools.py` creates auth instance based on `MICROSOFT_MCP_AUTH_METHOD` → sets on `graph` module → all Graph API calls use global instance

**Dual Auth Support**: Azure SDK (browser) or MSAL (device code) via `MICROSOFT_MCP_AUTH_METHOD=azure|msal`

**Multi-Account Support** (MSAL only): Install multiple accounts during setup, switch at runtime via `set_active_account()`. Tokens are stored per-account as `{email}_access_token.json`, and `MICROSOFT_MCP_ACCOUNT_ID` also drives cached-account selection.

**Authority Resolution** (MSAL only): If `MICROSOFT_MCP_TENANT_ID` is unset and the account exists in `outlook-creds`, `auth_msal.py` reuses that profile's tenant-specific authority instead of defaulting to `common`.

**401 auto-recovery (MSAL):** when Microsoft Graph returns 401, `graph.request` calls `auth.force_refresh()` and replays the request once.

**Dependency Injection**: Graph module uses global `_global_auth` instance; tests mock via `set_auth_instance()`. The `intel/` package and `bounces.py` take the request callable as a parameter and accept an injected `now`/`today` datetime.

**Graph-only, no EWS**: All mail operations go through Microsoft Graph REST. No Exchange Web Services (EWS/SOAP) and no lxml dependency.

**Draft-first reply/forward**: `reply_email_draft`, `reply_all_email_draft`, and `forward_email_draft` create drafts only. `send_email_draft` is the single tool that sends to the wire.

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
- `MICROSOFT_MCP_TENANT_ID` (optional) - explicit tenant override
- `MICROSOFT_MCP_TOKENS_DIR` (optional) - token storage directory (defaults to `~/.config/microsoft-mcp/tokens/`)
- `MICROSOFT_MCP_ACCOUNT_ID` (optional) - account identifier for token file naming, cached-account selection, and optional `outlook-creds` authority lookup
- `MICROSOFT_MCP_REFRESH_ON_STARTUP` (optional) - defaults to "1" (on for MSAL). Set to "0" to skip the refresh-all-accounts pass at server startup.

**Response Shaping:**
- `MICROSOFT_MCP_RESPONSE_PROFILE` (optional) - `legacy` (default) or `assistant`. Controls response shaping for list/search tools.

**Signatures (local plain-text store):**
- `MICROSOFT_MCP_SIGNATURES_DIR` (optional) - signature directory (default `~/.config/microsoft-mcp/signatures/`).
- `MICROSOFT_MCP_SIGNATURE_ACCOUNT` (optional) - account slug used in filenames.
- `MICROSOFT_MCP_DEFAULT_SIGNATURE` (optional) - signature name for new drafts.
- `MICROSOFT_MCP_REPLY_SIGNATURE` (optional) - signature name for reply/reply_all drafts.

**Email Templates:**
- `MICROSOFT_MCP_TEMPLATES_DIR` (optional) - user template directory (default `~/.config/microsoft-mcp/templates/`).

**Code-Mode:**
- `MICROSOFT_MCP_TOOL_MODE` (optional) - set to `codemode_only` to restrict to code-mode tools only.
- `MICROSOFT_MCP_CODE_MODE_DIR` (optional) - local directory for code-mode UTCP bridge dist files.
- `MICROSOFT_MCP_UTCP_BRIDGE_COMMAND` (optional) - override command for the UTCP bridge subprocess.

**Compatibility:**
- `OUTLOOK_CREDS_CONFIG_DIR` (optional) - config directory for `outlook-creds` token compatibility.

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

**Test conventions.** Import modules under test with the `src.` prefix — `from src.microsoft_mcp import tools` (source files themselves use `from microsoft_mcp...` with no `src.`). Call a `@mcp.tool`-decorated function via its `.fn` accessor: `tools.scan_bounces.fn(...)`. Pure modules (`intel/`, `bounces.py`, `todo.py`) take a dependency-injected `request` callable and an injected `now`/`today`, so tests pass a fake `request` and a fixed datetime — never the real clock.

## Development Guidelines

- Keep `IMPLEMENTATION.md` updated with any architectural changes
- Use virtual environment in `.venv` for all Python execution
- Run `black` or `ruff format` on edited files
- Logging goes to stderr only (MCP protocol uses stdout for JSON-RPC)
