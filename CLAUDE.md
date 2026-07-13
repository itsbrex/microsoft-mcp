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

- **`tools.py`** - FastMCP tool implementations (109 registered tools)
  - Account Management (9 tools): `list_accounts`, `set_active_account`, `get_active_account`, `get_user_details`, `authenticate_new_account`, `refresh_all_accounts`, `refresh_account`, `force_reauthenticate_account`, `verify_account_tokens`
  - Auth (2 tools): `is_logged_in`, `login`
  - Code-Mode / Tool Registry (5 tools): `list_tools`, `search_tools`, `tools_info`, `get_required_keys_for_tool`, `call_tool_chain`
  - Mail Folders (5 tools): `list_mail_folders`, `get_mail_folder`, `create_mail_folder`, `rename_mail_folder`, `delete_mail_folder`
  - Master Categories (6 tools): `list_master_categories`, `get_master_category`, `create_master_category`, `update_master_category`, `delete_master_category`, `ensure_master_categories`
  - Email (13 tools): `list_emails`, `get_email`, `create_email_draft`, `update_email_draft`, `mark_email_read`, `set_email_categories`, `move_email`, `archive_email`, `delete_email`, `bulk_manage_emails`, `list_invite_messages`, `delete_invite_message`, `get_mailtips`
  - Reply/Forward Drafts (4 tools): `reply_email_draft`, `reply_all_email_draft`, `forward_email_draft`, `send_email_draft` — `send_email_draft` is the only tool that sends to the wire; all others create drafts
  - Signatures (2 read-only tools): `list_signatures`, `get_signature` — the assistant can inspect local plain-text signatures but cannot mutate them.
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
  - Files live at `~/.config/microsoft-mcp/signatures/<account-slug>-<name>.txt` (override via `MICROSOFT_MCP_SIGNATURES_DIR`). Optional `.html` siblings used verbatim for HTML drafts; otherwise `.txt` is auto-converted (`\n` → `<br>\n`, wrapped in `<div class="signature">`).
  - Microsoft Graph does not expose signature settings (no `/me/signature`, not in `mailboxSettings`, no "apply default signature" flag), so this module owns the workaround: `create_email_draft` and `update_email_draft` call `apply_signature` and append to the body before POST/PATCH.
  - Account slug resolution: `MICROSOFT_MCP_SIGNATURE_ACCOUNT` if set, else slugify `MICROSOFT_MCP_ACCOUNT_ID` (lowercase, `@`/`.` → `-`, strip non-`[a-z0-9-]`, collapse repeated `-`).
  - Injection on draft tools: pass `signature="name"` to apply, `signature="none"` to suppress an env default. With no arg, `MICROSOFT_MCP_REPLY_SIGNATURE` is used for reply/reply_all (falling back to `MICROSOFT_MCP_DEFAULT_SIGNATURE`); `MICROSOFT_MCP_DEFAULT_SIGNATURE` is used for new drafts. `update_email_draft` only applies a signature when `body` is supplied.
  - Missing signature files do **not** fail the draft; the tool result includes a `signature_warning` field and the draft is created/updated without a signature.

- **`signatures_cli.py`** - CLI for managing the local store
  - Exposed two ways: standalone `microsoft-mcp-signatures <cmd>` console script and `microsoft-mcp signatures <cmd>` subcommand on the main entry point. `server.main()` dispatches `argv[0] == "signatures"` to the CLI before any Graph imports.
  - Subcommands: `list`, `show`, `set`, `edit`, `rm`, `path`, `dir`. `set` accepts `--from-file`, `--stdin`, or `--editor` ($VISUAL > $EDITOR > vi).

- **`auth_cli.py`** - CLI for refreshing/inspecting MSAL tokens (MSAL only). Mirrors `outlook auth refresh`.
  - Exposed two ways: standalone `microsoft-mcp-auth <cmd>` console script and `microsoft-mcp auth <cmd>` subcommand on the main entry point (mirroring the signatures CLI). `auth_refresh.py` is now a thin backward-compat shim over this module.
  - Subcommands: `auth refresh [email] [--api graph|outlook|both] [--force] [--json]`, `auth verify [--live] [--json]`, `auth status [--json]` (read-only, no network), `auth list [--json]`, `auth test [--json]` (live Graph `/me`), `auth doctor [--json]` (diagnose perms/dups/expiry).
  - Zero-dependency ANSI color, auto-disabled when stdout is not a TTY or `NO_COLOR` is set (`MICROSOFT_MCP_FORCE_COLOR=1` to force).

- **`rules.py`** + **`rules_cli.py`** - Outlook inbox message rules
  - `rules.py`: pure, Graph-free helpers — `build_rule_payload()` (snake_case kwargs → camelCase Graph payload), `template_to_rule_payload()` (YAML snake_case → Graph payload with optional folder resolver), `rule_to_template()` (inverse), `validate_template()`, `summarize_conditions()`, `summarize_actions()`.
  - `rules_cli.py`: CLI exposed two ways — standalone `microsoft-mcp-rules <cmd>` and `microsoft-mcp rules <cmd>`. Subcommands: `list`, `get`, `create`, `delete`, `toggle`, `export`, `import`. All support `--json`. `import` supports `--mode create|sync` and `--dry-run`. `delete` requires `--confirm`.

- **`todo.py`** - Microsoft To-Do payload builders
  - Pure module (no Graph imports). `parse_due_date(text, *, today)` accepts `"today"`, `"tomorrow"`, `"+Nd"`, `"YYYY-MM-DD"` and returns Graph `dueDateTime` format. `build_task_payload()` and `build_linked_resource()` build Graph `todoTask` and `linkedResource` payloads. The `today` parameter is always injected for deterministic testing.

- **`templates_engine.py`** + **`templates_data/`** - YAML email/calendar template system
  - Search path: `$MICROSOFT_MCP_TEMPLATES_DIR` (or `~/.config/microsoft-mcp/templates/`) user dir first, then bundled `templates_data/` directory. User templates shadow built-in ones by `(category, name)` key.
  - Templates are YAML files with `name`, `html_template`, `placeholders` (with `required`, `default`), and optional `conditional_sections` (with `condition` using `|`/`&` field expressions). All non-pre-rendered placeholder values are HTML-escaped before substitution (XSS-safe).
  - `list_templates()`, `load_template()`, `render_template()`, `validate_template_data()`. Variable substitution: `find_template_variables()` and `substitute_variables()` handle `{{var}}` tokens in plain-text content (also decodes HTML-encoded variants). `parse_recipients_csv()` reads a CSV file for bulk recipient expansion.

- **`signature_parser.py`** - Email signature and OOO contact/job-change extraction
  - `parse_signature_block(text)` → contact dict with `first_name`, `last_name`, `full_name`, `job_title`, `company`, `work_email`, `mobile_phone`, `business_phone`, `website`, `linkedin`, `twitter`, `confidence_score`.
  - `parse_email_body(body, *, html, extract_alternatives)` → `{contacts, job_changes}` — detects signature block, extracts primary contact plus alternative contacts from OOO prose, classifies job-change signals (`left_company`, `new_company`, `new_email`).
  - `normalize_phone_e164(phone, default_region="US")` → E.164 string or `""`.

- **`intel/`** - Intelligence report package: collectors → analyzers → engine
  - `intel/_utils.py`: `paginate(request, path, params, *, limit)` follows `@odata.nextLink` using injected request; `parse_graph_datetime()` handles both `Z` and `+00:00` formats.
  - `intel/collectors/`: `email.py`, `calendar.py`, `contacts.py`, `threads.py` — each accepts an injected `request` callable and an injected `now` datetime; no global Graph imports.
  - `intel/analyzers/`: `priority.py` (`score_priorities`), `relationships.py` (`analyze_relationships`), `schedule.py` (`analyze_schedule`).
  - `intel/engine.py`: `generate_briefing()`, `generate_signals()`, `generate_contact_report()`, `generate_recap()` — orchestrate collectors + analyzers and return typed dicts (`BriefingReport`, `SignalsReport`, `ContactReport`, `RecapReport` from `intel/types.py`). All take an injected `request` callable and injected `now` datetime.
  - `intel_cli.py`: CLI exposed two ways — standalone `microsoft-mcp-intel <cmd>` and `microsoft-mcp intel <cmd>`. Subcommands: `briefing [--timezone TZ] [--limit N] [--json]`, `signals [--timezone TZ] [--level all|critical|important|informational] [--json]`, `contact <email> [--days N] [--json]`, `recap [--timezone TZ] [--json]`.

- **`code_mode.py`** - Code-mode execution environment
  - Enables `MICROSOFT_MCP_TOOL_MODE=codemode_only` which restricts the server to the 5 code-mode tools (`list_tools`, `search_tools`, `tools_info`, `get_required_keys_for_tool`, `call_tool_chain`) and disables all other tools.
  - Loads a local UTCP bridge distribution from `MICROSOFT_MCP_CODE_MODE_DIR` or fetches it remotely.

- **`utcp_bridge_config.py`** - UTCP bridge configuration
  - Generates bridge configuration for code-mode. Exposed as `microsoft-mcp-utcp-config` console script.

- **`inbox_models.py`** - Inbox data models for assistant-native inbox tools
  - Pydantic-style models for inbox items, ranking metadata, and response shaping.

- **`inbox_ranking.py`** - Inbox ranking and prioritization
  - Scores and sorts inbox items by urgency, sender importance, and recency for the assistant-native inbox view.

- **`response_shaping.py`** - Response profile shaping logic
  - Transforms raw Graph API responses into shaped formats controlled by `MICROSOFT_MCP_RESPONSE_PROFILE` (`legacy` or `assistant`).

- **`search_cache.py`** - Search result caching
  - In-memory cache for search results to avoid redundant Graph API calls within a session.

- **`bounces.py`** + **`bounces_cli.py`** - NDR/bounce classifier and folder scanner
  - `bounces.py`: pure module with injected request — pattern catalogs (`SUBJECT_KEYWORDS`, `SENDER_PATTERNS`, `BODY_PATTERNS`, `BOUNCE_REASONS`, `STRONG_SUBJECT_INDICATORS`, `EXCLUDED_SUBJECT_PREFIXES`); `is_bounce_message()`, `determine_bounce_reason()`, `classify_bounce_message()`, `parse_dsn_content()`, `iter_folder_messages(request, folder_id, *, limit)` (follows `@odata.nextLink`), `scan_folder(request, folder_id, *, limit)`, `write_csv(rows, path)`.
  - `bounces_cli.py`: CLI exposed two ways — standalone `microsoft-mcp-bounces <cmd>` and `microsoft-mcp bounces <cmd>`. Subcommands: `scan [--folder FOLDER] [--limit N] [--output CSV_PATH] [--json]`, `patterns [--json]` (read-only, no Graph calls).

- **Dual Graph/Outlook tokens.** `MSALRefreshTokenAuth(api_type="outlook")` writes `{id}_outlook_access_token.json` using `outlook.office365.com/.default` and the SHARED `{id}_refresh_only.txt` (Graph tokens stay in `{id}_access_token.json`). `auth refresh --api=both` mints both off the one refresh token.

- **`server.py`** - MCP server entry point, validates `MICROSOFT_MCP_CLIENT_ID`. Dispatches `argv[0] in {"signatures", "auth", "rules", "intel", "bounces"}` to the respective CLI before importing the Graph stack.

### Key Patterns

**Authentication Flow**: `tools.py` creates auth instance based on `MICROSOFT_MCP_AUTH_METHOD` → sets on `graph` module → all Graph API calls use global instance

**Dual Auth Support**: Azure SDK (browser) or MSAL (device code) via `MICROSOFT_MCP_AUTH_METHOD=azure|msal`

**Multi-Account Support** (MSAL only): Install accounts during setup OR add them at runtime via `authenticate_new_account(email)` (triggers a device-code flow on the server's stderr). Switch the active account via `set_active_account(email)`. Inspect with `list_accounts()` / `get_active_account()`. Tokens stored per-account as `{email}_access_token.json`. Refresh tokens for every saved account at once via `refresh_all_accounts()` (mirrors `outlook auth refresh`). The server auto-refreshes all saved MSAL tokens on startup (opt-out via `MICROSOFT_MCP_REFRESH_ON_STARTUP=0`).

**401 auto-recovery (MSAL):** when Microsoft Graph returns 401 (e.g., after a long-idle session or upstream token revocation), `graph.request` calls `auth.force_refresh()` and replays the request once. If the second attempt also fails, the original 401 surfaces. Azure auth path is unchanged (its SDK manages refresh internally).

**Dependency Injection**: Graph module uses global `_global_auth` instance; tests mock via `set_auth_instance()`. The `intel/` package and `bounces.py` take the request callable as a parameter (never importing `graph` globally) and accept an injected `now`/`today` datetime — this keeps them pure and unit-testable.

**Graph-only, no EWS**: All mail operations go through Microsoft Graph REST (`/me/messages`, `/me/mailFolders`, `/me/messageRules`, etc.). No Exchange Web Services (EWS/SOAP) and no lxml dependency.

**Draft-first reply/forward**: `reply_email_draft`, `reply_all_email_draft`, and `forward_email_draft` create drafts only. `send_email_draft` is the single tool that actually sends to the wire.

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
- `MICROSOFT_MCP_NONINTERACTIVE` (optional) - set to `1`/`true`/`yes`/`on` to **disable** the interactive device-code fallback. When a silent token refresh fails (expired/revoked/65002), the code normally falls through to an interactive device-code flow on stderr; in a headless deployment (cron, CI, detached service) that would hang forever. With this set, those paths raise a clear actionable error instead. **Off by default** — interactive behavior is unchanged. Does not affect the explicit `authenticate`/`force_reauthenticate` entry points (those are interactive by contract).
- `MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS` (optional) - comma-separated account domains that must never authenticate, refresh, or make live verification calls. Domain matching includes subdomains. Pytest always adds `cresa.com` and points all credential stores at a temporary directory before project imports.

**Response Shaping:**
- `MICROSOFT_MCP_RESPONSE_PROFILE` (optional) - `legacy` (default) or `assistant`. Controls response shaping for list/search tools. Individual tool calls can override via `response_profile` parameter.

**Signatures (local plain-text store):**
- `MICROSOFT_MCP_SIGNATURES_DIR` (optional) - signature directory (default `~/.config/microsoft-mcp/signatures/`).
- `MICROSOFT_MCP_SIGNATURE_ACCOUNT` (optional) - account slug used in filenames; defaults to a slug derived from `MICROSOFT_MCP_ACCOUNT_ID`.
- `MICROSOFT_MCP_DEFAULT_SIGNATURE` (optional) - signature name appended to new drafts when `create_email_draft` is called without an explicit `signature`. Used as a fallback for replies when `MICROSOFT_MCP_REPLY_SIGNATURE` is unset.
- `MICROSOFT_MCP_REPLY_SIGNATURE` (optional) - signature name for reply/reply_all drafts.
- `MICROSOFT_MCP_SIGNATURE_RFC3676` (optional) - `1` to use the RFC 3676 `-- ` sig delimiter; default is a blank line.

**Email Templates:**
- `MICROSOFT_MCP_TEMPLATES_DIR` (optional) - user template directory (default `~/.config/microsoft-mcp/templates/`). Templates here shadow built-in templates of the same `(category, name)` key.

**Code-Mode:**
- `MICROSOFT_MCP_TOOL_MODE` (optional) - set to `codemode_only` to restrict the server to the 5 code-mode tools only.
- `MICROSOFT_MCP_CODE_MODE_DIR` (optional) - local directory for code-mode UTCP bridge dist files.
- `MICROSOFT_MCP_UTCP_BRIDGE_COMMAND` (optional) - override command for the UTCP code-mode bridge subprocess.

**Compatibility:**
- `OUTLOOK_CREDS_CONFIG_DIR` (optional) - config directory for `outlook-creds` token compatibility.
- `MICROSOFT_MCP_FORCE_COLOR` (optional) - set to `1` to force ANSI color output in CLIs even when stdout is not a TTY.

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

**Test conventions.** Import modules under test with the `src.` prefix — `from src.microsoft_mcp import tools` (source files themselves use `from microsoft_mcp...` with no `src.`). Call a `@mcp.tool`-decorated function via its `.fn` accessor: `tools.scan_bounces.fn(...)`. Pure modules (`intel/`, `bounces.py`, `todo.py`) take a dependency-injected `request` callable and an injected `now`/`today`, so tests pass a fake `request` and a fixed datetime — never the real clock. `tests/test_tool_surface.py` guards that every mail-port tool stays registered; update it when adding/removing tools.

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
- `.claude/commands/` — `/test`, `/lint`, `/format`, `/run`, `/auth`, `/auth-refresh`, `/auth-verify`, `/auth-status`, `/rules`, `/intel`, `/bounces`, `/commit-push-pr`, `/techdebt`, `/bridge-regen`, `/weekly-audit`, `/triage-inbox`, `/declutter-inbox`.
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
- **Draft-first design.** `reply_email_draft`, `reply_all_email_draft`, and `forward_email_draft` always create drafts. Only `send_email_draft` sends. Callers that want to send immediately must call `send_email_draft` after the draft is created.
- **Intel + bounces inject their clock.** `todo.parse_due_date()`, all `intel/` collectors, `intel/engine.py` functions, and `bounces.iter_folder_messages()` receive `today`/`now` as a parameter — never calling `datetime.now()` or `date.today()` internally. Tests must always pass an explicit datetime.
- **`pyyaml` is now a required dependency** (added for inbox rules YAML import/export and the template engine). It is listed in `pyproject.toml` as `pyyaml>=6.0,<7`.
- **The ruff PostToolUse hook strips unused `_`-aliased imports mid-edit.** When you add `from . import bounces as _bounces` (or `from .intel import engine as _intel_engine`) to `tools.py` in one edit and the *first use* lands in a later edit, the `ruff check --fix` hook deletes the import as unused (F401) before the use exists — producing a runtime `NameError`. Add the import and at least one use in the **same** edit, or append ` # noqa: F401` if the use is genuinely elsewhere. This bit `_intel_engine` and `_bounces` during the mail port.
- **No `auth force-reauthenticate` CLI subcommand.** The `auth` CLI exposes exactly `refresh`, `verify`, `status`, `list`, `doctor`, `test` (run `microsoft-mcp auth -h`). Force re-auth is **`auth refresh <email> --force`** — `--force` is a bare flag (clears that account's tokens + re-runs device-code flow), the email is the positional arg: `microsoft-mcp auth refresh broach@cresa.com --force --api both`. The MCP *tool* `force_reauthenticate_account` exists, but there is **no** matching `force-reauthenticate` CLI command — don't invent one.
- **`api_type` on the refresh tools/CLI.** Both the MCP tools `refresh_all_accounts(api_type=...)` and `refresh_account(email, api_type=...)` and the CLI `auth refresh --api ...` accept `graph` (default) / `outlook` / `both`. `both` on `refresh_account` returns a **list** of two result dicts (graph then outlook); the underlying `auth_msal.refresh_account()` only takes `graph`/`outlook`, so the tool and CLI both expand `both` into two calls (the refresh token is shared — see the persist invariant below).
- **Failed-refresh `hint` field.** `auth_msal.classify_refresh_error(error, identifier)` is a pure classifier that maps known AADSTS codes to `{code, summary, remedy}`. `_refresh_one` attaches it as a `hint` key on `status=="failed"` results, so the MCP tools' result dicts, the `auth refresh --json` output, and the human CLI output all carry actionable recovery guidance. **AADSTS65002** means a Microsoft-owned first-party client is not preauthorized for the requested first-party resource; repeating interactive consent with that client does not fix it. Use an app registration authorized for the requested API. `auth doctor`/`status` stay network-free and do not synthesize this (a refresh-time error can't be seen without refreshing).
- **Refresh token is shared across API resources.** `{id}_refresh_only.txt` belongs to the user/client pair, not Graph or Outlook. Each successful refresh persists Microsoft's latest replacement refresh token, following Microsoft identity-platform guidance.
- **`graph.request` takes no `headers` kwarg.** Its signature is `request(method, path, params=None, json=None, data=None, max_retries=3, auth=None)` (all parameters are positional-or-keyword, no `*` separator). Anything that needs a per-call header (e.g. `Prefer: outlook.timezone`, or `$count=true` which requires `ConsistencyLevel: eventual`) cannot send it — convert/derive client-side instead (the intel calendar collector converts Graph UTC datetimes locally; sent-count paginates instead of using `$count`).
- **Code-mode `call_tool_chain` sandbox gotchas.** When driving tools via `mcp__microsoft-mcp__call_tool_chain` (the `microsoft.<tool>()` namespace):
  - **Results truncate at ~1800 chars.** Do all filtering/classification *inside* the sandbox and `print()` compact output; never expect a full multi-email dump to come back.
  - **`list_emails` / `list_*` return a list directly**, not `{"result": [...]}`. Guard with `x if isinstance(x, list) else x.get("result", [])`.
  - **Dunder attribute access is banned** (`x.__name__`, `type(x).__name__` → `invalid attribute name` error). Use `isinstance()`.
  - **Email `body` is a dict** `{"contentType": ..., "content": ...}` — unwrap `body.get("content")` before running regex/text logic.
  - **No raw `conversationId` is exposed.** Derive a thread token from `conversation_url`: `url.split("readconv/")[1].split("?")[0]`. Same token across inbox/sent/drafts = same thread → use it for reply/draft de-duplication.
  - **The harness renders `[REDACTED_EMAIL]` in tool output**, but matching runs server-side on the real values, so address comparisons inside the sandbox are accurate despite the redaction. Reusable triage/declutter pipelines built on these facts live in `.claude/commands/triage-inbox.md` and `declutter-inbox.md`.
