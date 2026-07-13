# Implementation Overview

This document describes the key implementation concepts and architecture of the Microsoft MCP server for Microsoft Graph API integration.

## Project Overview

Microsoft MCP is a delegated-access MCP server for Microsoft 365 services including Outlook, Calendar, OneDrive, Contacts, and Teams. It includes an integrated code-mode orchestration surface and supports two public tool-surface modes: `codemode_only` by default and optional `hybrid`.

## Architecture

### Core Components

#### 1. Authentication System

The authentication system supports two pluggable providers via the `AuthProvider` protocol:

##### Azure SDK Authentication (`auth.py`) - Default
- `AzureAuthentication` uses Azure Identity and browser-based delegated access.
- Authentication records persist across sessions using `AuthenticationRecord`.
- Token refresh and caching are handled by Azure SDK.

##### MSAL Device Code Authentication (`auth_msal.py`) - Alternative
- `MSALRefreshTokenAuth` uses device code flow for CLI/headless environments.
- Token storage is file-based and account-aware.
- `MICROSOFT_MCP_ACCOUNT_ID` drives cached-account selection and optional authority reuse.
- Access and refresh-token files use atomic same-directory replacement with owner-only permissions.
- `MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS` rejects protected account domains before credential discovery, refresh, device-code auth, or live verification.

##### Authentication Provider Protocol (`auth_base.py`)
```python
class AuthProvider(Protocol):
    def get_token(self) -> str: ...
    def get_token_with_details(self) -> tuple[str, int]: ...
    def exists_valid_token(self) -> bool: ...
    def authenticate(self) -> dict: ...
    def clear_cache(self) -> None: ...
```

#### 2. Graph API Client (`graph.py`)

- Uses `httpx` for Microsoft Graph requests.
- Handles retries for 429 and 5xx responses.
- Supports pagination via `@odata.nextLink`.
- Supports chunked uploads and search endpoint integration.
- Uses the active auth provider injected through `set_auth_instance()`.

#### 3. MCP Tools (`tools.py`)

- Uses FastMCP 3 for tool registration and management while preserving the
  direct `.fn(...)` call contract used by CLIs, tests, and internal helpers.
- Initializes auth based on `MICROSOFT_MCP_AUTH_METHOD`.
- Builds an internal Microsoft business-tool registry: account, email, calendar, contacts, files, Teams, search, and inbox triage.
- Includes assistant-native inbox management helpers for email cleanup and organization:
  read state, categories, folder moves, archive, delete, and bulk actions.
- Includes mail-folder discovery and management helpers:
  list folders, resolve custom folder IDs, create folders, rename folders, and delete folders.
- Includes Outlook master-category helpers:
  list categories, inspect category colors, create categories, update category colors, delete categories, and ensure required categories exist.
- Includes draft-only compose support for Outlook mail:
  create brand-new drafts plus reply and reply-all drafts without sending.
- Exposes either:
  - only the code-mode tools publicly when `MICROSOFT_MCP_TOOL_MODE=codemode_only`
  - or both Graph tools and code-mode tools publicly when `MICROSOFT_MCP_TOOL_MODE=hybrid`
- Keeps responses compact via `response_shaping.py`.

#### 4. Integrated Code Mode Surface (`code_mode.py`)

The integrated code-mode layer is a Python-native orchestration runtime that reflects the internal auth-aware business-tool registry.

It provides:
- `search_tools` for discovery by task description.
- `list_tools` for the active auth-aware tool list.
- `tools_info` for tool metadata and generated interfaces.
- `get_required_keys_for_tool` for required environment/config inspection.
- `call_tool_chain` for sandboxed multi-step orchestration.
- `utcp_codemode_usage` prompt guidance.

The runtime should:
- Build its registry view from the internal auth-aware business-tool objects.
- Preserve auth-aware visibility, including hidden Teams tools under MSAL, even when the public registry is `codemode_only`.
- Generate stable interface text from live tool schemas.
- Return `result` and captured `logs` from code execution.

## Implementation Patterns

### Delegated Access Model

The server acts on behalf of the authenticated user rather than with its own identity. That preserves user-scoped data access and avoids introducing a separate service identity.

### Internal Registry Reflection

The code-mode layer should not duplicate the Microsoft Graph tool list by hand. It should reflect the internal business-tool registry so:
- auth-mode-specific tool visibility remains correct
- generated interfaces stay in sync with the actual tool schemas
- discovery works over the same tool set available to `call_tool_chain`

### Sandbox Model

The code-mode runtime uses a cooperative Python sandbox. It should:
- limit imports to safe modules
- capture console output
- enforce timeouts
- expose `interfaces` and `get_tool_interface(...)` helpers
- block obvious unsafe operations

This is a practical agent execution sandbox, not hardened multi-tenant isolation.

### Error Handling Strategy

```python
try:
    result = graph.request(...)
    return result
except Exception as e:
    logger.error(f"Operation failed: {str(e)}", exc_info=True)
    raise
```

### Response Shaping

The server-side shaping layer remains the first line of defense against token bloat. Code mode is for orchestration after shaping, not for replacing it.

### Compatibility Rules

- Existing Graph tools must remain available unchanged.
- Code mode must operate over the same internal auth-aware tool registry.
- Public exposure must support both `codemode_only` and `hybrid`.
- Documentation should distinguish shaping from orchestration.
- Tests should cover both the Graph tools and the code-mode surface.

## Mail-Port Feature Modules

### Inbox Rules (`rules.py`)

The inbox rules module is deliberately Graph-free. All network I/O lives in `tools.py`; `rules.py` only builds payloads and converts between representations.

**Data model convention:** All public functions in `rules.py` accept snake_case keyword arguments (`sender_contains`, `move_to_folder`, `mark_as_read`, etc.) and emit camelCase Graph payloads (`senderContains`, `moveToFolder`, `markAsRead`). The inverse path — Graph rule dicts to snake_case template dicts — goes through `rule_to_template()`. This keeps the camelCase↔snake_case mapping in one place (`build_rule_payload`) so callers never touch Graph field names directly.

YAML import/export uses the same snake_case template format. `template_to_rule_payload(tpl, folder_resolver)` accepts an optional callable that maps folder display names to Graph folder IDs before building the payload, allowing `import_inbox_rules` to resolve human-readable names at import time.

### Email Templates (`templates_engine.py` + `templates_data/`)

**Search path (user first, then bundled):**
1. `$MICROSOFT_MCP_TEMPLATES_DIR` if set, otherwise `~/.config/microsoft-mcp/templates/`
2. `src/microsoft_mcp/templates_data/` (bundled; organized into `email/` and `calendar/` subdirectories)

User-directory templates shadow built-in templates with the same `(category, name)` key. Files named with a leading underscore (`_foo.yaml`) are skipped during listing.

**Substitution model:** Templates use single-brace `{placeholder}` tokens inside `html_template`. All substituted values are HTML-escaped via `html.escape()` unless the key is in the pre-rendered set (e.g., `agenda_items`, `interviewer_items`) or is a named `conditional_section`. Conditional sections are evaluated with `|` (OR) or `&` (AND) field expressions. The `find_template_variables` / `substitute_variables` API uses double-brace `{{var}}` tokens for plain-text variable substitution in email bodies that don't use the YAML template format.

### Microsoft To-Do (`todo.py`)

Pure module. `parse_due_date(text, *, today)` accepts `"today"`, `"tomorrow"`, `"+Nd"`, or `"YYYY-MM-DD"` and returns a Graph `dueDateTime` dict. The `today` parameter is always injected by callers; the module never calls `date.today()`.

### Signature Parser (`signature_parser.py`)

Extracts contact information from plain-text email signatures and OOO auto-replies. `parse_email_body()` identifies the signature block (via delimiter patterns), extracts the primary contact, and optionally scans prose for alternative contacts (e.g., "while I'm away, contact Jane at jane@co.com"). Job-change signals (`left_company`, `new_company`, `new_email`) are detected from the message body and returned separately. Phone normalization uses `normalize_phone_e164()` which handles US 10-digit, 11-digit (1+10), and international (`+`) formats; returns `""` for unparseable input.

### Intelligence Package (`intel/`)

**Pipeline:** collectors → analyzers → engine.

**Collectors** (`intel/collectors/`): `email.py`, `calendar.py`, `contacts.py`, `threads.py`. Each collector function takes an injected `request: Callable` and an injected `now: datetime`. They never import `graph` at module level. Pagination follows `@odata.nextLink` via `intel/_utils.py:paginate()`.

**Analyzers** (`intel/analyzers/`): `priority.py` (`score_priorities`), `relationships.py` (`analyze_relationships`), `schedule.py` (`analyze_schedule`). Analyzers take collector output dicts and return scored/ranked structures; no network calls.

**Engine** (`intel/engine.py`): `generate_briefing()`, `generate_signals()`, `generate_contact_report()`, `generate_recap()`. Each function takes an injected `request` callable and an injected `now: datetime`. The engine runs the relevant collectors, passes results to analyzers, and assembles the final typed report dict. Return types are declared in `intel/types.py` as `TypedDict` subclasses (`BriefingReport`, `SignalsReport`, `ContactReport`, `RecapReport`).

**Injected clock rationale:** Passing `now` as a parameter rather than calling `datetime.now()` inside collectors/engine ensures that tests can assert on exact output without mocking the module's `datetime`. The CLI layer (`intel_cli.py`) is the only site that calls `datetime.now(ZoneInfo(tz))`.

### Bounce Scanner (`bounces.py`)

**Pattern catalogs:** `SUBJECT_KEYWORDS`, `SENDER_PATTERNS`, `BODY_PATTERNS`, `BOUNCE_REASONS` (regex priority list), `STRONG_SUBJECT_INDICATORS`, `EXCLUDED_SUBJECT_PREFIXES`. These are module-level constants so the CLI's `patterns` subcommand can print them without any network calls.

**Classification flow:** `is_bounce_message(subject, sender_email, body)` applies exclusion prefixes first, then postmaster/mailer-daemon sender matching, then strong subject indicators, then optional body pattern scan. `determine_bounce_reason(subject, body)` runs `BOUNCE_REASONS` regex patterns in priority order (first match wins). `classify_bounce_message(msg)` wraps both and extracts the bounced recipient email from DSN body or subject via `parse_dsn_content()`.

**DSN parsing:** `parse_dsn_content(text)` scans for `X-Failed-Recipients`, `Final-Recipient`, `Original-Recipient`, and `To:` headers (in priority order) plus `Action`, `Status`, `Diagnostic-Code`, and `X-Display-Name` fields.

**Folder scan:** `iter_folder_messages(request, folder_id, *, limit)` yields messages page by page following `@odata.nextLink`. Absolute Graph URLs are stripped to paths before being passed back to `request()`. `scan_folder()` wraps the iterator and filters via `classify_bounce_message()`. CSV output via `write_csv(rows, path)` uses column order defined by `_CSV_FIELDNAMES`.

## Tool Categories

### Account and Authentication Tools
- `list_accounts`
- `set_active_account`
- `get_active_account`
- `get_user_details`
- `is_logged_in`
- `login`

### Email Tools
- `list_emails`
- `get_email`
- `get_attachment`
- `search_emails`
- `create_email_draft`
- `list_mail_folders`
- `get_mail_folder`
- `create_mail_folder`
- `rename_mail_folder`
- `delete_mail_folder`
- `list_master_categories`
- `get_master_category`
- `create_master_category`
- `update_master_category`
- `delete_master_category`
- `ensure_master_categories`
- `mark_email_read`
- `set_email_categories`
- `move_email`
- `archive_email`
- `delete_email`
- `bulk_manage_emails`
- `list_invite_messages`
- `delete_invite_message`

### Calendar Tools
- `list_events`
- `get_event`
- `rsvp_to_event`
- `rsvp_to_invite_message`
- `check_availability`
- `search_events`

### Contact Tools
- `list_contacts`
- `get_contact`
- `search_contacts`

### File Tools
- `list_files`
- `get_file`
- `search_files`

### Teams Message Tools
- `list_chat_messages`
- `get_chat_message`
- `search_chat_messages`
- `list_channel_messages`
- `get_channel_message`
- `search_channel_messages`

### Search and Inbox Tools
- `unified_search`
- `list_inbox_items`
- `get_inbox_item_detail`

`list_inbox_items` and `get_inbox_item_detail` support the `invite_message` kind in addition to standard email and calendar event entries so code-mode inbox triage can act on Outlook meeting notifications that live in the mailbox.

### Inbox Rules Tools
- `list_inbox_rules`
- `get_inbox_rule`
- `create_inbox_rule`
- `update_inbox_rule`
- `delete_inbox_rule`
- `toggle_inbox_rule`
- `reorder_inbox_rules`
- `export_inbox_rules`
- `import_inbox_rules`

### Focused Inbox Override Tools
- `list_focused_overrides`
- `create_focused_override`
- `update_focused_override`
- `delete_focused_override`

### Reply/Forward Draft Tools
- `reply_email_draft`
- `reply_all_email_draft`
- `forward_email_draft`
- `send_email_draft` — the only tool that sends to the wire

### MailTips and Attachment Tools
- `get_mailtips`
- `list_attachments`
- `download_attachments`

### Microsoft To-Do Tools
- `list_todo_lists`
- `create_todo_list`
- `list_tasks`
- `create_task`
- `update_task`
- `complete_task`
- `delete_task`
- `create_task_from_email`

### Email Template Tools
- `list_email_templates`
- `render_email_template`
- `find_template_variables`
- `get_template_placeholders`
- `substitute_template_variables`

### Signature Parser and Phone Tools
- `parse_email_signature`
- `normalize_phone_number`

### Intelligence Report Tools
- `generate_morning_briefing`
- `get_priority_signals`
- `get_contact_intelligence`
- `get_end_of_day_recap`

### Bounce Scanning Tools
- `scan_bounces`

### Code Mode Tools
- `search_tools`
- `list_tools`
- `tools_info`
- `get_required_keys_for_tool`
- `call_tool_chain`

## Configuration

### Environment Variables
- `MICROSOFT_MCP_CLIENT_ID`: Azure AD application ID (required)
- `MICROSOFT_MCP_TENANT_ID`: Tenant ID (optional, defaults to `common`)
- `MICROSOFT_MCP_REDIRECT_URI`: Custom redirect URI (optional)
- `MICROSOFT_MCP_AUTH_METHOD`: `azure` or `msal`
- `MICROSOFT_MCP_TOKENS_DIR`: MSAL token storage directory
- `MICROSOFT_MCP_ACCOUNT_ID`: MSAL account selector and token-file identifier
- `MICROSOFT_MCP_NONINTERACTIVE`: Disable device-code fallback after silent refresh failure
- `MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS`: Comma-separated domains forbidden from authentication, refresh, or live verification
- `MICROSOFT_MCP_TOOL_MODE`: Public tool surface mode (`codemode_only` or `hybrid`)
- `MICROSOFT_MCP_RESPONSE_PROFILE`: Response shaping profile

## UTCP Bridge Conversion Utility

The repo now ships a UTCP bridge-config generator in `utcp_bridge_config.py`.

It should:
- read an existing Claude Desktop config without modifying it
- list available `mcpServers`
- optionally include only named servers
- optionally exclude named servers
- optionally override env vars for wrapped servers, such as forcing `MICROSOFT_MCP_TOOL_MODE=hybrid` for a wrapped `microsoft-mcp`
- write reviewable generated files for UTCP bridge usage

## Documentation Contract

The documentation should now describe:
- how the live registry powers code-mode discovery
- when to use direct Graph tools versus `call_tool_chain`
- how `search_tools` and `tools_info` support tool selection
- how to interpret the sandbox and logs returned by code execution

## Testing Expectations

- Tool contract tests should continue to validate the Graph outputs.
- New tests should validate code-mode registry reflection and sandbox behavior.
- Docs tests should ensure README and example paths remain accurate.
