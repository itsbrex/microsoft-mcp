# Microsoft MCP

Powerful MCP server for Microsoft Graph API with dual Azure browser auth and MSAL device-code auth for Outlook, Calendar, OneDrive, Contacts, and Teams.

## Features

- **Dual Authentication**: Azure SDK browser flow or MSAL device code flow
- **Account-Aware MSAL**: Per-account token files, cached-account selection, and optional tenant-specific authority reuse from `outlook-creds`
- **Email Access**: List, inspect, search, list/download attachments, create reply/reply-all/forward drafts, send drafts, discover/manage mail folders, manage Outlook master categories, manage Outlook inbox state, clean up Outlook invite messages, manage server-side inbox rules, manage Focused Inbox overrides, and get MailTips
- **Calendar Access**: List events, inspect event details, search calendars, check availability, and RSVP quietly by default from events or invite messages
- **OneDrive Access**: Browse, inspect, and search files
- **Contacts**: List, inspect, and search contacts from your address book
- **Teams Messages**: Read and search chat and channel messages
- **Unified Search**: Search across supported Microsoft Graph content types
- **Code Mode Surface**: Integrated discovery, interface introspection, and one-shot orchestration over the live Microsoft tool registry
- **Tool Surface Modes**: `codemode_only` by default, with optional `hybrid` mode for direct Graph tool exposure
- **Flexible Storage**: Configurable Azure credential storage and MSAL token directories
- **Inbox Rules**: List, create, update, delete, toggle, reorder, and YAML import/export server-side Outlook message rules
- **Microsoft To-Do**: Manage task lists, tasks (with due dates), checklists, and create tasks from emails
- **Email Templates**: YAML-based email and calendar templates with placeholder substitution (HTML-escaped, XSS-safe)
- **Intelligence Reports**: Morning briefing, priority signals, contact intelligence, and end-of-day recap from Graph data
- **Bounce Scanning**: NDR/bounce classifier with DSN parsing and CSV export for folder scans

## Quick Start

### One-Line Interactive Installer (Recommended)

```bash
curl -fsSL https://raw.githubusercontent.com/marc-hanheide/microsoft-mcp/main/scripts/install.sh | bash
```

This interactive installer:
- Lets you choose targets: **Claude Code**, **Cursor**, and/or **Claude Desktop**
- Lets you choose auth: **MSAL device code** (no setup) or **Azure SDK** (browser)
- Handles config merging with existing MCP servers
- Runs initial authentication

### Manual Quick Start (Claude Code)

```bash
# Using MSAL device code flow
claude mcp add microsoft-mcp \
  -e MICROSOFT_MCP_AUTH_METHOD=msal \
  -e MICROSOFT_MCP_ACCOUNT_ID=your-email@example.com \
  -e MICROSOFT_MCP_CLIENT_ID=d3590ed6-52b3-4102-aeff-aad2292ab01c \
  -- uvx --from git+https://github.com/marc-hanheide/microsoft-mcp.git microsoft-mcp

# Use hybrid if you want the direct Graph tools publicly exposed as well
claude mcp add microsoft-mcp-hybrid \
  -e MICROSOFT_MCP_AUTH_METHOD=msal \
  -e MICROSOFT_MCP_ACCOUNT_ID=your-email@example.com \
  -e MICROSOFT_MCP_CLIENT_ID=d3590ed6-52b3-4102-aeff-aad2292ab01c \
  -e MICROSOFT_MCP_TOOL_MODE=hybrid \
  -- uvx --from git+https://github.com/marc-hanheide/microsoft-mcp.git microsoft-mcp

# Or with Azure SDK (requires app registration)
claude mcp add microsoft-mcp -e MICROSOFT_MCP_CLIENT_ID=your-app-id-here -- uvx --from git+https://github.com/marc-hanheide/microsoft-mcp.git microsoft-mcp
```

## Available Tools

Microsoft MCP now has two public tool-surface modes controlled by `MICROSOFT_MCP_TOOL_MODE`.

### Default: `codemode_only`

By default, the public MCP registry exposes only the code-mode tools:

- `search_tools`
- `list_tools`
- `tools_info`
- `get_required_keys_for_tool`
- `call_tool_chain`

`call_tool_chain` still operates over the internal auth-aware Microsoft business-tool registry, so it can use mail, calendar, files, contacts, Teams, search, and inbox helpers even when those tools are not publicly exposed.

### Optional: `hybrid`

Set `MICROSOFT_MCP_TOOL_MODE=hybrid` if you want the public registry to expose both the Graph tools and the code-mode surface together.

### Graph Tools

- `list_accounts`, `set_active_account`, `get_active_account`, `get_user_details`, `is_logged_in`, `login`
- `list_emails`, `get_email`, `get_attachment`, `search_emails`, `create_email_draft`
- `list_mail_folders`, `get_mail_folder`, `create_mail_folder`, `rename_mail_folder`, `delete_mail_folder`
- `list_master_categories`, `get_master_category`, `create_master_category`, `update_master_category`, `delete_master_category`, `ensure_master_categories`
- `mark_email_read`, `set_email_categories`, `move_email`, `archive_email`, `delete_email`, `bulk_manage_emails`, `list_invite_messages`, `delete_invite_message`
- `list_events`, `get_event`, `rsvp_to_event`, `rsvp_to_invite_message`, `check_availability`, `search_events`
- `list_contacts`, `get_contact`, `search_contacts`
- `list_files`, `get_file`, `search_files`
- `list_chat_messages`, `get_chat_message`, `search_chat_messages`
- `list_channel_messages`, `get_channel_message`, `search_channel_messages`
- `unified_search`
- `list_inbox_items`, `get_inbox_item_detail`

### Integrated Code Mode Surface

The server exposes a code-mode orchestration layer that operates over the internal live Microsoft tool registry. Use it when you need discovery, tool inspection, or one-shot multi-step workflows.

- `search_tools` - find relevant Microsoft tools by natural language query
- `list_tools` - list the active, auth-aware tool set
- `tools_info` - return complete tool metadata and generated interfaces for selected tools
- `get_required_keys_for_tool` - inspect required configuration for a tool
- `call_tool_chain` - execute multi-step code against the active tool set and return `result` plus `logs`
- `utcp_codemode_usage` - prompt that teaches discovery first, code second

`call_tool_chain` executes Python and exposes runtime helpers as safe variable names:

- `interfaces`
- `available_tools`
- `availableTools`
- `get_tool_interface(name)` and `getToolInterface(name)`
- `interface_map_json` and `interfaceMapJson`

## Code Mode

The server-side code-mode layer is for batching follow-up calls, computing rankings, and reducing payloads after the Graph tools already shaped the data.

### Recommended workflow

1. Call `search_tools` to find the smallest useful tool set.
2. Call `tools_info` or inspect the prompt to confirm the tool contracts.
3. Call `call_tool_chain` with code that uses the live tool namespace.
4. Return a compact report, not raw Graph payloads.

### What to use it for

- Inbox triage where you hydrate only the top items
- Inbox cleanup where you batch archive, delete, or mark mail as read after triage
- Search-and-select flows over mail, calendar, files, or Teams
- Cross-tool workflows that need ranking, grouping, or deduplication
- Compact reports for assistants that should not see every intermediate payload

### What not to use it for

- Raw payload minimization that should be handled by tool shaping
- Single-item lookups where direct Graph tools are enough
- Workflows that do not benefit from batching or local computation

### Example

```python
# Integrated code-mode usage
tools = await mcp.search_tools("inbox triage and follow-up")
print(tools)

selected = await mcp.tools_info(["list_inbox_items", "get_inbox_item_detail"])
print(selected)

result = await mcp.call_tool_chain(
    """
summary = microsoft.list_inbox_items({"limit": 20})
top = summary["items"][:3]
details = [microsoft.get_inbox_item_detail({"item_id": item["id"], "kind": item["kind"]}) for item in top]

return {
    "top_titles": [item["title"] for item in top],
    "actions": [detail["action_hints"][0] if detail["action_hints"] else "review" for detail in details],
}
"""
)
print(result["result"])
print(result["logs"])
```

## Python + TypeScript Code Mode Interop

This repo supports both code-mode styles:

- Integrated Python code mode (built into `microsoft-mcp`)
- Standalone TypeScript UTCP bridge (`@utcp/code-mode-mcp`)

They can run independently or together.

When generating UTCP bridge configs from an existing Claude Desktop config, the converter now skips an existing bridge server (default name: `code-mode-mcp`) unless you explicitly include it. This prevents accidental self-wrapping.

## Manual Setup

### 1. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com) → Microsoft Entra ID → App registrations
2. New registration → Name: `microsoft-mcp`
3. Supported account types: Personal + Work/School
4. Authentication → Allow public client flows: Yes
5. API permissions → Add these delegated permissions:
   - User.Read
   - User.ReadBasic.All
   - Chat.Read
   - Mail.Read
   - Team.ReadBasic.All
   - TeamMember.ReadWrite.All
   - Calendars.Read
   - Files.Read
6. Copy Application ID

### 2. Installation

```bash
git clone https://github.com/marc-hanheide/microsoft-mcp.git
cd microsoft-mcp
uv sync
```

### 3. Authentication

#### Basic Authentication
```bash
# Set your Azure app ID
export MICROSOFT_MCP_CLIENT_ID="your-app-id-here"

# Optional: Set custom redirect URI for non-localhost deployments
# export MICROSOFT_MCP_REDIRECT_URI="https://your-app.azurewebsites.net/auth/callback"

# Run authentication script
uv run authenticate.py
```

#### Custom Credential Storage
You can specify custom locations for storing authentication credentials and tokens:

```bash
# Store credentials and tokens in custom locations
AZURE_CRED_CACHE_FILE=./creds/azure-credentials.json \
AZURE_TOKEN_CACHE_FILE=./creds/azure-token \
MICROSOFT_MCP_CLIENT_ID="your-app-id-here" \
./authenticate.py
```

**Environment Variables for Custom Storage:**
- `AZURE_CRED_CACHE_FILE`: Path to store AuthenticationRecord (authentication metadata)
- `AZURE_TOKEN_CACHE_FILE`: Base path for Azure SDK token cache (platform-specific secure storage)

This allows you to:
- Store credentials in project-specific directories
- Use different credentials for different projects
- Keep authentication data organized
- Facilitate team sharing of configuration (credentials only, not tokens)

#### Alternative: MSAL Device Code Flow Authentication

For CLI/headless environments without browser access, use MSAL device code flow:

```bash
# Use MSAL authentication
export MICROSOFT_MCP_AUTH_METHOD=msal
export MICROSOFT_MCP_ACCOUNT_ID="your-email@example.com"
export MICROSOFT_MCP_CLIENT_ID="d3590ed6-52b3-4102-aeff-aad2292ab01c"

# Optional: set a tenant explicitly if your org does not work with "common"
# export MICROSOFT_MCP_TENANT_ID="your-tenant-guid"

# Run authentication - displays a code to enter at https://login.microsoft.com/device
uv run authenticate.py
```

**MSAL Authentication Benefits:**
- Works in CLI and headless environments
- Supports the Microsoft Office client ID and custom public-client app IDs
- File-based token storage (easy to inspect/manage)
- Compatible with outlook-creds token files and account metadata
- Supports multiple cached accounts via `MICROSOFT_MCP_ACCOUNT_ID`

**MSAL Environment Variables:**
- `MICROSOFT_MCP_AUTH_METHOD`: Set to `msal` to use device code flow (default: `azure`)
- `MICROSOFT_MCP_CLIENT_ID`: Public client ID used for device-code flow
- `MICROSOFT_MCP_TENANT_ID`: Optional Azure AD tenant ID override
- `MICROSOFT_MCP_TOKENS_DIR`: Token storage directory (defaults to `~/.config/microsoft-mcp/tokens/`)
- `MICROSOFT_MCP_ACCOUNT_ID`: Account identifier used for token files, cached-account selection, and optional authority lookup
- `MICROSOFT_MCP_NONINTERACTIVE`: Set to `1`, `true`, `yes`, or `on` to disable device-code fallback after a silent refresh failure; cached credentials are preserved for retry
- `MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS`: Comma-separated account domains that must never authenticate, refresh, or make live verification calls
- `MICROSOFT_MCP_TOOL_MODE`: Public tool surface mode (`codemode_only` or `hybrid`, default: `codemode_only`)
- `MICROSOFT_MCP_RESPONSE_PROFILE`: Response shaping profile (`legacy` or `assistant`, default: `legacy`)

If `MICROSOFT_MCP_TENANT_ID` is not set and `MICROSOFT_MCP_ACCOUNT_ID` matches an existing `outlook-creds` profile, the MSAL auth provider will reuse that profile's tenant-specific authority. `AADSTS65002` is different: it means a Microsoft-owned client is not preauthorized for the requested Microsoft resource, so use your own app registration with the required delegated API permissions.

### Auth CLI (MSAL)

Refresh, inspect, and verify saved Microsoft tokens:

```bash
microsoft-mcp auth refresh             # refresh all accounts
microsoft-mcp auth refresh --api both  # refresh Graph + Outlook tokens
microsoft-mcp auth status              # read-only health (no network)
microsoft-mcp auth verify              # check tokens match their filenames
microsoft-mcp auth test                # live Graph /me check
microsoft-mcp auth doctor              # diagnose perms/dups/expiry
```

Example:

```text
broach@cresa.com
  ✓ Graph: Valid, expires 2026-06-15 21:25:06 UTC
broach@cresa.email
  ✓ Graph: Refreshed, expires 2026-06-15 22:40:11 UTC
```

`--api both` mints both a Graph token (`{id}_access_token.json`) and an Outlook token (`{id}_outlook_access_token.json`) off the shared `{id}_refresh_only.txt`. Color is zero-dependency ANSI, auto-disabled when stdout is not a TTY or `NO_COLOR` is set (`MICROSOFT_MCP_FORCE_COLOR=1` to force).

(Also available as a standalone `microsoft-mcp-auth <cmd>` command.)

### Response Shaping

Use the shaping parameters on the individual tools for raw payload control. Use the code-mode surface when you need orchestration and local reduction after the server has already trimmed the response.

## Email Signatures

Microsoft Graph does not expose Outlook signatures (no `/me/signature`, not in `mailboxSettings`, no draft flag), so this server keeps signatures as local plain-text files and appends them to draft bodies before they are POSTed/PATCHed to Graph.

**Where they live:** `~/.config/microsoft-mcp/signatures/<account-slug>-<name>.txt`. Optional `.html` siblings (e.g., `brian-work-default.html`) are used verbatim for HTML drafts; otherwise the `.txt` is auto-converted (`\n` → `<br>`).

**Account slug:** override with `MICROSOFT_MCP_SIGNATURE_ACCOUNT`, otherwise derived by slugifying `MICROSOFT_MCP_ACCOUNT_ID` (`brian@work.com` → `brian-work-com`).

**CLI** — two equivalent entry points:

```bash
microsoft-mcp-signatures list
microsoft-mcp-signatures set    default --from-file ./brian.txt
microsoft-mcp-signatures set    replies --editor                   # opens $EDITOR
microsoft-mcp-signatures set    default --html --from-file ./brian.html
microsoft-mcp-signatures show   default
microsoft-mcp-signatures edit   replies
microsoft-mcp-signatures rm     default --yes
microsoft-mcp-signatures path   default                            # absolute path
microsoft-mcp-signatures dir                                        # store directory
microsoft-mcp-signatures list --account jp-work
microsoft-mcp-signatures list --account "*"                        # all accounts

# equivalent subcommand on the main entry point:
microsoft-mcp signatures list
```

**Applying a signature to a draft:**

```python
# Explicit signature name
create_email_draft(draft_type="new", to_recipients=[...], body="Hi", signature="default")

# Suppress the env default for a single call
create_email_draft(draft_type="reply", email_id="...", body="Thanks", signature="none")
```

**Env defaults** (each is optional):

| Variable                              | Effect                                                                 |
| ------------------------------------- | ---------------------------------------------------------------------- |
| `MICROSOFT_MCP_DEFAULT_SIGNATURE`     | Name appended to new drafts when `signature=` is omitted.              |
| `MICROSOFT_MCP_REPLY_SIGNATURE`       | Name appended to reply / reply_all drafts (falls back to default).     |
| `MICROSOFT_MCP_SIGNATURES_DIR`        | Override the store directory.                                          |
| `MICROSOFT_MCP_SIGNATURE_ACCOUNT`     | Override the account slug used in filenames.                           |
| `MICROSOFT_MCP_SIGNATURE_RFC3676`     | `1` to use the RFC 3676 `-- ` delimiter; default is a blank line.      |

Missing signature files do not fail the draft — the tool result simply includes a `signature_warning` field and the draft is created without a signature. The assistant can inspect (but never modify) signatures via the read-only `list_signatures` / `get_signature` MCP tools.

## Inbox Rules CLI

Manage server-side Outlook message rules:

```bash
microsoft-mcp rules list
microsoft-mcp rules get <rule-id>
microsoft-mcp rules create --name "Archive newsletters" --subject-contains Newsletter --move-to Archive
microsoft-mcp rules toggle <rule-id>
microsoft-mcp rules export --output rules.yaml
microsoft-mcp rules import rules.yaml --mode sync --dry-run
microsoft-mcp rules delete <rule-id> --confirm

# equivalent standalone entry point:
microsoft-mcp-rules list
```

## Intel CLI

Generate intelligence reports from Microsoft 365 data:

```bash
microsoft-mcp intel briefing --timezone America/Chicago --json
microsoft-mcp intel signals --level critical
microsoft-mcp intel contact user@example.com --days 30 --json
microsoft-mcp intel recap --timezone UTC

# equivalent standalone entry point:
microsoft-mcp-intel briefing --json
```

## Bounces CLI

Scan Outlook folders for NDR/bounce messages:

```bash
microsoft-mcp bounces scan --folder inbox --limit 200 --output bounces.csv
microsoft-mcp bounces scan --json
microsoft-mcp bounces patterns --json

# equivalent standalone entry point:
microsoft-mcp-bounces scan --folder inbox --limit 200 --output bounces.csv
```

## UTCP Bridge Config Generator

The repo includes a non-destructive converter that wraps an existing Claude Desktop `mcpServers` config into a UTCP code-mode bridge configuration.

List the available servers in a Claude Desktop config:

```bash
PYTHONPATH=src python -m microsoft_mcp.utcp_bridge_config \
  "$HOME/Library/Application Support/Claude/claude_desktop_config.json" \
  --list-servers
```

Generate review files without touching the original config:

```bash
PYTHONPATH=src python -m microsoft_mcp.utcp_bridge_config \
  "$HOME/Library/Application Support/Claude/claude_desktop_config.json" \
  --output-dir ./tmp/claude-desktop-utcp-review \
  --include-server microsoft-mcp \
  --include-server notion-mcp \
  --exclude-server attio-old \
  --set-env microsoft-mcp MICROSOFT_MCP_TOOL_MODE hybrid
```

That command writes:

- `.utcp_config.json`
- `claude_desktop_config.utcp.json`
- `manual_map.json`

`manual_map.json` preserves source server names as stable manual aliases (sanitized for UTCP), for example `google-sheets` becomes `google_sheets`.

The original Claude Desktop config is only read, never modified.

The generated bridge config defaults to:

- server name: `code-mode-mcp`
- command: resolved via `shutil.which("npx")`, with `MICROSOFT_MCP_UTCP_BRIDGE_COMMAND` env var as override (falls back to literal `"npx"` for PATH lookup at exec time)
- args: `["@utcp/code-mode-mcp"]`
