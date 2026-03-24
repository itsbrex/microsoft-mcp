# Microsoft MCP

Powerful MCP server for Microsoft Graph API with dual Azure browser auth and MSAL device-code auth for Outlook, Calendar, OneDrive, Contacts, and Teams.

## Features

- **🔐 Dual Authentication**: Azure SDK browser flow or MSAL device code flow
- **👤 Account-Aware MSAL**: Per-account token files, cached-account selection, and optional tenant-specific authority reuse from `outlook-creds`
- **📧 Email Access**: List, inspect, search, and fetch attachments from Outlook mail
- **📅 Calendar Access**: List events, inspect event details, search calendars, and check availability
- **📁 OneDrive Access**: Browse, inspect, and search files
- **👥 Contacts**: List, inspect, and search contacts from your address book
- **💬 Teams Messages**: Read and search chat and channel messages
- **🔍 Unified Search**: Search across supported Microsoft Graph content types
- **🗂️ Flexible Storage**: Configurable Azure credential storage and MSAL token directories

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

# Or with Azure SDK (requires app registration)
claude mcp add microsoft-mcp -e MICROSOFT_MCP_CLIENT_ID=your-app-id-here -- uvx --from git+https://github.com/marc-hanheide/microsoft-mcp.git microsoft-mcp
```

### Usage Examples

```bash
# Email examples
> read my latest emails with full content
> show emails from last week

# Calendar examples  
> show my calendar for next week
> check if I'm free tomorrow at 2pm
> search my calendar for quarterly review

# File examples
> list files in my OneDrive
> search for "project proposal" across all my files

# Teams examples
> show the latest messages in my Teams chat
> search Teams messages for "incident review"

# Search examples
> search for "quarterly report" across all services
> find contacts named "Smith"
```

## Available Tools

Microsoft MCP currently exposes **29 MCP tools**:

### Account and Authentication (6 tools)
- **`list_accounts`** - List available MSAL accounts from the token directory
- **`set_active_account`** - Switch the active MSAL account
- **`get_active_account`** - Show the current account and auth method
- **`get_user_details`** - Get current user profile details
- **`is_logged_in`** - Check current authentication status
- **`login`** - Trigger authentication from a tool call

### Email (4 tools)
- **`list_emails`** - List emails with optional body content and date filtering
- **`get_email`** - Get a specific email with attachments
- **`get_attachment`** - Fetch email attachment content
- **`search_emails`** - Search emails by query

### Calendar (4 tools)
- **`list_events`** - List calendar events with details
- **`get_event`** - Get specific event details
- **`check_availability`** - Check free/busy times for scheduling
- **`search_events`** - Search calendar events

### Contacts (3 tools)
- **`list_contacts`** - List all contacts
- **`get_contact`** - Get specific contact details
- **`search_contacts`** - Search contacts by query

### Files (3 tools)
- **`list_files`** - Browse OneDrive files and folders
- **`get_file`** - Download file content
- **`search_files`** - Search files in OneDrive

### Teams Messages (6 tools)
- **`list_chat_messages`** - List messages in a Teams chat
- **`get_chat_message`** - Get a specific Teams chat message
- **`search_chat_messages`** - Search Teams chat messages
- **`list_channel_messages`** - List messages in a Teams channel
- **`get_channel_message`** - Get a specific Teams channel message
- **`search_channel_messages`** - Search Teams channel messages

### Search (1 tool)
- **`unified_search`** - Search across supported Microsoft Graph content types

### Inbox (2 tools)
- **`list_inbox_items`** - List prioritized inbox items (emails and calendar events) as normalized `InboxItem` summaries, ranked by urgency signals (unread, mentions, flagged, meeting proximity)
- **`get_inbox_item_detail`** - Hydrate a single inbox item by its `id` and `kind` to retrieve full body, participants, and action hints

## Code Mode Orchestration

For batch triage workflows, [Code Mode](https://docs.anthropic.com/en/docs/claude-code/overview) is the recommended way to orchestrate calls to this MCP server. Code Mode lets you write TypeScript that calls multiple MCP tools in sequence, compute decisions locally (prioritization, filtering, deduplication), and return a compact report — without sending full Graph payloads to the model on every step.

### When to use Code Mode vs server-side shaping

| Concern | Solution |
|---|---|
| Raw Graph payload bloat (large bodies, unused fields) | Server-side response shaping (already applied by this server) |
| Batching follow-up calls over only the selected items | Code Mode orchestration |
| Computing triage scores or ranking across items | Code Mode orchestration |
| Hydrating only the top-N items from a larger list | Code Mode orchestration |

Code Mode is **not** the primary fix for payload size — the server already trims Graph responses. Code Mode is the right tool when the assistant needs to make conditional, multi-step decisions over a set of items.

### Recommended inbox triage flow

1. Call `list_inbox_items` to get normalized summaries (low token cost)
2. Optionally call `unified_search` or `search_emails` to narrow by keyword or sender
3. Call `get_inbox_item_detail` only for the items that need full context (top 2-3)
4. Compute and return a triage report in the Code Mode script

See [`docs/code-mode-inbox-orchestration.md`](docs/code-mode-inbox-orchestration.md) and [`examples/code-mode/inbox_triage.ts`](examples/code-mode/inbox_triage.ts) for a complete walkthrough.

### Inbox triage example

```typescript
// examples/code-mode/inbox_triage.ts — fetch summaries, hydrate top 3, report
const summaries = await mcp.list_inbox_items({ limit: 20 });
const top3 = summaries.items.slice(0, 3);
const details = await Promise.all(
  top3.map(item => mcp.get_inbox_item_detail({ item_id: item.id, kind: item.kind }))
);
// ... build compact triage report from details
```

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
- `MICROSOFT_MCP_RESPONSE_PROFILE`: Response shaping profile (`legacy` or `assistant`, default: `legacy`). See [Response Shaping](#response-shaping) below.

If `MICROSOFT_MCP_TENANT_ID` is not set and `MICROSOFT_MCP_ACCOUNT_ID` matches an existing `outlook-creds` profile, the MSAL auth provider will reuse that profile's tenant-specific authority. This avoids tenant-specific device-code failures such as `AADSTS65002` on fresh login.

**Reuse outlook-creds account metadata:**
```bash
# Reuse an account identity that already exists in outlook-creds
export MICROSOFT_MCP_AUTH_METHOD=msal
export MICROSOFT_MCP_ACCOUNT_ID="your-email@example.com"
export MICROSOFT_MCP_CLIENT_ID="d3590ed6-52b3-4102-aeff-aad2292ab01c"
uv run microsoft-mcp
```

### Response Shaping

The server shapes Microsoft Graph API responses to reduce token usage when tools are called by AI assistants. You can control this behavior with the `MICROSOFT_MCP_RESPONSE_PROFILE` environment variable:

| Value | Behavior |
|---|---|
| `legacy` (default) | Current behavior. List/search tools return shaped summaries; detail tools return shaped details. Safe for existing integrations. |
| `assistant` | Optimized for AI assistant workflows. Forces summary mode on list tools (suppresses body/detail even when explicitly requested). Contacts use compact shaped output. |

**Per-call override:** The `list_emails`, `list_events`, `list_contacts`, and `list_chat_messages` tools accept an optional `response_profile` parameter (`"auto"`, `"legacy"`, or `"assistant"`) that overrides the environment variable for that single call. The default value `"auto"` defers to the env var.

**Inbox tools** (`list_inbox_items`, `get_inbox_item_detail`) always use assistant-shaped responses regardless of the profile setting.

**Token budgets:** Shaped summary responses are designed to stay within these approximate size budgets:
- `list_emails(limit=10)` -- under 12k characters
- `list_events(limit=10)` -- under 8k characters
- `list_contacts(limit=20)` -- under 10k characters
- `list_chat_messages(limit=10)` -- under 12k characters

```bash
# Opt in to assistant profile
export MICROSOFT_MCP_RESPONSE_PROFILE=assistant
```

### 4. Claude Desktop Configuration

Add to your Claude Desktop configuration:

**macOS**: `~/Library/Application Support/Claude/claude_desktop_config.json`  
**Windows**: `%APPDATA%\Claude\claude_desktop_config.json`  
**Linux**: `~/.config/claude/claude_desktop_config.json`

#### Basic Configuration
```json
{
  "mcpServers": {
    "microsoft": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/marc-hanheide/microsoft-mcp.git", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_CLIENT_ID": "your-app-id-here"
      }
    }
  }
}
```

#### Configuration with Custom Storage
```json
{
  "mcpServers": {
    "microsoft": {
      "command": "uvx",
      "args": ["--from", "git+https://github.com/marc-hanheide/microsoft-mcp.git", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_CLIENT_ID": "your-app-id-here",
        "AZURE_CRED_CACHE_FILE": "/path/to/creds/azure-credentials.json",
        "AZURE_TOKEN_CACHE_FILE": "/path/to/creds/azure-token",
        "MICROSOFT_MCP_REDIRECT_URI": "https://your-app.azurewebsites.net/auth/callback"
      }
    }
  }
}
```

#### Local Development Configuration
```json
{
  "mcpServers": {
    "microsoft": {
      "command": "uv",
      "args": ["--directory", "/path/to/microsoft-mcp", "run", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_CLIENT_ID": "your-app-id-here",
        "AZURE_CRED_CACHE_FILE": "./creds/azure-credentials.json",
        "AZURE_TOKEN_CACHE_FILE": "./creds/azure-token"
      }
    }
  }
}
```

#### MSAL Device Code Flow Configuration
```json
{
  "mcpServers": {
    "microsoft": {
      "command": "uv",
      "args": ["--directory", "/path/to/microsoft-mcp", "run", "microsoft-mcp"],
      "env": {
        "MICROSOFT_MCP_AUTH_METHOD": "msal",
        "MICROSOFT_MCP_ACCOUNT_ID": "your-email@example.com",
        "MICROSOFT_MCP_CLIENT_ID": "d3590ed6-52b3-4102-aeff-aad2292ab01c",
        "MICROSOFT_MCP_TOKENS_DIR": "/path/to/tokens"
      }
    }
  }
}
```

## Authentication & Credential Management

Microsoft MCP supports two runtime auth paths and loads local `.env` settings for `authenticate.py` and `microsoft-mcp` in local development.

### Authentication Flow
1. **Azure SDK mode**: Browser-based sign-in with an `AuthenticationRecord` plus Azure-managed token cache
2. **MSAL mode**: Device-code sign-in with file-based token storage under `~/.config/microsoft-mcp/tokens/`
3. **Account-aware MSAL startup**: `MICROSOFT_MCP_ACCOUNT_ID` selects the cached account and can seed tenant-specific authority from `outlook-creds`
4. **Server startup guard**: `microsoft-mcp` logs the configured and actual auth mode and exits on mismatch

### Storage Options

#### Azure SDK Default Storage Locations
- **AuthenticationRecord**: `~/.ms-graph-mcp-azure-auth-record.json`
- **Token Cache**: `~/.ms-graph-mcp-azure-token-cache.nocache` (platform-specific secure storage)

#### MSAL Default Storage Locations
- **Token Directory**: `~/.config/microsoft-mcp/tokens/`
- **Per-account files**:
  - `your-email@example.com_access_token.json`
  - `your-email@example.com_refresh_only.txt`
  - `your-email@example.com_access_only.txt`

#### Custom Storage Locations
```bash
# Specify custom paths for credentials and tokens
AZURE_CRED_CACHE_FILE=./creds/azure-credentials.json \
AZURE_TOKEN_CACHE_FILE=./creds/azure-token \
./authenticate.py
```

### Token Security
- **No Sensitive Data in AuthenticationRecord**: Only contains metadata for silent authentication
- **Secure Token Storage**: Azure SDK uses platform-specific secure storage (Windows Data Protection API, macOS Keychain, etc.)
- **Automatic Refresh**: Azure SDK handles refresh automatically; MSAL refreshes from the stored refresh token
- **Manual Cache Clearing**: Use `auth.clear_cache()` to force re-authentication

### Multi-Environment Setup
You can maintain separate credentials for different environments:

```bash
# Development environment
AZURE_CRED_CACHE_FILE=./dev-creds/azure-credentials.json ./authenticate.py

# Production environment  
AZURE_CRED_CACHE_FILE=./prod-creds/azure-credentials.json ./authenticate.py
```

## Development

```bash
# Run tests
uv run pytest tests/ -v

# Type checking
uv run pyright

# Format code
uvx ruff format .

# Lint
uvx ruff check --fix --unsafe-fixes .
```

## Example: AI Assistant Scenarios

### Smart Email Management
```python
# List latest emails with full content
emails = list_emails(limit=10, include_body=True)

# List emails from specific date range
recent_emails = list_emails(
    limit=20, 
    include_body=True,
    start_date="2024-01-01T00:00:00Z",
    end_date="2024-01-31T23:59:59Z"
)

# Search mail
matches = search_emails("quarterly report")
```

### Intelligent Scheduling
```python
# Check availability before scheduling
availability = check_availability("2024-01-15T10:00:00Z", "2024-01-15T18:00:00Z", ["colleague@company.com"])

# Search upcoming events
events = search_events("Project Review")
```

### File Management
```python
# Browse and inspect files
files = list_files(path="reports")

# Search across all services
results = unified_search("quarterly report")
```

### Teams Message Discovery
```python
# Read recent Teams messages
messages = list_chat_messages(limit=20)

# Search Teams conversations
hits = search_channel_messages("incident review")
```

## Security Notes

- **AuthenticationRecord**: Contains only authentication metadata, no sensitive tokens
- **Secure Token Storage**: Azure SDK manages tokens using platform-specific secure storage (Windows Data Protection API, macOS Keychain, Linux Secret Service)
- **MSAL Token Files**: Device-code auth stores per-account token files in the configured tokens directory
- **Automatic Token Refresh**: Azure SDK handles token refresh transparently; MSAL refreshes using the saved refresh token
- **Configurable Storage**: Credentials and tokens can be stored in custom locations
- **No Environment Tokens**: No tokens stored in environment variables or code
- **Delegated Permissions**: Uses delegated access (user-scoped) rather than application permissions
- **Principle of Least Privilege**: Only requests necessary permissions

## Troubleshooting

### Authentication Issues
- **Authentication fails**: Check your `MICROSOFT_MCP_CLIENT_ID` is correct
- **`AADSTS65002` during fresh MSAL login**: Set `MICROSOFT_MCP_TENANT_ID` to the account's tenant, or set `MICROSOFT_MCP_ACCOUNT_ID` to an email that already exists in `outlook-creds` so the server can reuse that tenant-specific authority
- **Personal Microsoft accounts**: Use `MICROSOFT_MCP_TENANT_ID=consumers` if needed
- **Token errors**: Clear cache and re-authenticate:
  ```bash
  # Azure SDK: remove AuthenticationRecord and force re-authentication
  rm ~/.ms-graph-mcp-azure-auth-record.json
  # Or if using custom location:
  rm ./creds/azure-credentials.json

  # MSAL: remove per-account token files and run authenticate.py again
  rm -rf ~/.config/microsoft-mcp/tokens
  ```
- **Permission errors**: Ensure all required API permissions are granted in Azure Portal

### Storage Issues
- **Custom storage not working**: Ensure directories exist and are writable:
  ```bash
  mkdir -p ./creds
  chmod 755 ./creds
  ```
- **Token cache issues**: Azure SDK manages token cache automatically, but you can force refresh by clearing AuthenticationRecord

### Connection Issues
- **Network timeouts**: Check internet connection and firewall settings
- **Rate limiting**: Tool automatically retries with exponential backoff
- **Device page not opening automatically**: Open `https://login.microsoft.com/device` manually and enter the code shown by `authenticate.py`

## License

MIT
