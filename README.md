# Microsoft MCP

Powerful MCP server for Microsoft Graph API with dual Azure browser auth and MSAL device-code auth for Outlook, Calendar, OneDrive, Contacts, and Teams.

## Features

- **Dual Authentication**: Azure SDK browser flow or MSAL device code flow
- **Account-Aware MSAL**: Per-account token files, cached-account selection, and optional tenant-specific authority reuse from `outlook-creds`
- **Email Access**: List, inspect, search, and fetch attachments from Outlook mail
- **Calendar Access**: List events, inspect event details, search calendars, and check availability
- **OneDrive Access**: Browse, inspect, and search files
- **Contacts**: List, inspect, and search contacts from your address book
- **Teams Messages**: Read and search chat and channel messages
- **Unified Search**: Search across supported Microsoft Graph content types
- **Code Mode Surface**: Integrated discovery, interface introspection, and one-shot orchestration over the live Microsoft tool registry
- **Flexible Storage**: Configurable Azure credential storage and MSAL token directories

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

## Available Tools

Microsoft MCP exposes the Graph tool surface plus an integrated code-mode surface.

### Graph Tools

- `list_accounts`, `set_active_account`, `get_active_account`, `get_user_details`, `is_logged_in`, `login`
- `list_emails`, `get_email`, `get_attachment`, `search_emails`
- `list_events`, `get_event`, `check_availability`, `search_events`
- `list_contacts`, `get_contact`, `search_contacts`
- `list_files`, `get_file`, `search_files`
- `list_chat_messages`, `get_chat_message`, `search_chat_messages`
- `list_channel_messages`, `get_channel_message`, `search_channel_messages`
- `unified_search`
- `list_inbox_items`, `get_inbox_item_detail`

### Integrated Code Mode Surface

The server now exposes a code-mode orchestration layer that operates over the live tool registry. Use it when you need discovery, tool inspection, or one-shot multi-step workflows.

- `search_tools` - find relevant Microsoft tools by natural language query
- `list_tools` - list the active, auth-aware tool set
- `tools_info` - return complete tool metadata and generated interfaces for selected tools
- `get_required_keys_for_tool` - inspect required configuration for a tool
- `call_tool_chain` - execute multi-step code against the active tool set and return `result` plus `logs`
- `utcp_codemode_usage` - prompt that teaches discovery first, code second

## Code Mode

The server-side code-mode layer is for batching follow-up calls, computing rankings, and reducing payloads after the Graph tools already shaped the data.

### Recommended workflow

1. Call `search_tools` to find the smallest useful tool set.
2. Call `tools_info` or inspect the prompt to confirm the tool contracts.
3. Call `call_tool_chain` with code that uses the live tool namespace.
4. Return a compact report, not raw Graph payloads.

### What to use it for

- Inbox triage where you hydrate only the top items
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
- `MICROSOFT_MCP_RESPONSE_PROFILE`: Response shaping profile (`legacy` or `assistant`, default: `legacy`)

If `MICROSOFT_MCP_TENANT_ID` is not set and `MICROSOFT_MCP_ACCOUNT_ID` matches an existing `outlook-creds` profile, the MSAL auth provider will reuse that profile's tenant-specific authority. This avoids tenant-specific device-code failures such as `AADSTS65002` on fresh login.

### Response Shaping

Use the shaping parameters on the individual tools for raw payload control. Use the code-mode surface when you need orchestration and local reduction after the server has already trimmed the response.
