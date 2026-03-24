# Implementation Overview

This document describes the key implementation concepts and architecture of the Microsoft MCP server for Microsoft Graph API integration.

## Project Overview

Microsoft MCP is a delegated-access MCP server for Microsoft 365 services including Outlook, Calendar, OneDrive, Contacts, and Teams. It now includes an integrated code-mode orchestration surface that operates over the live Microsoft tool registry.

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

- Uses `FastMCP` for tool registration and management.
- Initializes auth based on `MICROSOFT_MCP_AUTH_METHOD`.
- Exposes the Microsoft Graph tool surface: account, email, calendar, contacts, files, Teams, search, and inbox triage.
- Keeps responses compact via `response_shaping.py`.

#### 4. Integrated Code Mode Surface (`code_mode.py`)

The integrated code-mode layer is a Python-native orchestration runtime that reflects the live FastMCP registry.

It provides:
- `search_tools` for discovery by task description.
- `list_tools` for the active auth-aware tool list.
- `tools_info` for tool metadata and generated interfaces.
- `get_required_keys_for_tool` for required environment/config inspection.
- `call_tool_chain` for sandboxed multi-step orchestration.
- `utcp_codemode_usage` prompt guidance.

The runtime should:
- Build its registry view from the active FastMCP `FunctionTool` objects.
- Preserve auth-aware visibility, including hidden Teams tools under MSAL.
- Generate stable interface text from live tool schemas.
- Return `result` and captured `logs` from code execution.

## Implementation Patterns

### Delegated Access Model

The server acts on behalf of the authenticated user rather than with its own identity. That preserves user-scoped data access and avoids introducing a separate service identity.

### Live Registry Reflection

The code-mode layer should not duplicate the Microsoft Graph tool list by hand. It should reflect the active FastMCP registry so:
- auth-mode-specific tool visibility remains correct
- generated interfaces stay in sync with the actual tool schemas
- discovery works over the same tool set exposed to MCP clients

### Sandbox Model

The code-mode runtime uses a cooperative Python sandbox. It should:
- limit imports to safe modules
- capture console output
- enforce timeouts
- expose `__interfaces` and `__get_tool_interface(...)`
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
- Code mode must operate over the same active tool registry.
- Documentation should distinguish shaping from orchestration.
- Tests should cover both the Graph tools and the code-mode surface.

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

### Calendar Tools
- `list_events`
- `get_event`
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
- `MICROSOFT_MCP_RESPONSE_PROFILE`: Response shaping profile

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
