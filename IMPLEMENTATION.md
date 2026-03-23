# Implementation Overview

This document describes the key implementation concepts and architecture of the Microsoft MCP (Model Context Protocol) server for Microsoft Graph API integration.

## Project Overview

Microsoft MCP is a comprehensive MCP server that provides AI assistants with delegated access to Microsoft 365 services including Outlook (Email), Calendar, OneDrive (Files), Contacts, and Teams messages through the Microsoft Graph API.

## Architecture

### Core Components

#### 1. Authentication System (Dual Provider Architecture)

The authentication system supports two pluggable providers via the `AuthProvider` protocol:

##### 1a. Azure SDK Authentication (`auth.py`) - Default
- **AzureAuthentication** class leverages Azure SDK's built-in capabilities
- **Azure SDK Integration**: Uses Azure Identity's automatic token caching and refresh token handling
- **AuthenticationRecord**: Persistent authentication across sessions using `~/.ms-graph-mcp-azure-auth-record.json`
- **Delegated Access**: Uses Azure Identity's `InteractiveBrowserCredential` for user authentication
- **Modern Authentication Flow**: Implements authorization code flow with PKCE (Proof Key for Code Exchange)
- **No Manual Token Management**: Eliminates custom token caching, refresh services, and background threads
- **Browser-based Auth**: Opens browser for user sign-in
- **Requires Azure App Registration**: User must register their own Azure AD app

**Key Features:**
- Simplified object-oriented design with minimal state management
- Azure SDK handles all token refresh automatically
- AuthenticationRecord enables silent authentication across application restarts
- No background threads or manual refresh services needed
- Platform-specific secure token storage (Windows Data Protection API, macOS Keychain, etc.)
- Support for multiple tenants (common, consumers, organization-specific)
- Robust error handling with automatic fallback to interactive authentication

##### 1b. MSAL Device Code Authentication (`auth_msal.py`) - Alternative
- **MSALRefreshTokenAuth** class uses MSAL's device code flow
- **CLI/Headless Support**: Works in environments without browser access
- **File-based Token Storage**: Tokens stored in JSON/TXT files (compatible with outlook-creds)
- **Account-Aware Startup**: `MICROSOFT_MCP_ACCOUNT_ID` selects cached accounts and names per-account token files
- **Authority Resolution**: When tenant is unset, the provider can reuse tenant-specific authority metadata from an existing `outlook-creds` account profile
- **Default Client ID**: Uses Microsoft Office client ID (`d3590ed6-52b3-4102-aeff-aad2292ab01c`) unless overridden
- **Manual Token Refresh**: Uses HTTP POST to refresh tokens when expired
- **Configurable Storage**: Tokens stored in `~/.config/microsoft-mcp/tokens/` by default

**Token File Format:**
```json
{
  "email": "user@example.com",
  "access_token": "eyJ0...",
  "expires_at": "2026-01-21T15:30:00Z",
  "scopes": "https://graph.microsoft.com/.default",
  "api_type": "graph"
}
```

##### Authentication Provider Protocol (`auth_base.py`)
```python
class AuthProvider(Protocol):
    def get_token(self) -> str: ...
    def get_token_with_details(self) -> tuple[str, int]: ...
    def exists_valid_token(self) -> bool: ...
    def authenticate(self) -> dict: ...
    def clear_cache(self) -> None: ...
```

##### Provider Selection (`tools.py`)
```python
auth_method = os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower()
if auth_method == "msal":
    auth = MSALRefreshTokenAuth(...)
else:
    auth = AzureAuthentication(...)
graph.set_auth_instance(auth)
```

#### 2. Graph API Client (`graph.py`)
- **HTTP Client**: Uses `httpx` for robust HTTP communication
- **Environment-Aware Fallback**: Loads `.env` and recreates the configured auth provider if no global auth instance was injected
- **Retry Logic**: Implements exponential backoff for rate limiting (429) and server errors (5xx)
- **Pagination Support**: Handles Microsoft Graph `@odata.nextLink` pagination automatically
- **Large File Uploads**: Chunked upload sessions for files >3MB (emails) or custom chunk sizes (OneDrive)
- **Search Integration**: Modern `/search/query` API endpoint support

**Key Features:**
- Request/response logging
- Automatic header management (Authorization, Content-Type, ConsistencyLevel)
- Upload session management for large attachments
- Download capabilities with streaming support

#### 3. MCP Tools (`tools.py`)
- **FastMCP Framework**: Uses FastMCP for tool registration and management
- **Authentication Integration**: Initializes either Azure SDK auth or MSAL auth from the environment
- **Current Tool Surface**: 27 tools covering account management, auth status, email, calendar, contacts, files, Teams messages, and unified search
- **Error Handling**: Consistent error logging and exception propagation
- **Response Optimization**: Configurable body truncation, attachment handling
- **Microsoft Search Integration**: Universal search API with KQL support and interleaved results

#### 4. Server Implementation (`server.py`)
- **Local Env Loading**: Calls `load_dotenv()` before importing modules that read auth configuration
- **Environment Validation**: Checks for required `MICROSOFT_MCP_CLIENT_ID`
- **Startup Guard**: Logs configured auth mode, actual auth class, and active account, then exits on mismatch
- **Error Recovery**: Graceful failure with helpful error messages on configuration mistakes

## Implementation Patterns

### 1. Delegated Access Model
The system implements delegated access where the application acts on behalf of the authenticated user rather than with its own identity. This provides:
- User-scoped data access
- Respect for user permissions
- No need for administrative consent in most cases
- Secure token management

### 2. MSAL Authority Resolution
- **Explicit Tenant Wins**: `MICROSOFT_MCP_TENANT_ID` is used when set
- **Account Metadata Fallback**: If tenant is unset and `MICROSOFT_MCP_ACCOUNT_ID` matches an `outlook-creds` profile, the stored authority and tenant are reused
- **Default Fallback**: Otherwise the provider falls back to `common`
- **Operational Impact**: Fresh device-code login can succeed for tenants where `common` plus the Office client ID fails

### 3. Error Handling Strategy
```python
try:
    # Operation
    result = graph.request(...)
    logger.info(f"Operation successful: {details}")
    return result
except Exception as e:
    logger.error(f"Operation failed: {str(e)}", exc_info=True)
    raise
```

### 4. Pagination Pattern
```python
def request_paginated(path, params=None, limit=None):
    items_returned = 0
    next_link = None
    
    while True:
        result = request("GET", next_link or path, params=params)
        for item in result.get("value", []):
            if limit and items_returned >= limit:
                return
            yield item
            items_returned += 1
        
        next_link = result.get("@odata.nextLink")
        if not next_link:
            break
```

### 5. Large File Handling
- **Email Attachments**: 3MB threshold for chunked uploads
- **OneDrive Files**: Configurable chunk size (15 x 320KB = ~5MB chunks)
- **Upload Sessions**: Create session → Upload chunks → Finalize

### 6. Microsoft Search API Integration
- **Unified Search Endpoint**: `/search/query` API for cross-service content discovery
- **Entity Type Support**: Supports all Microsoft Graph searchable entities with proper validation
- **KQL Query Language**: Full Keyword Query Language support for precise filtering
  - Property restrictions: `from:user@domain.com`, `filetype:pdf`, `sent>=2024-01-01`
  - Boolean operators: `AND`, `OR`, `NOT` for complex query construction
  - Date intervals: `today`, `yesterday`, `"this week"`, `"last month"`
  - Wildcard matching: `serv*` for prefix matching
- **Interleaved Results**: Returns unified results across all content types ranked by relevance
- **Entity Type Restrictions**: Automatic validation and adjustment for Microsoft Graph API limitations
  - `event` and `person` entity types cannot be combined with others
  - `chatMessage` cannot be combined with file-related entity types (`driveItem`, `site`, etc.)
  - Automatic fallback to compatible entity combinations with user warnings
- **Response Processing**: 
  - Automatic entity type detection from `@odata.type`
  - Metadata extraction specific to each entity type
  - HTML-to-markdown conversion for body content
  - Deep link generation for direct content access
- **Performance Optimization**:
  - Configurable response minimization to reduce token usage
  - Body content truncation with length limits
  - Entity type result counting and summaries
- **Error Handling**: Comprehensive error handling for search API limitations and failures
  - Detailed error analysis with specific suggestions for 400 Bad Request errors
  - Entity type compatibility validation to prevent unsupported combinations
  - Proper exception raising instead of returning error responses
  - Diagnostic information for authentication and permission issues

## Tool Categories

### Account and Authentication Tools (6 tools)
- **Account Switching**: `list_accounts`, `set_active_account`, `get_active_account`
- **Auth Status**: `is_logged_in`, `login`, `get_user_details`

### Email Tools (4 tools)
- **Core Operations**: `list_emails`, `get_email`, `get_attachment`
- **Search**: `search_emails`
- **Features**: Attachment support, folder context, date filtering
- **Date Filtering**: `list_emails` supports `start_date` and `end_date` parameters (ISO format, UTC timezone) for precise email retrieval by date range

### Calendar Tools (4 tools)
- **Core Operations**: `list_events`, `get_event`
- **Interaction**: `check_availability`
- **Search**: `search_events`

### Contact Tools (3 tools)
- **Core Operations**: `list_contacts`, `get_contact`
- **Search**: `search_contacts`

### File Tools (3 tools)
- **Core Operations**: `list_files`, `get_file`
- **Search**: `search_files`
- **Features**: Path-based navigation, download/upload, metadata management

### Teams Message Tools (6 tools)
- **Chat Messages**: `list_chat_messages`, `get_chat_message`, `search_chat_messages`
- **Channel Messages**: `list_channel_messages`, `get_channel_message`, `search_channel_messages`
- **Features**: Message content search, HTML-to-markdown conversion, date filtering, and chat/channel context information

### Universal Search Tools (1 tool)
- **unified_search**: Comprehensive Microsoft Search API integration with advanced KQL filtering

### Inbox Tools (2 tools)
- **list_inbox_items**: Returns a ranked list of `InboxItem` summaries drawn from emails and calendar events. Items are scored by urgency signals: unread, direct mentions, flagged state, meeting proximity, and newsletter suppression. Accepts `limit` and `include_kinds` parameters.
- **get_inbox_item_detail**: Hydrates a single item by `item_id` and `kind`, returning the full body, participants, and action hints from the underlying Graph API.
- **Supported Entity Types**: 
  - `message` - Outlook emails
  - `event` - Calendar events
  - `driveItem` - OneDrive/SharePoint files and folders
  - `list` - SharePoint lists  
  - `listItem` - SharePoint list items
  - `site` - SharePoint sites
  - `drive` - OneDrive/SharePoint drives
  - `chatMessage` - Teams chat and channel messages
  - `person` - People in organization
- **KQL Filtering**: Supports Keyword Query Language for precise searches
  - Date filters: `sent>=2024-01-01`, `lastModified="this week"`
  - Sender/recipient: `from:john@company.com`, `to:manager@company.com`
  - Content type: `filetype:pdf`, `filetype:docx`
  - Teams mentions: `IsMentioned:true`
  - Content author: `author:"John Smith"`
- **Response Optimization**: 
  - Configurable body inclusion with length limits
  - Minimal response mode to reduce token usage
  - Entity type result counts and summaries
  - Relevance ranking and deep links
- **Interleaved Results**: Returns unified results across all content types ranked by relevance

## Configuration

### Environment Variables
- `MICROSOFT_MCP_CLIENT_ID`: Azure AD application ID (required)
- `MICROSOFT_MCP_TENANT_ID`: Tenant ID (optional, defaults to "common")
- `MICROSOFT_MCP_REDIRECT_URI`: Custom redirect URI (optional, for non-localhost deployments)
- `MICROSOFT_MCP_AUTH_METHOD`: `azure` or `msal`
- `MICROSOFT_MCP_TOKENS_DIR`: MSAL token storage directory
- `MICROSOFT_MCP_ACCOUNT_ID`: MSAL account selector and token-file identifier

### Required Azure Permissions
```python
SCOPES = [
    "User.Read",                    # Read user profile
    "User.ReadBasic.All",          # Read basic user info
    "Chat.Read",                   # Read chat messages
    "ChannelMessage.Read.All",     # Read channel messages from all channels
    "Mail.Read",                   # Read emails
    "Team.ReadBasic.All",          # Read team info
    "TeamMember.ReadWrite.All",    # Manage team membership
    "Calendars.Read",              # Read calendars
    "Files.Read",                  # Read OneDrive files
]
```

## Key Design Decisions

### 1. Dual Authentication Architecture
- **Chosen**: Azure SDK browser auth plus MSAL device-code auth
- **Rationale**: Preserve Azure SDK secure storage while supporting CLI/headless and multi-account MSAL workflows
- **Benefits**: 
  - Azure SDK path for browser-first delegated access
  - MSAL path for headless environments and account-scoped token files
  - Reuse of known-good tenant authority from `outlook-creds` when available

### 2. Delegated vs Application Access
- **Chosen**: Delegated access
- **Rationale**: User-scoped permissions, better security model, no admin consent required
- **Trade-off**: Requires user authentication vs automatic background access

### 3. Authentication Storage Strategy
- **Azure SDK**: Stores authentication metadata in an `AuthenticationRecord` plus platform-specific token cache
- **MSAL**: Stores per-account token files in the configured tokens directory
- **Benefits**: Seamless auth across restarts, clear separation between browser and device-code flows, and explicit multi-account support

### 4. Error Handling Philosophy
- **Fail Fast**: Validate inputs early, provide clear error messages
- **Logging**: Comprehensive stderr logging for MCP-safe diagnostics
- **User Experience**: Helpful error messages, recovery suggestions

### 5. Logging Configuration
- **Centralized Logging**: Single configuration in `tools.py` that sets up logging for all modules
- **Dual Output**: Logs are written to both console (stderr) and local file (`mcp.log`)
- **Formatted Output**: Includes timestamp, module name, log level, and message
- **File Location**: `mcp.log` is created in the current working directory
- **Log Level**: INFO level by default for comprehensive debugging information

### 6. Response Size Management
- **Body Truncation**: Configurable limits for email body content
- **Attachment Handling**: Metadata only unless explicitly requested
- **Pagination**: Limit-based result sets to manage response sizes

### 7. Inbox Ranking and Normalization
- **InboxItem dataclass** (`inbox_models.py`): Normalized representation of emails and calendar events with fields: `id`, `kind`, `source_tool`, `title`, `snippet`, `participants`, `when`, `state`, `score`, `reason`, `action_hints`, `web_url`, plus ranking signals
- **rank_items()** (`inbox_ranking.py`): Scores items using heuristics: unread (+weight), direct mentions (+weight), flagged (+weight), meeting proximity (+weight), newsletter suppression (-weight)
- **Search cache** (`search_cache.py`): In-memory TTL cache with degraded fallback so inbox listing can tolerate transient Graph API errors

### 8. Code Mode Orchestration Pattern
Code Mode is the recommended orchestration layer for multi-step inbox triage. The server handles payload shaping; Code Mode handles conditional batching over only the items that need full hydration.

**Preferred flow:**
1. `list_inbox_items` — get ranked summaries (low token cost)
2. `search_inbox_items` (optional) — narrow by keyword/sender
3. `get_inbox_item_detail` — hydrate only the top 2-3 items
4. Code Mode computes and returns a triage report locally

Code Mode is **not** the fix for raw Graph payload size — the server's response shaping (`response_shaping.py`) already handles that. Code Mode is the right layer for conditional, multi-step decisions across a set of items.

See `docs/code-mode-inbox-orchestration.md` and `examples/code-mode/inbox_triage.ts`.

## Security Considerations

### 1. Token Security
- AuthenticationRecord file (`~/.azure-graph-auth.json`) contains no sensitive data
- Tokens managed by Azure SDK using platform-specific secure storage
- Tokens have expiration times managed automatically
- Cache can be cleared manually via `clear_cache()` method
- No tokens in environment variables or code

### 2. Permission Model
- Principle of least privilege
- Specific scopes requested
- User consent required
- Delegated (not application) permissions

### 3. Data Handling
- No persistent data storage
- Temporary files for downloads/uploads
- Memory-efficient streaming for large files

## Development and Testing

### Project Structure
```
src/microsoft_mcp/
├── __init__.py          # Package initialization
├── server.py            # MCP server entry point
├── auth.py              # Azure SDK authentication (browser-based)
├── auth_msal.py         # MSAL device code flow authentication (CLI/headless)
├── auth_base.py         # AuthProvider protocol definition
├── graph.py             # Microsoft Graph API client
├── tools.py             # MCP tool implementations (30+ tools)
├── response_shaping.py  # ResponseProfile, BudgetHints, type-specific shapers
├── inbox_models.py      # InboxItem dataclass (normalized cross-service item)
├── inbox_ranking.py     # rank_items() scoring heuristics
└── search_cache.py      # In-memory TTL cache with degraded fallback

tests/
├── conftest.py          # Shared test fixtures
├── test_auth.py         # Azure authentication tests
├── test_auth_msal.py    # MSAL authentication tests
├── test_graph.py        # Graph API client tests
├── test_integration.py  # Module integration tests
├── test_tools_simple.py # Tools logic tests
├── test_inbox_ranking.py # Inbox ranking heuristics tests
├── test_inbox_tools.py  # Inbox tool integration tests
└── README.md            # Test documentation

docs/
└── code-mode-inbox-orchestration.md  # Code Mode guidance for inbox triage

examples/code-mode/
└── inbox_triage.ts      # TypeScript example: fetch summaries, hydrate top 3, report

authenticate.py          # Standalone authentication script
```

### Testing Strategy
- **Unit Tests**: Individual functions and classes with mocked dependencies
- **Integration Tests**: Module interactions and dependency injection
- **Authentication Testing**: Both Azure SDK and MSAL auth providers
- **Tool Coverage**: Logic and parameter validation tests
- **Error Scenarios**: Network failures, invalid inputs, permission issues

### Test Coverage (81 tests)
- **`test_auth.py`** (17 tests): Azure authentication class tests
- **`test_auth_msal.py`** (27 tests): MSAL device code flow tests
- **`test_graph.py`** (17 tests): Graph API client and search tests
- **`test_integration.py`** (12 tests): Module integration tests
- **`test_tools_simple.py`** (9 tests): Tools logic tests

### Development Workflow
1. Environment setup with Azure app registration
2. Authentication using `authenticate.py`
3. Development with uv/Python tooling
4. Testing with pytest
5. Code formatting with black/ruff

## Performance Characteristics

### Typical Response Times
- **Token acquisition**: 50-200ms (cached) / 2-5s (interactive auth)
- **Silent authentication**: 100-300ms (using AuthenticationRecord)
- **Simple API calls**: 200-800ms
- **Paginated requests**: 500ms-2s per page
- **File uploads**: Depends on size, ~1MB/s
- **Search operations**: 800ms-2s

### Rate Limiting
- Microsoft Graph: ~1000 requests/minute/tenant
- Automatic retry with exponential backoff
- 429 status code handling with Retry-After headers

### Memory Usage
- Minimal memory footprint
- Streaming for large files
- No background threads or refresh services
- HTTP connection pooling via httpx
- Azure SDK handles token management efficiently

## Future Considerations

### Potential Enhancements
1. **Multi-account support**: Manage multiple Microsoft accounts simultaneously
2. **Webhook subscriptions**: Real-time notifications for changes
3. **Batch operations**: Multiple API calls in single request
4. **Advanced search**: More sophisticated query capabilities
5. **Collaborative features**: Teams, SharePoint integration

### Scalability Considerations
- **Connection pooling**: Already implemented via httpx
- **Token management**: Azure SDK handles automatic refresh without interruptions
- **Caching strategies**: Response caching for frequently accessed data
- **Resource management**: Connection limits, timeout configuration
- **Thread safety**: Azure SDK provides thread-safe token management

This implementation provides a robust, secure, and comprehensive interface to Microsoft 365 services while maintaining simplicity and reliability for AI assistant integration.
