# Implementation Overview

This document describes the key implementation concepts and architecture of the Microsoft MCP (Model Context Protocol) server for Microsoft Graph API integration.

## Project Overview

Microsoft MCP is a comprehensive MCP server that provides AI assistants with seamless access to Microsoft 365 services including Outlook (Email), Calendar, OneDrive (Files), and Contacts through the Microsoft Graph API. The implementation focuses on delegated access authentication, allowing the application to act on behalf of signed-in users.

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
- **Default Client ID**: Uses Microsoft Office client ID (`d3590ed6-52b3-4102-aeff-aad2292ab01c`) - works out of box
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
- **Authentication Integration**: Uses `AzureAuthentication` class instance for authentication
- **Comprehensive Coverage**: 25+ tools covering email, calendar, contacts, files, and universal search
- **Error Handling**: Consistent error logging and exception propagation
- **Response Optimization**: Configurable body truncation, attachment handling
- **Microsoft Search Integration**: Universal search API with KQL support and interleaved results

#### 4. Server Implementation (`server.py`)
- **Environment Validation**: Checks for required `MICROSOFT_MCP_CLIENT_ID`
- **Authentication Options**: Supports both class-based and function-based authentication approaches
- **Startup Authentication**: Optional validation of authentication before starting MCP server
- **Error Recovery**: Graceful failure with helpful error messages

## Implementation Patterns

### 1. Delegated Access Model
The system implements delegated access where the application acts on behalf of the authenticated user rather than with its own identity. This provides:
- User-scoped data access
- Respect for user permissions
- No need for administrative consent in most cases
- Secure token management

### 2. Error Handling Strategy
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

### 3. Pagination Pattern
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

### 4. Large File Handling
- **Email Attachments**: 3MB threshold for chunked uploads
- **OneDrive Files**: Configurable chunk size (15 x 320KB = ~5MB chunks)
- **Upload Sessions**: Create session → Upload chunks → Finalize

### 5. Microsoft Search API Integration
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

### Email Tools (9 tools)
- **Core Operations**: list, get, create_draft, send, reply, reply_all
- **Management**: update, move, delete
- **Search**: search_emails, get_attachment
- **Features**: Attachment support, folder management, thread handling, date filtering
- **Date Filtering**: `list_emails` supports `start_date` and `end_date` parameters (ISO format, UTC timezone) for precise email retrieval by date range

### Calendar Tools (7 tools)
- **Core Operations**: list_events, get_event, create_event, update_event, delete_event
- **Interaction**: respond_event, check_availability
- **Features**: Recurring events, attendee management, availability checking

### Contact Tools (6 tools)
- **Core Operations**: list_contacts, get_contact, create_contact, update_contact, delete_contact
- **Search**: search_contacts
- **Features**: Multiple email addresses, phone numbers, addresses

### File Tools (6 tools)
- **Core Operations**: list_files, get_file, create_file, update_file, delete_file
- **Search**: search_files
- **Features**: Path-based navigation, download/upload, metadata management

### Teams Message Tools (6 tools)
- **Chat Messages**: list_chat_messages, get_chat_message, search_chat_messages
- **Channel Messages**: list_channel_messages, get_channel_message, search_channel_messages
- **Features**: Message content search, HTML-to-markdown conversion, date filtering, context information (chat/channel details), attachment support, reply handling

### Universal Search Tools (1 tool)
- **unified_search**: Comprehensive Microsoft Search API integration with advanced KQL filtering
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

### 1. Simplified Authentication Architecture
- **Chosen**: Azure SDK-managed authentication with AuthenticationRecord
- **Rationale**: Eliminates complex manual token management, leverages tested Azure SDK functionality
- **Benefits**: 
  - Reduced code complexity and maintenance burden
  - More reliable token refresh handled by Microsoft's own library
  - Platform-specific secure storage automatically handled
  - Better security through established patterns
- **Backward Compatibility**: Module-level functions maintained for existing code
- **Migration Path**: Existing code continues to work with simplified backend

### 2. Delegated vs Application Access
- **Chosen**: Delegated access
- **Rationale**: User-scoped permissions, better security model, no admin consent required
- **Trade-off**: Requires user authentication vs automatic background access

### 3. Authentication Storage Strategy
- **AuthenticationRecord**: Stores authentication metadata (not tokens) in `~/.azure-graph-auth.json`
- **Platform-Specific Token Cache**: Azure SDK manages secure token storage per platform
- **Benefits**: Seamless authentication across restarts, improved security, reduced complexity
- **Implementation**: File-based AuthenticationRecord with Azure SDK's secure token cache

### 4. Error Handling Philosophy
- **Fail Fast**: Validate inputs early, provide clear error messages
- **Logging**: Comprehensive logging to both console and local file (`mcp.log`)
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
└── tools.py             # MCP tool implementations (30+ tools)

tests/
├── conftest.py          # Shared test fixtures
├── test_auth.py         # Azure authentication tests
├── test_auth_msal.py    # MSAL authentication tests
├── test_graph.py        # Graph API client tests
├── test_integration.py  # Module integration tests
├── test_tools_simple.py # Tools logic tests
└── README.md            # Test documentation

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
