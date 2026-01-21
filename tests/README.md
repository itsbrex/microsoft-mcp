# Microsoft MCP Tests

This directory contains unit tests for the Microsoft Graph MCP (Model Context Protocol) server.

## Test Structure

### Core Test Files

- **`test_auth.py`** (17 tests) - Tests for Azure SDK authentication module
  - Authentication credential creation and configuration
  - Token management and caching
  - AuthenticationRecord persistence
  - Scopes configuration (including extended search scopes)
  - Graph client creation
  - Environment variable handling

- **`test_auth_msal.py`** (27 tests) - Tests for MSAL device code flow authentication
  - Token file management (access, refresh tokens)
  - Token validation and expiration handling
  - Device code flow authentication
  - Protocol compliance with AuthProvider
  - Cache clearing and token refresh

- **`test_graph.py`** (17 tests) - Tests for Microsoft Graph API interaction module
  - HTTP request handling and retries
  - Pagination support
  - Search query functionality
  - Error handling for various HTTP status codes
  - Rate limiting and exponential backoff

- **`test_tools_simple.py`** (9 tests) - Tests for MCP tools core functionality
  - Folder mapping configuration
  - Parameter validation logic
  - Search and pagination parameter construction
  - Endpoint URL construction logic

- **`test_integration.py`** (12 tests) - Integration tests between modules
  - Module imports and dependencies
  - Auth and graph module interaction
  - Configuration validation
  - Logging setup verification
  - Unified search tool verification
  - Search hit processing

### Test Configuration

- **`conftest.py`** - Shared test fixtures and configuration
  - Mock authentication instances
  - Sample data fixtures (users, emails)
  - Environment cleanup utilities

## Running Tests

```bash
# Run all tests
pytest tests/

# Run tests with verbose output
pytest tests/ -v

# Run specific test file
pytest tests/test_auth.py -v

# Run tests with coverage (if coverage package is installed)
pytest tests/ --cov=src/microsoft_mcp
```

## Test Coverage (81 tests)

The tests cover the following key areas:

### Azure Authentication (`test_auth.py`) - 17 tests
- ✅ Azure credential creation with various configurations
- ✅ Token acquisition and caching mechanisms
- ✅ AuthenticationRecord serialization/deserialization
- ✅ Default and custom auth file paths
- ✅ Scopes configuration (base + search scopes)
- ✅ Environment variable validation
- ✅ Error handling for missing credentials
- ✅ Graph client instantiation

### MSAL Authentication (`test_auth_msal.py`) - 27 tests
- ✅ Initialization with default and custom values
- ✅ Environment variable configuration
- ✅ Token file path generation
- ✅ Token saving and loading
- ✅ Token expiration validation
- ✅ Token refresh mechanism
- ✅ Device code flow authentication
- ✅ AuthProvider protocol compliance
- ✅ Cache clearing operations

### Graph API (`test_graph.py`) - 17 tests
- ✅ HTTP request construction and execution
- ✅ Authentication header injection
- ✅ Search request special headers (consistency level, prefer)
- ✅ Rate limiting (429) and server error (5xx) retry logic
- ✅ Pagination handling with @odata.nextLink
- ✅ Search query with entity type filtering
- ✅ Various HTTP error code handling
- ✅ Network error handling

### Tools Logic (`test_tools_simple.py`) - 9 tests
- ✅ Folder name mapping and case-insensitive lookup
- ✅ Email parameter validation and limits
- ✅ Search parameter construction
- ✅ Endpoint URL building logic
- ✅ Availability check payload creation
- ✅ Dependency injection verification

### Integration (`test_integration.py`) - 12 tests
- ✅ Module import verification
- ✅ Cross-module communication (auth ↔ graph)
- ✅ Configuration constants
- ✅ Base URL and endpoint validation
- ✅ Unified search tool definition
- ✅ Search entity type validation
- ✅ Search hit processing helper function

## Test Strategy

The tests focus on:

1. **Unit Testing**: Individual functions and classes in isolation
2. **Logic Testing**: Core business logic without external dependencies
3. **Integration Testing**: Module interactions and dependency injection
4. **Error Handling**: Various failure scenarios and edge cases
5. **Configuration Testing**: Environment setup and parameter validation

## Limitations

Due to the FastMCP decorator pattern used in the tools module, the tests focus on:
- Testing the underlying logic and dependencies
- Validating configuration and constants
- Ensuring proper module imports and setup
- Testing core business logic components

Direct function execution testing is handled through dependency mocking and logic validation rather than calling decorated functions directly.

## Dependencies

The tests use:
- `pytest` - Test framework
- `unittest.mock` - Mocking and patching
- Standard library modules for test utilities

All test dependencies are managed through the `[dependency-groups.dev]` section in `pyproject.toml`.
