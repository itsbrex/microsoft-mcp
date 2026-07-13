"""
Test configuration and fixtures for Microsoft MCP tests.
"""

import os
import tempfile
from pathlib import Path
from unittest.mock import Mock

# Install safety boundaries before importing any project module. Tests must
# never discover real credentials or attempt authentication for cresa.com.
_TEST_RUNTIME = tempfile.TemporaryDirectory(prefix="microsoft-mcp-tests-")
_TEST_ROOT = Path(_TEST_RUNTIME.name)
os.environ["MICROSOFT_MCP_ACCOUNT_ID"] = "test-user@cresa.email"
os.environ["MICROSOFT_MCP_TOKENS_DIR"] = str(_TEST_ROOT / "tokens")
os.environ["OUTLOOK_CREDS_CONFIG_DIR"] = str(_TEST_ROOT / "outlook-creds")
os.environ["AZURE_CRED_CACHE_FILE"] = str(_TEST_ROOT / "azure-auth-record.json")
os.environ["AZURE_TOKEN_CACHE_FILE"] = str(_TEST_ROOT / "azure-token-cache")
os.environ["MICROSOFT_MCP_NONINTERACTIVE"] = "1"

_blocked_domains = {
    domain.strip().casefold()
    for domain in os.getenv("MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS", "").split(",")
    if domain.strip()
}
_blocked_domains.add("cresa.com")
os.environ["MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS"] = ",".join(sorted(_blocked_domains))

import pytest  # noqa: E402
import pytest_asyncio  # noqa: E402

from fastmcp import FastMCP  # noqa: E402

from microsoft_mcp.code_mode import CodeModeRuntime  # noqa: E402


@pytest_asyncio.fixture
async def mcp_with_runtime():
    mcp = FastMCP("test")
    return await CodeModeRuntime.create(mcp)


@pytest.fixture
def mock_auth():
    """Fixture providing a mock authentication instance."""
    auth = Mock()
    auth.exists_valid_token.return_value = True
    auth.get_token.return_value = "mock-access-token"
    return auth


@pytest.fixture
def clean_env():
    """Fixture that provides a clean environment for testing."""
    # Store original environment variables
    original_env = dict(os.environ)

    # Clear Microsoft-related environment variables for testing
    env_vars_to_clear = [
        "MICROSOFT_MCP_CLIENT_ID",
        "MICROSOFT_MCP_TENANT_ID",
        "MICROSOFT_MCP_REDIRECT_URI",
    ]

    for var in env_vars_to_clear:
        if var in os.environ:
            del os.environ[var]

    yield

    # Restore original environment
    os.environ.clear()
    os.environ.update(original_env)


@pytest.fixture
def sample_user_data():
    """Fixture providing sample user data for testing."""
    return {
        "id": "12345",
        "displayName": "John Doe",
        "mail": "john.doe@company.com",
        "userPrincipalName": "john.doe@company.com",
        "givenName": "John",
        "surname": "Doe",
        "jobTitle": "Software Engineer",
        "department": "Engineering",
        "companyName": "Test Company",
    }


@pytest.fixture
def sample_email_data():
    """Fixture providing sample email data for testing."""
    return [
        {
            "id": "email1",
            "subject": "Test Email 1",
            "from": {
                "emailAddress": {"address": "sender1@test.com", "name": "Sender One"}
            },
            "toRecipients": [{"emailAddress": {"address": "recipient@test.com"}}],
            "receivedDateTime": "2024-09-01T10:00:00Z",
            "hasAttachments": False,
            "isRead": False,
            "conversationId": "conv1",
        },
        {
            "id": "email2",
            "subject": "Test Email 2",
            "from": {
                "emailAddress": {"address": "sender2@test.com", "name": "Sender Two"}
            },
            "toRecipients": [{"emailAddress": {"address": "recipient@test.com"}}],
            "receivedDateTime": "2024-09-01T11:00:00Z",
            "hasAttachments": True,
            "isRead": True,
            "conversationId": "conv2",
        },
    ]
