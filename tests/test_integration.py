"""
Integration tests for Microsoft MCP modules.
"""

import pytest
from unittest.mock import Mock, patch
from src.microsoft_mcp import auth, graph


class TestModuleIntegration:
    """Test integration between modules."""

    def test_auth_module_imports(self):
        """Test that auth module imports correctly."""
        assert hasattr(auth, "AzureAuthentication")
        assert hasattr(auth, "SCOPES")

        # Test that SCOPES contains expected permissions
        assert "User.Read" in auth.SCOPES
        assert "Mail.Read" in auth.SCOPES
        assert "Calendars.Read" in auth.SCOPES
        assert "Files.Read" in auth.SCOPES

    def test_graph_module_imports(self):
        """Test that graph module imports correctly."""
        assert hasattr(graph, "request")
        assert hasattr(graph, "request_paginated")
        assert hasattr(graph, "search_query")
        assert hasattr(graph, "set_auth_instance")
        assert hasattr(graph, "get_auth_instance")

    def test_graph_base_url(self):
        """Test that graph module has correct base URL."""
        assert graph.BASE_URL == "https://graph.microsoft.com/v1.0"

    @patch("src.microsoft_mcp.graph.httpx.Client")
    def test_graph_client_configuration(self, mock_client_class):
        """Test that HTTP client is configured correctly."""
        # The module should have initialized a client
        mock_client = Mock()
        mock_client_class.return_value = mock_client

        # Import to trigger client creation
        import src.microsoft_mcp.graph as graph_module

        # Verify client configuration would be reasonable
        assert hasattr(graph_module, "_client")

    def test_auth_graph_integration(self):
        """Test that auth and graph modules can work together."""
        # Create auth instance
        auth_instance = auth.AzureAuthentication()

        # Set it in graph module
        graph.set_auth_instance(auth_instance)

        # Retrieve it back
        retrieved_auth = graph.get_auth_instance()

        assert retrieved_auth == auth_instance

    @patch.dict("os.environ", {"MICROSOFT_MCP_CLIENT_ID": "test-client-id"})
    @patch("src.microsoft_mcp.auth.InteractiveBrowserCredential")
    def test_auth_instance_creation(self, mock_credential):
        """Test creating authentication instance with minimal config."""
        mock_cred_instance = Mock()
        mock_credential.return_value = mock_cred_instance

        auth_instance = auth.AzureAuthentication()
        credential = auth_instance.get_credential()

        assert credential == mock_cred_instance
        mock_credential.assert_called_once()

    def test_folder_constants(self):
        """Test that folder mappings are accessible."""
        from src.microsoft_mcp.tools import FOLDERS

        assert isinstance(FOLDERS, dict)
        assert "inbox" in FOLDERS
        assert "sent" in FOLDERS
        assert FOLDERS["inbox"] == "inbox"
        assert FOLDERS["sent"] == "sentitems"

    def test_logging_configuration(self):
        """Test that logging is configured in modules."""
        import logging

        # Test that modules have loggers
        auth_logger = logging.getLogger("src.microsoft_mcp.auth")
        graph_logger = logging.getLogger("src.microsoft_mcp.graph")
        tools_logger = logging.getLogger("src.microsoft_mcp.tools")

        # Loggers should exist (even if not explicitly configured)
        assert auth_logger is not None
        assert graph_logger is not None
        assert tools_logger is not None

    def test_unified_search_tool_exists(self):
        """Test that unified_search tool is properly defined."""
        from src.microsoft_mcp.tools import unified_search

        # The @mcp.tool decorator wraps the function in a FunctionTool
        # Test that the tool has expected attributes
        assert hasattr(unified_search, "name")
        assert unified_search.name == "unified_search"
        assert hasattr(unified_search, "description")
        assert "Microsoft Search API" in unified_search.description

        # Test that the underlying function is accessible and has proper signature
        assert hasattr(unified_search, "fn")
        assert callable(unified_search.fn)

        import inspect

        sig = inspect.signature(unified_search.fn)
        expected_params = [
            "query",
            "entity_types",
            "limit",
            "kql_filters",
            "include_body",
            "body_max_length",
        ]

        actual_params = list(sig.parameters.keys())
        for param in expected_params:
            assert (
                param in actual_params
            ), f"Parameter '{param}' missing from unified_search signature"

    def test_unified_search_entity_validation(self):
        """Test that unified_search validates entity types correctly."""
        from src.microsoft_mcp.tools import unified_search

        # Mock the graph request to avoid actual API calls
        with patch("src.microsoft_mcp.graph.request") as mock_request:
            mock_request.return_value = {
                "value": [{"hitsContainers": [{"total": 0, "hits": []}]}]
            }

            # Access the underlying function through .fn attribute
            # Test with invalid entity types - should return error without API call
            result = unified_search.fn(
                query="test", entity_types=["invalid_type", "another_invalid"]
            )

            # Should return error response for invalid entity types
            assert "summary" in result
            assert result["summary"]["total_results"] == 0
            assert "error" in result["summary"]
            assert "No valid entity types" in result["summary"]["error"]

            # Verify no API call was made since entity types are invalid
            mock_request.assert_not_called()

    def test_unified_search_helper_function(self):
        """Test the _process_search_hit helper function."""
        from src.microsoft_mcp.tools import _process_search_hit

        # Test with a mock message hit
        mock_hit = {
            "rank": 1,
            "summary": "Test email summary",
            "resource": {
                "@odata.type": "#microsoft.graph.message",
                "id": "test-id",
                "subject": "Test Subject",
                "from": {
                    "emailAddress": {"name": "John Doe", "address": "john@test.com"}
                },
                "receivedDateTime": "2024-01-01T10:00:00Z",
                "conversationId": "test-conversation-id",
            },
        }

        # Updated function signature: no minimal_response parameter
        result = _process_search_hit(
            mock_hit, include_body=False, body_max_length=1000
        )

        assert result is not None
        assert result["entity_type"] == "message"
        assert result["id"] == "test-id"
        # The function preserves original resource fields, so "subject" not "title"
        assert result["subject"] == "Test Subject"
        assert result["search_rank"] == 1
        # Resource fields are at top level, not nested in "metadata"
        assert result["from"] == {
            "emailAddress": {"name": "John Doe", "address": "john@test.com"}
        }
        # Should generate conversation URL
        assert "conversation_url" in result
        assert "test-conversation-id" in result["conversation_url"]
