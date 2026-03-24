import importlib
import sys
import time
from unittest.mock import patch

import pytest


@pytest.fixture
def load_tools_module(monkeypatch):
    def _load(auth_method: str = "msal"):
        monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", auth_method)
        monkeypatch.setenv(
            "MICROSOFT_MCP_CLIENT_ID",
            "d3590ed6-52b3-4102-aeff-aad2292ab01c"
            if auth_method == "msal"
            else "test-client-id",
        )
        if auth_method == "msal":
            monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "tester@example.com")
        sys.modules.pop("microsoft_mcp.tools", None)
        sys.modules.pop("microsoft_mcp.code_mode", None)
        return importlib.import_module("microsoft_mcp.tools")

    yield _load
    sys.modules.pop("microsoft_mcp.tools", None)
    sys.modules.pop("microsoft_mcp.code_mode", None)


def test_list_tools_returns_active_business_tools(load_tools_module):
    module = load_tools_module("msal")

    result = module.list_tools.fn()
    tool_names = {tool["name"] for tool in result["tools"]}

    assert result["namespace"] == "microsoft"
    assert "list_emails" in tool_names
    assert "call_tool_chain" not in tool_names
    assert "search_tools" not in tool_names
    assert not (set(module.TEAMS_TOOL_NAMES) & tool_names)


def test_search_tools_returns_interfaces(load_tools_module):
    module = load_tools_module("azure")

    result = module.search_tools.fn("email inbox messages", limit=5)

    assert result["count"] >= 1
    first = result["tools"][0]
    assert "python_interface" in first
    assert "required_keys" in first
    assert first["score"] > 0


def test_tools_info_accepts_access_pattern(load_tools_module):
    module = load_tools_module("msal")

    result = module.tools_info.fn(["microsoft.list_emails"])

    assert result["count"] == 1
    tool = result["tools"][0]
    assert tool["found"] is True
    assert tool["name"] == "list_emails"
    assert "class list_emailsInput" in tool["python_interface"]


def test_get_required_keys_for_tool_reflects_auth_mode(load_tools_module):
    module = load_tools_module("msal")

    result = module.get_required_keys_for_tool.fn("list_emails")

    assert result["found"] is True
    assert "MICROSOFT_MCP_CLIENT_ID" in result["required_keys"]
    assert "MICROSOFT_MCP_ACCOUNT_ID" in result["required_keys"]


def test_call_tool_chain_can_use_live_tool_namespace(load_tools_module):
    module = load_tools_module("msal")

    with patch.object(module.auth, "exists_valid_token", return_value=True):
        result = module.call_tool_chain.fn(
            """
status = microsoft.is_logged_in()
print("login status", status)
return {"status": status}
"""
        )

    assert result["result"] == {"status": True}
    assert any("login status True" in line for line in result["logs"])


@patch("microsoft_mcp.tools.graph")
def test_call_tool_chain_accepts_dict_arguments(mock_graph, load_tools_module):
    module = load_tools_module("msal")
    mock_graph.request_paginated.side_effect = [iter([]), iter([])]

    result = module.call_tool_chain.fn(
        """
summary = microsoft.list_inbox_items({"limit": 3})
return {"returned": summary["meta"]["returned"]}
"""
    )

    assert result["result"] == {"returned": 0}


def test_call_tool_chain_timeout_is_reported(load_tools_module):
    module = load_tools_module("msal")

    with pytest.raises(TimeoutError):
        module.call_tool_chain.fn(
            """
import time
time.sleep(0.2)
return {"done": True}
""",
            timeout=0.01,
        )

    time.sleep(0.25)
