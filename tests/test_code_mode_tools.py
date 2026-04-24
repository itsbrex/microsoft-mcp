import importlib
import asyncio
import sys
import time
import warnings
from unittest.mock import patch

import pytest


@pytest.fixture
def load_tools_module(monkeypatch):
    def _load(auth_method: str = "msal", tool_mode: str | None = None):
        monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", auth_method)
        monkeypatch.setenv(
            "MICROSOFT_MCP_CLIENT_ID",
            "d3590ed6-52b3-4102-aeff-aad2292ab01c"
            if auth_method == "msal"
            else "test-client-id",
        )
        if auth_method == "msal":
            monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "tester@example.com")
        if tool_mode is None:
            monkeypatch.delenv("MICROSOFT_MCP_TOOL_MODE", raising=False)
        else:
            monkeypatch.setenv("MICROSOFT_MCP_TOOL_MODE", tool_mode)
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


def test_tools_info_accepts_server_prefixed_names(load_tools_module):
    module = load_tools_module("msal")

    result = module.tools_info.fn(
        ["code-mode-mcp:list_emails", "code-mode-mcp:microsoft.list_emails"]
    )

    assert result["count"] == 2
    assert all(tool["found"] is True for tool in result["tools"])
    assert all(tool["name"] == "list_emails" for tool in result["tools"])


def test_get_required_keys_for_tool_reflects_auth_mode(load_tools_module):
    module = load_tools_module("msal")

    result = module.get_required_keys_for_tool.fn("list_emails")

    assert result["found"] is True
    assert "MICROSOFT_MCP_CLIENT_ID" in result["required_keys"]
    assert "MICROSOFT_MCP_ACCOUNT_ID" in result["required_keys"]


def test_get_required_keys_accepts_server_prefixed_tool_names(load_tools_module):
    module = load_tools_module("msal")

    result = module.get_required_keys_for_tool.fn("code-mode-mcp:list_emails")

    assert result["found"] is True
    assert result["tool"] == "list_emails"


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


def test_call_tool_chain_exposes_safe_helper_globals(load_tools_module):
    module = load_tools_module("msal")

    result = module.call_tool_chain.fn(
        """
available = available_tools
iface = get_tool_interface("list_emails")
iface_map = interfaceMapJson
return {
    "available_count": len(available),
    "iface_present": iface is not None,
    "map_has_list_emails": "list_emails" in iface_map,
}
""",
        include_interfaces=True,
    )

    assert result["result"]["available_count"] > 0
    assert result["result"]["iface_present"] is True
    assert result["result"]["map_has_list_emails"] is True
    assert "interface_map_json" in result
    assert "available_tools" in result


def test_call_tool_chain_does_not_emit_restrictedpython_print_warning(
    load_tools_module,
):
    module = load_tools_module("msal")

    with warnings.catch_warnings(record=True) as captured:
        warnings.simplefilter("always")
        result = module.call_tool_chain.fn(
            """
print("hello from sandbox")
return {"ok": True}
"""
        )

    warning_text = [
        str(item.message)
        for item in captured
        if "Prints, but never reads 'printed' variable." in str(item.message)
    ]

    assert warning_text == []
    assert result["result"] == {"ok": True}
    assert "hello from sandbox" in result["logs"]


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


def test_public_tool_registry_defaults_to_code_mode_only(load_tools_module):
    module = load_tools_module("msal")

    tool_names = {
        tool.name for tool in asyncio.run(module.mcp._list_tools_middleware())
    }

    assert tool_names == set(module.CODE_MODE_TOOL_NAMES)


def test_public_tool_registry_can_be_switched_back_to_hybrid(load_tools_module):
    module = load_tools_module("msal", tool_mode="hybrid")

    tool_names = {
        tool.name for tool in asyncio.run(module.mcp._list_tools_middleware())
    }

    assert "list_emails" in tool_names
    assert "search_tools" in tool_names
    assert "call_tool_chain" in tool_names
    assert not (set(module.TEAMS_TOOL_NAMES) & tool_names)


@pytest.mark.asyncio
async def test_call_tool_chain_supports_list_comprehensions(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain("return [n * 2 for n in range(4)]")
    assert result["result"] == [0, 2, 4, 6]


@pytest.mark.asyncio
async def test_call_tool_chain_supports_for_loops(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain(
        """
acc = []
for n in range(3):
    acc.append(n * n)
return acc
"""
    )
    assert result["result"] == [0, 1, 4]


@pytest.mark.asyncio
async def test_call_tool_chain_supports_augmented_assignment(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain(
        """
total = 0
for n in range(5):
    total += n
return total
"""
    )
    assert result["result"] == 10


@pytest.mark.asyncio
async def test_call_tool_chain_default_excludes_interface_catalog(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain("return 42")
    assert result["result"] == 42
    assert "interfaces" not in result
    assert "interface_map_json" not in result
    assert "available_tools" not in result
    assert "available_access_patterns" not in result


@pytest.mark.asyncio
async def test_call_tool_chain_include_interfaces_flag(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain("return 1", include_interfaces=True)
    assert "interfaces" in result
    assert "interface_map_json" in result
    assert "available_tools" in result
    assert "available_access_patterns" in result
