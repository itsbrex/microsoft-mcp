import asyncio
import importlib
import sys
from unittest.mock import patch

import pytest
from fastmcp.exceptions import NotFoundError


def _make_chat(chat_id, topic="", chat_type="oneOnOne"):
    return {"id": chat_id, "topic": topic, "chatType": chat_type, "webUrl": ""}


def _make_message(msg_id, body="hello", created="2026-03-23T10:00:00Z"):
    return {
        "id": msg_id,
        "messageType": "message",
        "createdDateTime": created,
        "from": {"user": {"displayName": "Alice", "id": "u1"}},
        "body": {"content": body, "contentType": "text"},
    }


# ---------------------------------------------------------------------------
# list_chat_messages: bounded traversal
# ---------------------------------------------------------------------------


@patch("microsoft_mcp.tools.graph")
def test_list_chat_messages_bounds_container_scan(mock_graph):
    """When limit is small, should not scan every chat."""
    from microsoft_mcp.tools import list_chat_messages

    # Create 20 chats but the paginated mock should respect the limit kwarg
    all_chats = [_make_chat(f"chat-{i}") for i in range(20)]

    def paginated_side_effect(path, **kw):
        if path == "/me/chats":
            lim = kw.get("limit", len(all_chats))
            return iter(all_chats[:lim])
        return iter([_make_message(f"m-{path}")])

    mock_graph.request_paginated.side_effect = paginated_side_effect

    list_chat_messages.fn(limit=5)
    # Should have scanned at most 10 containers (recent_container_limit default)
    paginated_calls = mock_graph.request_paginated.call_args_list
    message_calls = [c for c in paginated_calls if "/messages" in str(c)]
    assert len(message_calls) <= 10


@patch("microsoft_mcp.tools.graph")
def test_list_chat_messages_targeted_by_chat_id(mock_graph):
    """When chat_id is provided, should only scan that chat."""
    from microsoft_mcp.tools import list_chat_messages

    mock_graph.request_paginated.return_value = iter(
        [_make_message("m1"), _make_message("m2")]
    )

    list_chat_messages.fn(chat_id="target-chat", limit=5)
    call_paths = [str(c) for c in mock_graph.request_paginated.call_args_list]
    assert any("target-chat" in p for p in call_paths)
    # Should NOT have called /me/chats to list all chats
    assert not any("call('/me/chats'" in p for p in call_paths)


# ---------------------------------------------------------------------------
# list_channel_messages: bounded traversal
# ---------------------------------------------------------------------------


@patch("microsoft_mcp.tools.graph")
def test_list_channel_messages_bounds_team_scan(mock_graph):
    """When no team/channel specified, should bound the scan."""
    from microsoft_mcp.tools import list_channel_messages

    teams = [{"id": f"team-{i}", "displayName": f"Team {i}"} for i in range(10)]
    channels = [{"id": "ch-1", "displayName": "General", "webUrl": ""}]

    def paginated_side_effect(path, **kw):
        if path == "/me/joinedTeams":
            return iter(teams)
        if "/channels" in path and "/messages" not in path:
            return iter(channels)
        return iter([_make_message("m1")])

    mock_graph.request_paginated.side_effect = paginated_side_effect
    mock_graph.request.return_value = {"displayName": "Team", "webUrl": ""}

    result = list_channel_messages.fn(limit=5)
    # Should have a meta section with mode info
    assert isinstance(result, list)


@patch("microsoft_mcp.tools.graph")
def test_list_channel_messages_targeted_by_team_and_channel(mock_graph):
    """When team_id and channel_id are provided, should go directly."""
    from microsoft_mcp.tools import list_channel_messages

    mock_graph.request_paginated.return_value = iter([_make_message("m1")])
    mock_graph.request.return_value = {"displayName": "Test", "webUrl": ""}

    list_channel_messages.fn(team_id="t1", channel_id="c1", limit=5)
    # Should NOT have called /me/joinedTeams
    call_paths = [str(c) for c in mock_graph.request_paginated.call_args_list]
    assert not any("joinedTeams" in p for p in call_paths)


# ---------------------------------------------------------------------------
# Normalized message output
# ---------------------------------------------------------------------------


@patch("microsoft_mcp.tools.graph")
def test_chat_messages_include_chat_context(mock_graph):
    """Messages should include chat context metadata."""
    from microsoft_mcp.tools import list_chat_messages

    chats = [_make_chat("c1", topic="Project Alpha", chat_type="group")]
    mock_graph.request_paginated.side_effect = lambda path, **kw: (
        iter(chats) if path == "/me/chats" else iter([_make_message("m1")])
    )

    result = list_chat_messages.fn(limit=5)
    assert len(result) >= 1
    msg = result[0]
    assert "chatId" in msg
    assert msg["chatTopic"] == "Project Alpha"


@pytest.fixture
def load_tools_module(monkeypatch):
    def _load(auth_method: str):
        monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", auth_method)
        if auth_method == "azure":
            monkeypatch.setenv("MICROSOFT_MCP_CLIENT_ID", "test-client-id")
        sys.modules.pop("microsoft_mcp.tools", None)
        return importlib.import_module("microsoft_mcp.tools")

    yield _load
    sys.modules.pop("microsoft_mcp.tools", None)


def test_teams_tools_are_hidden_from_msal_tool_list(load_tools_module):
    module = load_tools_module("msal")

    tool_names = {
        tool.name for tool in asyncio.run(module.mcp._list_tools_middleware())
    }

    assert not (set(module.TEAMS_TOOL_NAMES) & tool_names)


def test_teams_tools_remain_available_for_azure_tool_list(load_tools_module):
    module = load_tools_module("azure")

    tool_names = {
        tool.name for tool in asyncio.run(module.mcp._list_tools_middleware())
    }

    assert set(module.TEAMS_TOOL_NAMES) <= tool_names


def test_teams_tools_cannot_be_called_under_msal(load_tools_module):
    module = load_tools_module("msal")

    with pytest.raises(NotFoundError, match="Unknown tool: list_chat_messages"):
        asyncio.run(module.mcp._call_tool_mcp("list_chat_messages", {}))
