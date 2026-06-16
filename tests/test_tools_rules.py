from unittest.mock import patch
from src.microsoft_mcp.tools import list_inbox_rules


@patch("src.microsoft_mcp.tools.graph.request_paginated")
def test_list_inbox_rules_assistant_profile(mock_paged):
    mock_paged.return_value = iter(
        [
            {
                "id": "r1",
                "displayName": "News",
                "sequence": 1,
                "isEnabled": True,
                "conditions": {"senderContains": ["news"]},
                "actions": {"markAsRead": True},
            },
        ]
    )
    out = list_inbox_rules.fn(response_profile="assistant")
    assert out[0]["id"] == "r1"
    assert out[0]["display_name"] == "News"
    assert "news" in out[0]["conditions_summary"]
    called_path = mock_paged.call_args[0][0]
    assert called_path == "/me/mailFolders/inbox/messageRules"


# ---------------------------------------------------------------------------
# get_inbox_rule
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools.graph.request")
def test_get_inbox_rule_assistant(mock_req):
    from src.microsoft_mcp.tools import get_inbox_rule

    mock_req.return_value = {
        "id": "abc",
        "displayName": "My Rule",
        "sequence": 2,
        "isEnabled": True,
        "conditions": {"subjectContains": ["invoice"]},
        "actions": {"moveToFolder": "AAA"},
        "exceptions": None,
    }
    out = get_inbox_rule.fn(rule_id="abc", response_profile="assistant")
    mock_req.assert_called_once()
    args, kwargs = mock_req.call_args
    assert args[0] == "GET"
    assert "abc" in args[1]
    assert "/me/mailFolders/inbox/messageRules/" in args[1]
    assert out["id"] == "abc"
    assert out["display_name"] == "My Rule"
    assert "conditions_summary" in out
    assert "exceptions_summary" in out


@patch("src.microsoft_mcp.tools.graph.request")
def test_get_inbox_rule_legacy(mock_req):
    from src.microsoft_mcp.tools import get_inbox_rule

    mock_req.return_value = {
        "id": "xyz",
        "displayName": "Foo",
        "@odata.context": "should-be-stripped",
    }
    out = get_inbox_rule.fn(rule_id="xyz", response_profile="legacy")
    assert out["id"] == "xyz"
    assert "@odata.context" not in out


# ---------------------------------------------------------------------------
# create_inbox_rule
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools._resolve_mail_folder")
@patch("src.microsoft_mcp.tools.graph.request")
def test_create_inbox_rule_with_folder_name(mock_req, mock_resolve):
    from src.microsoft_mcp.tools import create_inbox_rule

    mock_resolve.return_value = "FOLDER_ID_123"
    mock_req.return_value = {
        "id": "new1",
        "displayName": "Archive Newsletter",
        "sequence": 5,
        "isEnabled": True,
        "conditions": {"senderContains": ["newsletter"]},
        "actions": {"moveToFolder": "FOLDER_ID_123"},
    }
    out = create_inbox_rule.fn(
        display_name="Archive Newsletter",
        sequence=5,
        sender_contains=["newsletter"],
        move_to_folder="Archive",
        response_profile="assistant",
    )
    mock_resolve.assert_called_once_with("Archive")
    mock_req.assert_called_once()
    args, kwargs = mock_req.call_args
    assert args[0] == "POST"
    assert args[1] == "/me/mailFolders/inbox/messageRules"
    payload = kwargs["json"]
    assert payload["actions"]["moveToFolder"] == "FOLDER_ID_123"
    assert out["id"] == "new1"
    assert out["display_name"] == "Archive Newsletter"


@patch("src.microsoft_mcp.tools._resolve_mail_folder")
@patch("src.microsoft_mcp.tools.graph.request")
def test_create_inbox_rule_no_folder(mock_req, mock_resolve):
    from src.microsoft_mcp.tools import create_inbox_rule

    mock_req.return_value = {
        "id": "new2",
        "displayName": "Mark Read",
        "sequence": 1,
        "isEnabled": True,
        "conditions": {},
        "actions": {"markAsRead": True},
    }
    create_inbox_rule.fn(
        display_name="Mark Read",
        mark_as_read=True,
        response_profile="legacy",
    )
    mock_resolve.assert_not_called()
    args, kwargs = mock_req.call_args
    assert args[0] == "POST"


# ---------------------------------------------------------------------------
# update_inbox_rule
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools.graph.request")
def test_update_inbox_rule_only_is_enabled(mock_req):
    from src.microsoft_mcp.tools import update_inbox_rule

    mock_req.return_value = {
        "id": "r1",
        "displayName": "News",
        "sequence": 1,
        "isEnabled": False,
    }
    out = update_inbox_rule.fn(
        rule_id="r1", is_enabled=False, response_profile="legacy"
    )
    args, kwargs = mock_req.call_args
    assert args[0] == "PATCH"
    assert "r1" in args[1]
    assert kwargs["json"] == {"isEnabled": False}
    assert out["id"] == "r1"


@patch("src.microsoft_mcp.tools.graph.request")
def test_update_inbox_rule_partial_conditions(mock_req):
    from src.microsoft_mcp.tools import update_inbox_rule

    mock_req.return_value = {
        "id": "r2",
        "displayName": "Updated",
        "sequence": 3,
        "isEnabled": True,
        "conditions": {"subjectContains": ["bill"]},
        "actions": {},
    }
    update_inbox_rule.fn(
        rule_id="r2", subject_contains=["bill"], response_profile="legacy"
    )
    _, kwargs = mock_req.call_args
    payload = kwargs["json"]
    assert "conditions" in payload
    assert payload["conditions"]["subjectContains"] == ["bill"]
    assert "actions" not in payload
    assert "displayName" not in payload


# ---------------------------------------------------------------------------
# delete_inbox_rule
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools.graph.request")
def test_delete_inbox_rule(mock_req):
    from src.microsoft_mcp.tools import delete_inbox_rule

    mock_req.return_value = None
    out = delete_inbox_rule.fn(rule_id="r99")
    args, kwargs = mock_req.call_args
    assert args[0] == "DELETE"
    assert "r99" in args[1]
    assert out == {"status": "deleted", "rule_id": "r99"}


# ---------------------------------------------------------------------------
# toggle_inbox_rule
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools.graph.request")
def test_toggle_inbox_rule_true_to_false(mock_req):
    from src.microsoft_mcp.tools import toggle_inbox_rule

    mock_req.side_effect = [
        {"id": "r5", "isEnabled": True},
        {"id": "r5", "isEnabled": False},
    ]
    out = toggle_inbox_rule.fn(rule_id="r5")
    assert mock_req.call_count == 2
    get_args = mock_req.call_args_list[0][0]
    patch_args = mock_req.call_args_list[1][0]
    patch_kwargs = mock_req.call_args_list[1][1]
    assert get_args[0] == "GET"
    assert patch_args[0] == "PATCH"
    assert "r5" in patch_args[1]
    assert patch_kwargs["json"] == {"isEnabled": False}
    assert out == {"rule_id": "r5", "is_enabled": False}


@patch("src.microsoft_mcp.tools.graph.request")
def test_toggle_inbox_rule_false_to_true(mock_req):
    from src.microsoft_mcp.tools import toggle_inbox_rule

    mock_req.side_effect = [
        {"id": "r6", "isEnabled": False},
        {"id": "r6", "isEnabled": True},
    ]
    out = toggle_inbox_rule.fn(rule_id="r6")
    patch_kwargs = mock_req.call_args_list[1][1]
    assert patch_kwargs["json"] == {"isEnabled": True}
    assert out == {"rule_id": "r6", "is_enabled": True}
