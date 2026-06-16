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


# ---------------------------------------------------------------------------
# reorder_inbox_rules
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools.graph.request")
def test_reorder_inbox_rules(mock_req):
    from src.microsoft_mcp.tools import reorder_inbox_rules

    mock_req.return_value = {}
    out = reorder_inbox_rules.fn(rule_ids_in_order=["rB", "rA"])

    assert mock_req.call_count == 2
    call0_args, call0_kwargs = mock_req.call_args_list[0]
    call1_args, call1_kwargs = mock_req.call_args_list[1]

    assert call0_args[0] == "PATCH"
    assert "rB" in call0_args[1]
    assert call0_kwargs["json"] == {"sequence": 1}

    assert call1_args[0] == "PATCH"
    assert "rA" in call1_args[1]
    assert call1_kwargs["json"] == {"sequence": 2}

    assert out == [
        {"rule_id": "rB", "sequence": 1},
        {"rule_id": "rA", "sequence": 2},
    ]


# ---------------------------------------------------------------------------
# export_inbox_rules
# ---------------------------------------------------------------------------

_MOCK_RULES = [
    {
        "id": "r1",
        "displayName": "Newsletter",
        "sequence": 1,
        "isEnabled": True,
        "conditions": {"senderContains": ["newsletter"]},
        "actions": {"moveToFolder": "FOLDER_ID_NEWSLETTERS"},
    }
]

_MOCK_FOLDERS = [
    {"id": "FOLDER_ID_NEWSLETTERS", "displayName": "Newsletters"},
    {"id": "FOLDER_ID_ARCHIVE", "displayName": "Archive"},
]


@patch("src.microsoft_mcp.tools.graph.request_paginated")
def test_export_inbox_rules_yaml_contains_name_and_folder_name(mock_paged):
    from src.microsoft_mcp.tools import export_inbox_rules
    import yaml

    # First call: rules, second call: folders
    mock_paged.side_effect = [iter(_MOCK_RULES), iter(_MOCK_FOLDERS)]

    out = export_inbox_rules.fn()
    assert "yaml" in out
    assert out["count"] == 1

    data = yaml.safe_load(out["yaml"])
    assert len(data["rules"]) == 1
    rule = data["rules"][0]
    assert rule["name"] == "Newsletter"
    # folder ID must be resolved to display name
    assert rule["actions"]["move_to"] == "Newsletters"


@patch("src.microsoft_mcp.tools.graph.request_paginated")
def test_export_inbox_rules_writes_file(mock_paged, tmp_path):
    from src.microsoft_mcp.tools import export_inbox_rules

    mock_paged.side_effect = [iter(_MOCK_RULES), iter(_MOCK_FOLDERS)]
    out_path = str(tmp_path / "rules.yaml")
    out = export_inbox_rules.fn(path=out_path)
    assert out == {"path": out_path, "count": 1}
    assert (tmp_path / "rules.yaml").exists()


# ---------------------------------------------------------------------------
# import_inbox_rules
# ---------------------------------------------------------------------------

_IMPORT_YAML = """\
rules:
  - name: Newsletter
    enabled: true
    sequence: 1
    conditions:
      sender_contains:
        - newsletter
    actions:
      move_to: Newsletters
"""


@patch("src.microsoft_mcp.tools._resolve_mail_folder")
@patch("src.microsoft_mcp.tools.graph.request_paginated")
@patch("src.microsoft_mcp.tools.graph.request")
def test_import_inbox_rules_dry_run_no_mutations(mock_req, mock_paged, mock_resolve):
    from src.microsoft_mcp.tools import import_inbox_rules

    # Existing rules (no match) returned by paginated
    mock_paged.return_value = iter([])
    mock_resolve.return_value = "FOLDER_ID_NEWSLETTERS"

    out = import_inbox_rules.fn(yaml_text=_IMPORT_YAML, dry_run=True)

    # dry_run must not POST or PATCH
    for call in mock_req.call_args_list:
        method = call[0][0] if call[0] else call[1].get("method", "")
        assert method not in ("POST", "PATCH"), f"unexpected mutating call: {call}"

    assert len(out["created"]) == 1
    assert out["created"][0] == "Newsletter"
    assert out["updated"] == []
    assert out["skipped"] == []
    assert out["errors"] == []


@patch("src.microsoft_mcp.tools._resolve_mail_folder")
@patch("src.microsoft_mcp.tools.graph.request_paginated")
@patch("src.microsoft_mcp.tools.graph.request")
def test_import_inbox_rules_create_skips_existing_name(
    mock_req, mock_paged, mock_resolve
):
    from src.microsoft_mcp.tools import import_inbox_rules

    existing_rule = {
        "id": "r1",
        "displayName": "Newsletter",
        "sequence": 1,
        "isEnabled": True,
        "conditions": {"senderContains": ["newsletter"]},
        "actions": {"moveToFolder": "FOLDER_ID_NEWSLETTERS"},
    }
    mock_paged.return_value = iter([existing_rule])
    mock_resolve.return_value = "FOLDER_ID_NEWSLETTERS"

    out = import_inbox_rules.fn(yaml_text=_IMPORT_YAML, mode="create")

    # must not POST or PATCH
    mock_req.assert_not_called()
    assert out["skipped"] == ["Newsletter"]
    assert out["created"] == []
    assert out["errors"] == []


@patch("src.microsoft_mcp.tools._resolve_mail_folder")
@patch("src.microsoft_mcp.tools.graph.request_paginated")
@patch("src.microsoft_mcp.tools.graph.request")
def test_import_inbox_rules_invalid_mode_raises(mock_req, mock_paged, mock_resolve):
    from src.microsoft_mcp.tools import import_inbox_rules
    import pytest

    with pytest.raises(ValueError, match="mode"):
        import_inbox_rules.fn(yaml_text=_IMPORT_YAML, mode="replace")
