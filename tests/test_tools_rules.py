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
