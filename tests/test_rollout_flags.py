"""Tests for MICROSOFT_MCP_RESPONSE_PROFILE rollout flag."""

from unittest.mock import patch


def _make_raw_email(email_id, subject="Test", body_preview="Hello..."):
    """Build a realistic raw Graph email payload."""
    return {
        "id": email_id,
        "subject": subject,
        "from": {"emailAddress": {"name": "Sender", "address": "sender@x.com"}},
        "toRecipients": [{"emailAddress": {"address": "me@x.com"}}],
        "receivedDateTime": "2026-03-23T10:00:00Z",
        "isRead": False,
        "hasAttachments": False,
        "bodyPreview": body_preview,
        "conversationId": "conv-1",
    }


# ---------------------------------------------------------------------------
# Profile flag controls list_emails output
# ---------------------------------------------------------------------------


def _make_raw_email_with_body(email_id):
    """Build a raw Graph email with body content included."""
    raw = _make_raw_email(email_id)
    raw["body"] = {"content": "Full email body here", "contentType": "text"}
    return raw


@patch("microsoft_mcp.tools.graph")
def test_assistant_profile_suppresses_body_even_when_requested(mock_graph, monkeypatch):
    """When MICROSOFT_MCP_RESPONSE_PROFILE=assistant, list_emails suppresses
    body even if include_body=True is passed."""
    monkeypatch.setenv("MICROSOFT_MCP_RESPONSE_PROFILE", "assistant")
    from microsoft_mcp.tools import list_emails

    mock_graph.request_paginated.return_value = iter([_make_raw_email("e-1")])

    # Explicitly request body, but assistant profile overrides to summary
    result = list_emails.fn(limit=5, include_body=True)
    assert "body" not in result[0]


@patch("microsoft_mcp.tools.graph")
def test_legacy_profile_allows_body_when_requested(mock_graph, monkeypatch):
    """Default legacy profile respects include_body=True."""
    monkeypatch.delenv("MICROSOFT_MCP_RESPONSE_PROFILE", raising=False)
    from microsoft_mcp.tools import list_emails

    mock_graph.request_paginated.return_value = iter(
        [_make_raw_email_with_body("e-2")]
    )

    result = list_emails.fn(limit=5, include_body=True)
    assert "body" in result[0]


@patch("microsoft_mcp.tools.graph")
def test_response_profile_parameter_overrides_env(mock_graph, monkeypatch):
    """An explicit response_profile='assistant' overrides legacy env var."""
    monkeypatch.setenv("MICROSOFT_MCP_RESPONSE_PROFILE", "legacy")
    from microsoft_mcp.tools import list_emails

    mock_graph.request_paginated.return_value = iter([_make_raw_email("e-3")])

    # Even with legacy env, parameter wins
    result = list_emails.fn(limit=5, include_body=True, response_profile="assistant")
    assert "body" not in result[0]
    assert result[0]["from"] == "Sender <sender@x.com>"


# ---------------------------------------------------------------------------
# Profile flag affects list_events
# ---------------------------------------------------------------------------


def _make_event_with_body():
    return {
        "id": "evt-1",
        "subject": "Standup",
        "start": {"dateTime": "2026-03-24T09:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-03-24T09:30:00", "timeZone": "UTC"},
        "location": {"displayName": "Room A"},
        "organizer": {
            "emailAddress": {"name": "Boss", "address": "boss@x.com"}
        },
        "body": {"content": "<p>Standup details</p>", "contentType": "html"},
        "attendees": [
            {"emailAddress": {"name": "Dev", "address": "dev@x.com"}, "status": {"response": "accepted"}}
        ],
    }


@patch("microsoft_mcp.tools.graph")
def test_list_events_assistant_suppresses_details(mock_graph, monkeypatch):
    """Assistant profile forces include_details=False even when True is passed."""
    monkeypatch.setenv("MICROSOFT_MCP_RESPONSE_PROFILE", "assistant")
    from microsoft_mcp.tools import list_events

    mock_graph.request_paginated.return_value = iter([_make_event_with_body()])

    result = list_events.fn(include_details=True, response_profile="assistant")
    evt = result[0]
    assert "body" not in evt
    assert evt["organizer"] == "Boss <boss@x.com>"


# ---------------------------------------------------------------------------
# Profile flag affects list_contacts
# ---------------------------------------------------------------------------


@patch("microsoft_mcp.tools.graph")
def test_list_contacts_respects_assistant_profile(mock_graph, monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_RESPONSE_PROFILE", "assistant")
    from microsoft_mcp.tools import list_contacts

    mock_graph.request_paginated.return_value = iter(
        [
            {
                "id": "c-1",
                "displayName": "Alice",
                "emailAddresses": [{"address": "alice@x.com"}],
                "businessPhones": ["+1234"],
            }
        ]
    )

    result = list_contacts.fn(limit=5, response_profile="assistant")
    first = result[0]
    assert "body" not in first
    assert first["displayName"] == "Alice"


# ---------------------------------------------------------------------------
# get_response_profile helper
# ---------------------------------------------------------------------------


def test_get_response_profile_returns_env_value(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_RESPONSE_PROFILE", "assistant")
    from microsoft_mcp.tools import get_response_profile

    assert get_response_profile() == "assistant"


def test_get_response_profile_defaults_to_legacy(monkeypatch):
    monkeypatch.delenv("MICROSOFT_MCP_RESPONSE_PROFILE", raising=False)
    from microsoft_mcp.tools import get_response_profile

    assert get_response_profile() == "legacy"
