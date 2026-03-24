"""Token budget assertions for shaped tool responses.

These tests verify that shaped outputs stay within reasonable size budgets,
preventing context-window bloat when tools are called by AI assistants.
"""

import json
from unittest.mock import patch


# ---------------------------------------------------------------------------
# Helpers: realistic mock data generators
# ---------------------------------------------------------------------------


def _make_raw_email(i):
    """Realistic email payload from Graph API."""
    return {
        "id": f"AAMkAGE{i:04d}",
        "subject": f"Re: Q{(i % 4) + 1} Budget Review follow-up discussion thread #{i}",
        "from": {
            "emailAddress": {
                "name": f"Person {i} McLastname",
                "address": f"person{i}@longdomainname-corporation.com",
            }
        },
        "toRecipients": [
            {"emailAddress": {"address": "me@longdomainname-corporation.com"}},
            {"emailAddress": {"address": f"cc{i}@longdomainname-corporation.com"}},
        ],
        "receivedDateTime": f"2026-03-{10 + (i % 20):02d}T{8 + (i % 12):02d}:30:00Z",
        "isRead": i % 3 == 0,
        "hasAttachments": i % 4 == 0,
        "bodyPreview": f"Hi team, following up on our earlier conversation about the Q{(i % 4) + 1} budget allocations. "
        f"We need to finalize the numbers by end of week. Please review the attached spreadsheet "
        f"and provide your feedback. Best regards, Person {i}",
        "conversationId": f"AAQkAGE{i:04d}conv",
    }


def _make_raw_event(i):
    """Realistic calendar event payload from Graph API."""
    return {
        "id": f"AAMkAGEvt{i:04d}",
        "subject": f"{'Weekly Standup' if i % 3 == 0 else 'Design Review' if i % 3 == 1 else '1:1 with Manager'} - Sprint {i}",
        "start": {
            "dateTime": f"2026-03-{24 + (i % 5):02d}T{9 + i}:00:00",
            "timeZone": "UTC",
        },
        "end": {
            "dateTime": f"2026-03-{24 + (i % 5):02d}T{10 + i}:00:00",
            "timeZone": "UTC",
        },
        "location": {
            "displayName": f"Conference Room {'Alpha' if i % 2 == 0 else 'Beta'}"
        },
        "organizer": {
            "emailAddress": {
                "name": f"Organizer {i}",
                "address": f"org{i}@longdomainname-corporation.com",
            }
        },
        "isAllDay": False,
        "seriesMasterId": f"series-{i % 3}" if i % 3 == 0 else None,
    }


def _make_raw_contact(i):
    """Realistic contact payload from Graph API."""
    return {
        "id": f"AAMkAGCt{i:04d}",
        "@odata.etag": f'W/"etag{i}"',
        "displayName": f"{'Alice' if i % 3 == 0 else 'Bob' if i % 3 == 1 else 'Charlie'} {'Smith' if i % 2 == 0 else 'Johnson'} {i}",
        "jobTitle": f"{'Engineer' if i % 3 == 0 else 'Manager' if i % 3 == 1 else 'Director'}",
        "companyName": f"{'Acme Corp' if i % 2 == 0 else 'Globex Inc'}",
        "emailAddresses": [
            {"address": f"contact{i}@longdomainname-corporation.com"},
            {"address": f"contact{i}.personal@email.com"},
        ],
        "businessPhones": [f"+1-555-{1000 + i:04d}"],
        "mobilePhone": f"+1-555-{2000 + i:04d}",
        "homePhones": [],
        "personalNotes": None,
        "changeKey": f"changekey{i}",
    }


def _make_raw_chat_message(i):
    """Realistic Teams chat message payload from Graph API."""
    return {
        "id": f"msg-{i:04d}",
        "messageType": "message",
        "createdDateTime": f"2026-03-{10 + (i % 20):02d}T{8 + (i % 12):02d}:15:00Z",
        "from": {"user": {"displayName": f"Team Member {i}", "id": f"user-{i}"}},
        "body": {
            "content": f"Hey team, I wanted to share an update on the sprint {i} deliverables. "
            f"We have completed the backend integration and are now moving to testing phase. "
            f"Please review the PR #{1000 + i} when you get a chance.",
            "contentType": "text",
        },
        "chatId": f"chat-{i % 3}",
        "chatTopic": f"Project {'Alpha' if i % 3 == 0 else 'Beta' if i % 3 == 1 else 'Gamma'}",
        "chatType": "group",
        "webUrl": "",
    }


# ---------------------------------------------------------------------------
# Budget tests
# ---------------------------------------------------------------------------


@patch("microsoft_mcp.tools.graph")
def test_list_emails_summary_stays_under_budget(mock_graph):
    """list_emails(limit=10) summary output must stay under 12k chars."""
    from microsoft_mcp.tools import list_emails

    mock_graph.request_paginated.return_value = iter(
        [_make_raw_email(i) for i in range(10)]
    )

    result = list_emails.fn(limit=10)
    serialized = json.dumps(result)
    assert len(serialized) < 12_000, (
        f"list_emails(limit=10) produced {len(serialized)} chars, budget is 12000"
    )


@patch("microsoft_mcp.tools.graph")
def test_list_events_summary_stays_under_budget(mock_graph):
    """list_events summary output for 10 events must stay under 8k chars."""
    from microsoft_mcp.tools import list_events

    mock_graph.request_paginated.return_value = iter(
        [_make_raw_event(i) for i in range(10)]
    )

    result = list_events.fn()
    serialized = json.dumps(result)
    assert len(serialized) < 8_000, (
        f"list_events produced {len(serialized)} chars, budget is 8000"
    )


@patch("microsoft_mcp.tools.graph")
def test_list_contacts_summary_stays_under_budget(mock_graph):
    """list_contacts(limit=20) summary output must stay under 10k chars."""
    from microsoft_mcp.tools import list_contacts

    mock_graph.request_paginated.return_value = iter(
        [_make_raw_contact(i) for i in range(20)]
    )

    result = list_contacts.fn(limit=20)
    serialized = json.dumps(result)
    assert len(serialized) < 10_000, (
        f"list_contacts(limit=20) produced {len(serialized)} chars, budget is 10000"
    )


@patch("microsoft_mcp.tools.graph")
def test_list_chat_messages_summary_stays_under_budget(mock_graph):
    """list_chat_messages(limit=10) summary output must stay under 12k chars."""
    from microsoft_mcp.tools import list_chat_messages

    # For chat messages, we provide a chat_id to avoid the container scan
    # and directly supply messages
    mock_graph.request_paginated.return_value = iter(
        [_make_raw_chat_message(i) for i in range(10)]
    )

    result = list_chat_messages.fn(chat_id="test-chat", limit=10)
    serialized = json.dumps(result)
    assert len(serialized) < 12_000, (
        f"list_chat_messages(limit=10) produced {len(serialized)} chars, budget is 12000"
    )
