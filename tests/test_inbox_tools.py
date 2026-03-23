from unittest.mock import patch


def _make_emails(n=3):
    return [
        {
            "id": f"e-{i}",
            "subject": f"Email {i}",
            "from": {"emailAddress": {"name": f"User{i}", "address": f"u{i}@x.com"}},
            "receivedDateTime": f"2026-03-23T{10 + i}:00:00Z",
            "isRead": i % 2 == 0,
            "hasAttachments": False,
            "conversationId": f"conv-{i}",
        }
        for i in range(n)
    ]


def _make_events(n=2):
    return [
        {
            "id": f"evt-{i}",
            "subject": f"Meeting {i}",
            "start": {"dateTime": f"2026-03-23T{14 + i}:00:00", "timeZone": "UTC"},
            "end": {"dateTime": f"2026-03-23T{15 + i}:00:00", "timeZone": "UTC"},
            "location": {"displayName": "Room A"},
            "organizer": {"emailAddress": {"name": "Org", "address": "org@x.com"}},
        }
        for i in range(n)
    ]


@patch("microsoft_mcp.tools.graph")
def test_list_inbox_items_returns_mixed_ranked_items(mock_graph):
    from microsoft_mcp.tools import list_inbox_items

    emails = _make_emails(3)
    events = _make_events(2)

    def paginated_side_effect(path, **kw):
        if "mailFolders" in path:
            return iter(emails)
        if "calendarView" in path:
            return iter(events)
        return iter([])

    mock_graph.request_paginated.side_effect = paginated_side_effect

    result = list_inbox_items.fn(limit=20)
    assert "items" in result
    assert "meta" in result
    assert len(result["items"]) > 0
    first = result["items"][0]
    assert {"id", "kind", "title", "score"} <= set(first)


@patch("microsoft_mcp.tools.graph")
def test_list_inbox_items_respects_limit(mock_graph):
    from microsoft_mcp.tools import list_inbox_items

    emails = _make_emails(5)

    def paginated_side_effect(path, **kw):
        if "mailFolders" in path:
            return iter(emails)
        return iter([])

    mock_graph.request_paginated.side_effect = paginated_side_effect

    result = list_inbox_items.fn(limit=3)
    assert len(result["items"]) <= 3


@patch("microsoft_mcp.tools.graph")
def test_list_inbox_items_filters_by_kind(mock_graph):
    from microsoft_mcp.tools import list_inbox_items

    emails = _make_emails(3)
    events = _make_events(2)

    def paginated_side_effect(path, **kw):
        if "mailFolders" in path:
            return iter(emails)
        if "calendarView" in path:
            return iter(events)
        return iter([])

    mock_graph.request_paginated.side_effect = paginated_side_effect

    result = list_inbox_items.fn(include_kinds=["event"], limit=20)
    for item in result["items"]:
        assert item["kind"] == "event"


@patch("microsoft_mcp.tools.graph")
def test_get_inbox_item_detail_hydrates_email(mock_graph):
    from microsoft_mcp.tools import get_inbox_item_detail

    mock_graph.request.return_value = {
        "id": "e-1",
        "subject": "Test Email",
        "from": {"emailAddress": {"name": "A", "address": "a@x.com"}},
        "receivedDateTime": "2026-03-23T10:00:00Z",
        "isRead": False,
        "hasAttachments": False,
        "conversationId": "conv-1",
        "body": {"content": "Hello world", "contentType": "text"},
    }

    result = get_inbox_item_detail.fn(item_id="e-1", kind="email")
    assert result["kind"] == "email"
    assert "body" in result


@patch("microsoft_mcp.tools.graph")
def test_get_inbox_item_detail_hydrates_event(mock_graph):
    from microsoft_mcp.tools import get_inbox_item_detail

    mock_graph.request.return_value = {
        "id": "evt-1",
        "subject": "Meeting",
        "start": {"dateTime": "2026-03-23T14:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-03-23T15:00:00", "timeZone": "UTC"},
        "location": {"displayName": "Room A"},
        "organizer": {"emailAddress": {"name": "Org", "address": "org@x.com"}},
        "body": {"content": "Agenda here", "contentType": "text"},
        "attendees": [],
    }

    result = get_inbox_item_detail.fn(item_id="evt-1", kind="event")
    assert result["kind"] == "event"
    assert "body" in result
