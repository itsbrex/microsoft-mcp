from unittest.mock import patch


@patch("microsoft_mcp.tools.graph")
def test_list_events_returns_shaped_summaries(mock_graph):
    from microsoft_mcp.tools import list_events

    mock_graph.request_paginated.return_value = iter(
        [
            {
                "id": "evt-1",
                "subject": "Standup",
                "start": {"dateTime": "2026-03-24T09:00:00", "timeZone": "UTC"},
                "end": {"dateTime": "2026-03-24T09:30:00", "timeZone": "UTC"},
                "location": {"displayName": "Teams"},
                "organizer": {"emailAddress": {"name": "Boss", "address": "boss@x.com"}},
                "@odata.etag": "junk",
                "seriesMasterId": "series-1",
            }
        ]
    )

    result = list_events.fn()
    evt = result[0]
    assert evt["id"] == "evt-1"
    assert evt["location"] == "Teams"
    assert evt["organizer"] == "Boss <boss@x.com>"
    assert "@odata.etag" not in evt


@patch("microsoft_mcp.tools.graph")
def test_get_event_extracts_teams_meeting_info(mock_graph):
    from microsoft_mcp.tools import get_event

    mock_graph.request.return_value = {
        "id": "evt-2",
        "subject": "Design Review",
        "start": {"dateTime": "2026-03-24T14:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-03-24T15:00:00", "timeZone": "UTC"},
        "location": {"displayName": "Teams"},
        "organizer": {"emailAddress": {"name": "Org", "address": "org@x.com"}},
        "body": {
            "contentType": "html",
            "content": '<a href="https://teams.microsoft.com/l/meetup-join/abc123">Join</a>',
        },
        "onlineMeeting": {"joinUrl": "https://teams.microsoft.com/l/meetup-join/abc123"},
        "attendees": [
            {
                "emailAddress": {"name": "A", "address": "a@x.com"},
                "status": {"response": "accepted"},
            }
        ],
        "@odata.context": "junk",
    }

    result = get_event.fn("evt-2")
    assert "meeting" in result
    assert result["meeting"]["join_url"] == "https://teams.microsoft.com/l/meetup-join/abc123"
    assert "@odata.context" not in result
    assert result["attendees"][0]["name"] == "A <a@x.com>"
    assert result["attendees"][0]["status"] == "accepted"
