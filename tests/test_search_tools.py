from unittest.mock import patch


@patch("microsoft_mcp.tools.graph")
def test_unified_search_defaults_to_inbox_entities(mock_graph):
    from microsoft_mcp.tools import unified_search

    mock_graph.request.return_value = {"value": []}

    result = unified_search.fn("AI pilot")
    assert result["summary"]["entity_types_searched"] == [
        "message",
        "event",
        "chatMessage",
    ]


@patch("microsoft_mcp.tools.graph")
def test_unified_search_results_are_normalized(mock_graph):
    from microsoft_mcp.tools import unified_search

    mock_graph.request.return_value = {
        "value": [
            {
                "hitsContainers": [
                    {
                        "total": 1,
                        "hits": [
                            {
                                "rank": 1,
                                "summary": "Budget meeting tomorrow",
                                "resource": {
                                    "@odata.type": "#microsoft.graph.message",
                                    "id": "msg-1",
                                    "subject": "Budget Q4",
                                    "from": {
                                        "emailAddress": {
                                            "name": "JP",
                                            "address": "jp@x.com",
                                        }
                                    },
                                    "receivedDateTime": "2026-03-23T10:00:00Z",
                                    "conversationId": "conv-1",
                                },
                            }
                        ],
                    }
                ]
            }
        ]
    }

    result = unified_search.fn("budget", entity_types=["message"])
    hit = result["results"][0]
    assert "id" in hit
    assert "kind" in hit
    assert "title" in hit
    assert "snippet" in hit
    assert "score" in hit
    assert hit["kind"] == "message"


@patch("microsoft_mcp.tools.graph")
def test_unified_search_uses_hit_id_when_resource_id_is_missing(mock_graph):
    from microsoft_mcp.tools import unified_search

    mock_graph.request.return_value = {
        "value": [
            {
                "hitsContainers": [
                    {
                        "total": 1,
                        "hits": [
                            {
                                "hitId": "msg-hit-1",
                                "rank": 1,
                                "summary": "Budget meeting tomorrow",
                                "resource": {
                                    "@odata.type": "#microsoft.graph.message",
                                    "subject": "Budget Q4",
                                    "from": {
                                        "emailAddress": {
                                            "name": "JP",
                                            "address": "jp@x.com",
                                        }
                                    },
                                    "receivedDateTime": "2026-03-23T10:00:00Z",
                                    "conversationId": "conv-1",
                                },
                            }
                        ],
                    }
                ]
            }
        ]
    }

    result = unified_search.fn("budget", entity_types=["message"])

    assert result["results"][0]["id"] == "msg-hit-1"


@patch("microsoft_mcp.tools.graph")
def test_search_emails_returns_normalized_shape(mock_graph):
    from microsoft_mcp.tools import search_emails

    mock_graph.search_query.return_value = iter(
        [
            {
                "id": "e-1",
                "subject": "Meeting",
                "from": {"emailAddress": {"name": "A", "address": "a@x.com"}},
                "receivedDateTime": "2026-03-23T10:00:00Z",
                "isRead": False,
                "hasAttachments": False,
                "conversationId": "conv-1",
            }
        ]
    )

    result = search_emails.fn("meeting", limit=1)
    first = result[0]
    assert first["from"] == "A <a@x.com>"
    assert "conversation_url" in first


@patch("microsoft_mcp.tools.graph")
def test_search_emails_resolves_custom_folder_name(mock_graph):
    from microsoft_mcp.tools import search_emails

    mock_graph.request_paginated.side_effect = [
        iter(
            [
                {
                    "id": "folder-1",
                    "displayName": "Cresa Deals of the Week",
                    "childFolderCount": 0,
                    "totalItemCount": 12,
                    "unreadItemCount": 2,
                }
            ]
        ),
        iter(
            [
                {
                    "id": "e-1",
                    "subject": "Cresa Deals of the Week 3/13/26",
                    "from": {
                        "emailAddress": {
                            "name": "Cresa Communications",
                            "address": "communications@cresa.com",
                        }
                    },
                    "receivedDateTime": "2026-03-13T15:01:33Z",
                    "isRead": False,
                    "conversationId": "conv-1",
                }
            ]
        ),
    ]

    result = search_emails.fn(
        "Cresa Deals of the Week", limit=1, folder="Cresa Deals of the Week"
    )

    assert (
        mock_graph.request_paginated.call_args_list[1][0][0]
        == "/me/mailFolders/folder-1/messages"
    )
    assert result[0]["id"] == "e-1"


@patch("microsoft_mcp.tools.graph")
def test_search_events_returns_normalized_shape(mock_graph):
    from microsoft_mcp.tools import search_events

    mock_graph.search_query.return_value = iter(
        [
            {
                "id": "evt-1",
                "subject": "Design Review",
                "start": {"dateTime": "2026-03-24T14:00:00", "timeZone": "UTC"},
                "end": {"dateTime": "2026-03-24T15:00:00", "timeZone": "UTC"},
                "location": {"displayName": "Teams"},
                "organizer": {"emailAddress": {"name": "Org", "address": "org@x.com"}},
            }
        ]
    )

    result = search_events.fn("design", limit=1)
    first = result[0]
    assert first["id"] == "evt-1"
    assert first["location"] == "Teams"
    assert first["organizer"] == "Org <org@x.com>"


@patch("microsoft_mcp.tools.graph")
def test_search_contacts_returns_shaped_results(mock_graph):
    from microsoft_mcp.tools import search_contacts

    mock_graph.request_paginated.return_value = iter(
        [
            {
                "id": "c-1",
                "displayName": "John",
                "emailAddresses": [{"address": "john@x.com"}],
                "@odata.etag": "junk",
                "changeKey": "abc",
            }
        ]
    )

    result = search_contacts.fn("john", limit=1)
    first = result[0]
    assert first["email_addresses"] == ["john@x.com"]
    assert "@odata.etag" not in first
