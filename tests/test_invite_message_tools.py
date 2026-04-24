from unittest.mock import call, patch

import httpx


def _invite_message(invite_id: str, **overrides):
    raw = {
        "id": invite_id,
        "subject": f"Invite {invite_id}",
        "from": {
            "emailAddress": {
                "name": "Organizer",
                "address": "organizer@example.com",
            }
        },
        "toRecipients": [
            {
                "emailAddress": {
                    "name": "Invitee",
                    "address": "invitee@example.com",
                }
            }
        ],
        "receivedDateTime": "2026-04-02T10:00:00Z",
        "isRead": False,
        "conversationId": f"conv-{invite_id}",
        "bodyPreview": "Please join this meeting.",
        "meetingMessageType": "meetingRequest",
        "responseRequested": True,
        "allowNewTimeProposals": True,
        "isOutOfDate": False,
        "startDateTime": {"dateTime": "2026-04-03T17:00:00", "timeZone": "UTC"},
        "endDateTime": {"dateTime": "2026-04-03T18:00:00", "timeZone": "UTC"},
        "location": {"displayName": "Teams"},
        "webLink": f"https://outlook.office.com/mail/{invite_id}",
        "event": {
            "id": f"evt-{invite_id}",
            "subject": f"Invite {invite_id}",
            "start": {"dateTime": "2026-04-03T17:00:00", "timeZone": "UTC"},
            "end": {"dateTime": "2026-04-03T18:00:00", "timeZone": "UTC"},
            "location": {"displayName": "Teams"},
            "organizer": {
                "emailAddress": {
                    "name": "Organizer",
                    "address": "organizer@example.com",
                }
            },
        },
    }
    raw.update(overrides)
    return raw


@patch("microsoft_mcp.tools.graph")
def test_list_invite_messages_filters_event_messages(mock_graph):
    from microsoft_mcp.tools import list_invite_messages

    raw_messages = [
        {
            "id": "e-1",
            "subject": "Plain email",
            "receivedDateTime": "2026-04-02T09:00:00Z",
            "isRead": False,
            "conversationId": "conv-e-1",
        },
        {
            "id": "im-1",
            "subject": "Invite im-1",
            "from": {
                "emailAddress": {
                    "name": "Organizer",
                    "address": "organizer@example.com",
                }
            },
            "receivedDateTime": "2026-04-02T10:00:00Z",
            "isRead": False,
            "conversationId": "conv-im-1",
            "bodyPreview": "Please join this meeting.",
            "webLink": "https://outlook.office.com/mail/im-1",
        },
        {
            "id": "im-2",
            "subject": "Invite im-2",
            "from": {
                "emailAddress": {
                    "name": "Organizer",
                    "address": "organizer@example.com",
                }
            },
            "receivedDateTime": "2026-04-02T11:00:00Z",
            "isRead": False,
            "conversationId": "conv-im-2",
            "bodyPreview": "This meeting was canceled.",
            "webLink": "https://outlook.office.com/mail/im-2",
        },
    ]
    mock_graph.request_paginated.return_value = iter(raw_messages)
    mock_graph.request.side_effect = [
        {
            "id": "e-1",
            "subject": "Plain email",
            "@odata.type": "#microsoft.graph.message",
        },
        _invite_message("im-1"),
        _invite_message(
            "im-2",
            meetingMessageType="meetingCancelled",
            isOutOfDate=True,
        ),
    ]

    result = list_invite_messages.fn(limit=10)

    mock_graph.request_paginated.assert_called_once_with(
        "/me/mailFolders/inbox/messages",
        params={
            "$top": 50,
            "$select": (
                "id,subject,from,toRecipients,receivedDateTime,hasAttachments,"
                "bodyPreview,conversationId,isRead,webLink,flag"
            ),
            "$orderby": "receivedDateTime desc",
        },
        limit=50,
    )
    assert mock_graph.request.call_args_list == [
        call(
            "GET",
            "/me/messages/e-1",
            params={"$expand": "microsoft.graph.eventMessage/event"},
        ),
        call(
            "GET",
            "/me/messages/im-1",
            params={"$expand": "microsoft.graph.eventMessage/event"},
        ),
        call(
            "GET",
            "/me/messages/im-2",
            params={"$expand": "microsoft.graph.eventMessage/event"},
        ),
    ]
    assert [item["id"] for item in result] == ["im-1", "im-2"]
    assert result[0]["kind"] == "invite_message"
    assert result[0]["meeting_message_type"] == "meetingRequest"
    assert result[0]["event"]["id"] == "evt-im-1"
    assert result[1]["is_out_of_date"] is True


@patch("microsoft_mcp.tools.graph")
def test_delete_invite_message_removes_event_message(mock_graph):
    from microsoft_mcp.tools import delete_invite_message

    mock_graph.request.return_value = None

    result = delete_invite_message.fn(invite_message_id="im-3")

    mock_graph.request.assert_called_once_with("DELETE", "/me/messages/im-3")
    assert result == {
        "status": "deleted",
        "invite_message_id": "im-3",
        "resource": "eventMessage",
    }


@patch("microsoft_mcp.tools.graph")
def test_rsvp_to_invite_message_resolves_associated_event(mock_graph):
    from microsoft_mcp.tools import rsvp_to_invite_message

    mock_graph.request.side_effect = [
        _invite_message("im-4"),
        None,
    ]

    result = rsvp_to_invite_message.fn("im-4", response="accept")

    assert mock_graph.request.call_args_list == [
        call(
            "GET",
            "/me/messages/im-4",
            params={"$expand": "microsoft.graph.eventMessage/event"},
        ),
        call(
            "POST",
            "/me/events/evt-im-4/accept",
            json={"comment": None, "sendResponse": False},
        ),
    ]
    assert result == {
        "status": "responded",
        "invite_message_id": "im-4",
        "event_id": "evt-im-4",
        "meeting_message_type": "meetingRequest",
        "response": "accept",
        "send_response": False,
    }


@patch("microsoft_mcp.tools.graph")
def test_list_inbox_items_can_include_invite_messages(mock_graph):
    from microsoft_mcp.tools import list_inbox_items

    def paginated_side_effect(path, **kwargs):
        if "mailFolders" in path:
            return iter([_invite_message("im-5")])
        if "calendarView" in path:
            return iter([])
        return iter([])

    mock_graph.request_paginated.side_effect = paginated_side_effect
    mock_graph.request.return_value = _invite_message("im-5")

    result = list_inbox_items.fn(include_kinds=["invite_message"], limit=20)

    assert result["items"][0]["kind"] == "invite_message"
    assert result["items"][0]["id"] == "im-5"


@patch("microsoft_mcp.tools.graph")
def test_get_inbox_item_detail_hydrates_invite_message(mock_graph):
    from microsoft_mcp.tools import get_inbox_item_detail

    mock_graph.request.return_value = _invite_message(
        "im-6",
        body={"content": "<p>Join us</p>", "contentType": "html"},
    )

    result = get_inbox_item_detail.fn(item_id="im-6", kind="invite_message")

    assert result["kind"] == "invite_message"
    assert result["meeting_message_type"] == "meetingRequest"
    assert result["event"]["id"] == "evt-im-6"
    assert "body" in result


@patch("microsoft_mcp.tools.graph")
def test_list_inbox_items_keeps_email_and_event_results_when_invite_probe_fails(
    mock_graph,
):
    from microsoft_mcp.tools import list_inbox_items

    request = httpx.Request("GET", "https://graph.microsoft.com/v1.0/me/messages/e-1")
    response = httpx.Response(400, request=request)

    def paginated_side_effect(path, **kwargs):
        if "mailFolders" in path:
            return iter(
                [
                    {
                        "id": "e-1",
                        "subject": "Plain email",
                        "from": {
                            "emailAddress": {
                                "name": "Sender",
                                "address": "s@example.com",
                            }
                        },
                        "receivedDateTime": "2026-04-02T09:00:00Z",
                        "isRead": False,
                        "conversationId": "conv-e-1",
                        "bodyPreview": "Normal email",
                    }
                ]
            )
        if "calendarView" in path:
            return iter(
                [
                    {
                        "id": "evt-1",
                        "subject": "Meeting",
                        "start": {
                            "dateTime": "2026-04-03T17:00:00",
                            "timeZone": "UTC",
                        },
                        "end": {
                            "dateTime": "2026-04-03T18:00:00",
                            "timeZone": "UTC",
                        },
                        "location": {"displayName": "Room A"},
                        "organizer": {
                            "emailAddress": {
                                "name": "Org",
                                "address": "org@example.com",
                            }
                        },
                    }
                ]
            )
        return iter([])

    mock_graph.request_paginated.side_effect = paginated_side_effect
    mock_graph.request.side_effect = httpx.HTTPStatusError(
        "bad request",
        request=request,
        response=response,
    )

    result = list_inbox_items.fn(limit=20)

    assert {item["kind"] for item in result["items"]} == {"email", "event"}
