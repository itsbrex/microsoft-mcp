from microsoft_mcp.response_shaping import (
    cleanup_graph_payload,
    flatten_email_address,
    shape_email_summary,
    shape_email_detail,
    shape_event_summary,
    shape_event_detail,
    shape_contact_summary,
    shape_contact_detail,
    shape_message_summary,
)


# --- Task 2: cleanup_graph_payload ---


def test_cleanup_graph_payload_strips_odata_and_empty_values():
    raw = {
        "@odata.context": "x",
        "@odata.etag": "y",
        "displayName": "John",
        "mobilePhone": None,
        "otherAddress": {},
        "businessPhones": [],
    }
    assert cleanup_graph_payload(raw) == {"displayName": "John"}


def test_cleanup_graph_payload_keeps_false_and_zero():
    raw = {"isRead": False, "size": 0, "subject": "Test"}
    assert cleanup_graph_payload(raw) == raw


def test_cleanup_graph_payload_strips_nested_odata():
    raw = {
        "value": [
            {"@odata.type": "#microsoft.graph.message", "id": "1", "extra": None},
            {"id": "2", "tags": []},
        ]
    }
    assert cleanup_graph_payload(raw) == {"value": [{"id": "1"}, {"id": "2"}]}


def test_cleanup_graph_payload_strips_noise_keys():
    raw = {
        "id": "1",
        "changeKey": "abc",
        "parentFolderId": "xyz",
        "subject": "Hello",
    }
    assert cleanup_graph_payload(raw) == {"id": "1", "subject": "Hello"}


# --- Task 3: Entity-specific shapers ---


def test_flatten_email_address():
    nested = {"emailAddress": {"name": "JP", "address": "jp@example.com"}}
    assert flatten_email_address(nested) == "JP <jp@example.com>"


def test_flatten_email_address_no_name():
    nested = {"emailAddress": {"address": "jp@example.com"}}
    assert flatten_email_address(nested) == "jp@example.com"


def test_shape_email_summary_drops_body_and_flattens_sender():
    raw = {
        "id": "1",
        "subject": "Hello",
        "from": {"emailAddress": {"name": "JP", "address": "jp@example.com"}},
        "body": {"content": "huge"},
        "conversationId": "abc",
        "receivedDateTime": "2026-03-23T10:00:00Z",
        "isRead": False,
        "hasAttachments": False,
    }
    shaped = shape_email_summary(raw)
    assert shaped["id"] == "1"
    assert shaped["subject"] == "Hello"
    assert shaped["from"] == "JP <jp@example.com>"
    assert "body" not in shaped
    assert "conversation_url" in shaped


def test_shape_email_summary_flattens_recipients():
    raw = {
        "id": "2",
        "subject": "Re: Hello",
        "from": {"emailAddress": {"name": "A", "address": "a@x.com"}},
        "toRecipients": [
            {"emailAddress": {"name": "B", "address": "b@x.com"}},
            {"emailAddress": {"address": "c@x.com"}},
        ],
        "receivedDateTime": "2026-03-23T10:00:00Z",
        "isRead": True,
        "hasAttachments": True,
        "conversationId": "conv-2",
    }
    shaped = shape_email_summary(raw)
    assert shaped["to"] == ["B <b@x.com>", "c@x.com"]


def test_shape_email_detail_includes_body():
    raw = {
        "id": "1",
        "subject": "Hello",
        "from": {"emailAddress": {"name": "JP", "address": "jp@example.com"}},
        "body": {"contentType": "text", "content": "Hello world"},
        "conversationId": "abc",
        "receivedDateTime": "2026-03-23T10:00:00Z",
        "isRead": True,
        "hasAttachments": False,
    }
    shaped = shape_email_detail(raw)
    assert "body" in shaped
    assert shaped["body"]["content"] == "Hello world"


def test_shape_event_summary_keeps_actionable_fields():
    raw = {
        "id": "evt-1",
        "subject": "AI Pilot Intro",
        "start": {"dateTime": "2026-03-24T18:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-03-24T19:00:00", "timeZone": "UTC"},
        "location": {"displayName": "Microsoft Teams Meeting"},
        "organizer": {"emailAddress": {"name": "Org", "address": "org@x.com"}},
    }
    shaped = shape_event_summary(raw)
    assert shaped["id"] == "evt-1"
    assert shaped["subject"] == "AI Pilot Intro"
    assert shaped["start"] == raw["start"]
    assert shaped["location"] == "Microsoft Teams Meeting"
    assert shaped["organizer"] == "Org <org@x.com>"


def test_shape_event_detail_extracts_teams_meeting():
    raw = {
        "id": "evt-2",
        "subject": "Sync",
        "start": {"dateTime": "2026-03-24T18:00:00", "timeZone": "UTC"},
        "end": {"dateTime": "2026-03-24T19:00:00", "timeZone": "UTC"},
        "location": {"displayName": "Teams"},
        "organizer": {"emailAddress": {"name": "Org", "address": "org@x.com"}},
        "body": {
            "contentType": "html",
            "content": '<a href="https://teams.microsoft.com/l/meetup-join/abc123">Join</a>',
        },
        "onlineMeeting": {
            "joinUrl": "https://teams.microsoft.com/l/meetup-join/abc123"
        },
        "attendees": [
            {
                "emailAddress": {"name": "A", "address": "a@x.com"},
                "status": {"response": "accepted"},
            }
        ],
    }
    shaped = shape_event_detail(raw)
    assert "meeting" in shaped
    assert (
        shaped["meeting"]["join_url"]
        == "https://teams.microsoft.com/l/meetup-join/abc123"
    )
    assert shaped["attendees"] == [{"name": "A <a@x.com>", "status": "accepted"}]


def test_shape_contact_summary_filters_empty_email_entries():
    raw = {
        "id": "c-1",
        "displayName": "Brian Roach",
        "emailAddresses": [{"address": "roach7@gmail.com"}, {}, {}],
    }
    shaped = shape_contact_summary(raw)
    assert shaped["email_addresses"] == ["roach7@gmail.com"]


def test_shape_contact_summary_separates_unresolved_addresses():
    raw = {
        "id": "c-2",
        "displayName": "Someone",
        "emailAddresses": [
            {"address": "good@smtp.com"},
            {
                "address": "/o=ExchangeLabs/ou=Exchange Administrative Group/cn=Recipients/cn=abc"
            },
        ],
    }
    shaped = shape_contact_summary(raw)
    assert shaped["email_addresses"] == ["good@smtp.com"]
    assert len(shaped["unresolved_addresses"]) == 1


def test_shape_contact_detail_includes_full_info():
    raw = {
        "id": "c-3",
        "displayName": "Full Contact",
        "emailAddresses": [{"address": "full@x.com"}],
        "businessPhones": ["+1234567890"],
        "mobilePhone": "+0987654321",
        "jobTitle": "Engineer",
        "companyName": "Acme",
        "businessAddress": {"street": "123 Main", "city": "NYC"},
    }
    shaped = shape_contact_detail(raw)
    assert shaped["jobTitle"] == "Engineer"
    assert shaped["businessAddress"]["city"] == "NYC"


def test_shape_message_summary_flattens_sender():
    raw = {
        "id": "msg-1",
        "body": {"contentType": "text", "content": "Hey team, updates?"},
        "from": {
            "user": {"displayName": "Alice", "id": "user-1"},
        },
        "createdDateTime": "2026-03-23T10:00:00Z",
    }
    shaped = shape_message_summary(raw)
    assert shaped["id"] == "msg-1"
    assert shaped["from"] == "Alice"
    assert "snippet" in shaped
    assert len(shaped["snippet"]) <= 200


import inspect
from microsoft_mcp import tools as tools_mod

LIST_OR_SEARCH_TOOLS_THAT_MUST_ACCEPT_PROFILE = [
    "list_emails",
    "list_events",
    "list_contacts",
    "list_chat_messages",
    "list_mail_folders",
    "list_master_categories",
    "list_invite_messages",
    "list_files",
    "unified_search",
    "search_files",
    "search_emails",
    "search_events",
    "search_contacts",
    "list_channel_messages",
    "search_chat_messages",
    "search_channel_messages",
    "list_inbox_items",
]


def test_all_list_and_search_tools_accept_response_profile():
    missing = []
    for name in LIST_OR_SEARCH_TOOLS_THAT_MUST_ACCEPT_PROFILE:
        tool = getattr(tools_mod, name, None)
        assert tool is not None, f"{name} not exported"
        fn = getattr(tool, "fn", tool)
        sig = inspect.signature(fn)
        if "response_profile" not in sig.parameters:
            missing.append(name)
    assert not missing, f"tools missing response_profile param: {missing}"


import microsoft_mcp.response_shaping as rs


def test_response_shaping_does_not_export_dead_types():
    assert not hasattr(rs, "ResponseProfile"), (
        "ResponseProfile enum was unused and should be removed"
    )
    assert not hasattr(rs, "BudgetHints"), (
        "BudgetHints dataclass was unused and should be removed"
    )
