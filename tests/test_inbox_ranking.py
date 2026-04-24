import datetime as dt

from microsoft_mcp import tools as tools_mod
from microsoft_mcp.inbox_models import InboxItem
from microsoft_mcp.inbox_ranking import _compute_score, rank_items


def _future_iso(minutes: int) -> str:
    return (
        dt.datetime.now(dt.timezone.utc) + dt.timedelta(minutes=minutes)
    ).isoformat()


def test_inbox_item_creation():
    item = InboxItem(
        id="m1",
        kind="email",
        source_tool="list_emails",
        title="Test email",
    )
    assert item.id == "m1"
    assert item.kind == "email"
    assert item.score == 0.0


def test_rank_items_prioritizes_unread_over_read():
    items = [
        InboxItem(id="m1", kind="email", title="Read email", unread=False),
        InboxItem(id="m2", kind="email", title="Unread email", unread=True),
    ]
    ranked = rank_items(items)
    assert ranked[0].id == "m2"


def test_rank_items_prioritizes_soon_events():
    items = [
        InboxItem(id="m1", kind="email", title="FYI", unread=False),
        InboxItem(
            id="e1",
            kind="event",
            title="Starts soon",
            starts_in_minutes=10,
        ),
    ]
    ranked = rank_items(items)
    assert ranked[0].title == "Starts soon"


def test_rank_items_prioritizes_mentioned():
    items = [
        InboxItem(id="m1", kind="email", title="No mention", unread=True),
        InboxItem(
            id="m2", kind="email", title="Mentioned", unread=True, mentioned=True
        ),
    ]
    ranked = rank_items(items)
    assert ranked[0].id == "m2"


def test_rank_items_prioritizes_flagged():
    items = [
        InboxItem(id="m1", kind="email", title="Normal", unread=True),
        InboxItem(id="m2", kind="email", title="Flagged", unread=True, flagged=True),
    ]
    ranked = rank_items(items)
    assert ranked[0].id == "m2"


def test_rank_items_suppresses_newsletters():
    items = [
        InboxItem(id="m1", kind="email", title="Normal email", unread=True),
        InboxItem(
            id="m2", kind="email", title="Newsletter", unread=True, is_newsletter=True
        ),
    ]
    ranked = rank_items(items)
    # Newsletter should be ranked lower
    assert ranked[0].id == "m1"


def test_rank_items_returns_scores():
    items = [
        InboxItem(id="m1", kind="email", title="Test", unread=True),
    ]
    ranked = rank_items(items)
    assert ranked[0].score > 0


def test_inbox_item_to_dict():
    item = InboxItem(
        id="m1",
        kind="email",
        source_tool="list_emails",
        title="Test",
        snippet="preview text",
        web_url="https://example.com",
    )
    d = item.to_dict()
    assert d["id"] == "m1"
    assert d["kind"] == "email"
    assert d["title"] == "Test"
    assert "snippet" in d


def test_invite_message_populates_starts_in_minutes_under_15():
    raw = [
        {
            "id": "msg-1",
            "subject": "Imminent standup",
            "meetingMessageType": "meetingRequest",
            "startDateTime": {"dateTime": _future_iso(5)},
            "isRead": False,
        }
    ]
    items = tools_mod._invite_messages_to_inbox_items(raw)
    assert items[0].starts_in_minutes is not None
    assert items[0].starts_in_minutes <= 15
    # Ranker awards +25 (<=15 min meeting) on top of unread (+10) = >=35
    assert _compute_score(items[0]) >= 35


def test_event_populates_starts_in_minutes_1_to_2_hours():
    raw = [
        {
            "id": "evt-1",
            "subject": "Later meeting",
            "start": {"dateTime": _future_iso(90)},
        }
    ]
    items = tools_mod._events_to_inbox_items(raw)
    assert items[0].starts_in_minutes is not None
    assert 60 < items[0].starts_in_minutes <= 120
    assert _compute_score(items[0]) == 5.0


def test_past_events_have_none_starts_in_minutes():
    raw = [
        {
            "id": "evt-past",
            "subject": "Already happened",
            "start": {"dateTime": _future_iso(-30)},
        }
    ]
    items = tools_mod._events_to_inbox_items(raw)
    assert items[0].starts_in_minutes is None
    assert _compute_score(items[0]) == 0.0


def test_email_flagged_status_feeds_ranker():
    raw = [
        {
            "id": "m-1",
            "subject": "Action needed",
            "isRead": True,
            "flag": {"flagStatus": "flagged"},
        }
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].flagged is True
    assert _compute_score(items[0]) == 8.0


def test_email_not_flagged_when_status_missing_or_none():
    raw = [
        {
            "id": "m-2",
            "subject": "None",
            "isRead": True,
            "flag": {"flagStatus": "notFlagged"},
        },
        {"id": "m-3", "subject": "Missing", "isRead": True},
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert all(not i.flagged for i in items)


def test_newsletter_sender_heuristic_flags_item():
    raw = [
        {
            "id": "m-news",
            "subject": "Weekly digest",
            "isRead": False,
            "from": {
                "emailAddress": {"address": "noreply@substack.com", "name": "Substack"}
            },
        }
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].is_newsletter is True
    # unread(+10) + newsletter(-20) = -10
    assert _compute_score(items[0]) == -10.0


def test_human_sender_not_newsletter():
    raw = [
        {
            "id": "m-human",
            "subject": "Hey",
            "isRead": False,
            "from": {"emailAddress": {"address": "alice@company.com", "name": "Alice"}},
        }
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].is_newsletter is False


def test_mentioned_signal_fires_when_mentionspreview_present():
    raw = [
        {
            "id": "m-ment",
            "subject": "FYI",
            "isRead": True,
            "mentionsPreview": {"isMentioned": True},
        }
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].mentioned is True
    # mentioned (+15) only, not unread
    assert _compute_score(items[0]) == 15.0


def test_not_mentioned_when_field_absent_or_false():
    raw = [
        {"id": "m-nm1", "subject": "a", "isRead": True},
        {
            "id": "m-nm2",
            "subject": "b",
            "isRead": True,
            "mentionsPreview": {"isMentioned": False},
        },
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert all(not i.mentioned for i in items)
