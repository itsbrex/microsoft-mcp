from microsoft_mcp.inbox_models import InboxItem
from microsoft_mcp.inbox_ranking import rank_items


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
