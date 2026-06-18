"""Tests for the intel signal collectors.

Each collector accepts a dependency-injected ``request`` callable instead of
importing ``graph`` directly.  Tests pass a canned fake to keep them pure and
deterministic.  The fixed ``now`` value (2026-06-17 09:00 UTC) is always
injected so "today's events" calculations are reproducible.
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any
from unittest.mock import MagicMock

from src.microsoft_mcp.intel.collectors.calendar import collect_calendar_signals
from src.microsoft_mcp.intel.collectors.contacts import collect_contact_signals
from src.microsoft_mcp.intel.collectors.email import collect_email_signals
from src.microsoft_mcp.intel.collectors.threads import collect_thread_signals

# Fixed "now" used throughout — 2026-06-17 09:00 UTC.
NOW = datetime(2026, 6, 17, 9, 0, 0, tzinfo=UTC)
YESTERDAY_ISO = "2026-06-16T10:00:00Z"
TWO_DAYS_AGO_ISO = "2026-06-15T08:00:00Z"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _email_request(
    folders: list | None = None,
    inbox_msgs: list | None = None,
    sent_count: int = 0,
) -> MagicMock:
    """Build a fake request callable for email collector tests.

    ``sent_count`` is satisfied by returning that many stub message dicts in
    ``value`` (no @odata.count) so the paginating _fetch_sent_count counts them
    correctly.
    """
    folder_data = folders or []
    inbox_data = inbox_msgs or []
    sent_data = [{"id": f"s{i}"} for i in range(sent_count)]

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        if "/mailFolders" in path and "messages" not in path:
            return {"value": folder_data}
        if "sentItems" in path or "sentitems" in path:
            return {"value": sent_data}
        return {"value": inbox_data}

    return MagicMock(side_effect=fake)


def _inbox_msg(mid: str, sender: str, name: str, *, is_read: bool = False) -> dict:
    return {
        "id": mid,
        "subject": f"Subj {mid}",
        "from": {"emailAddress": {"name": name, "address": sender}},
        "receivedDateTime": YESTERDAY_ISO,
        "isRead": is_read,
        "importance": "normal",
        "hasAttachments": False,
        "bodyPreview": "",
    }


def _cal_event(
    eid: str, start: str, end: str, *, is_all_day: bool = False, attendees: int = 1
) -> dict:
    return {
        "id": eid,
        "subject": f"Evt {eid}",
        "start": {"dateTime": start},
        "end": {"dateTime": end},
        "location": {"displayName": ""},
        "isAllDay": is_all_day,
        "isOnlineMeeting": False,
        "organizer": {"emailAddress": {"name": "Me", "address": "me@example.com"}},
        "attendees": [
            {"emailAddress": {"name": f"P{i}", "address": f"p{i}@example.com"}}
            for i in range(attendees)
        ],
        "responseStatus": {"response": "accepted"},
        "showAs": "busy",
    }


def _cal_request(today: list | None = None, tomorrow: list | None = None) -> MagicMock:
    today_data = today or []
    tomorrow_data = tomorrow or []
    call_count = [0]

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        call_count[0] += 1
        return {"value": today_data if call_count[0] == 1 else tomorrow_data}

    return MagicMock(side_effect=fake)


def _thread_request(
    received: list | None = None, sent: list | None = None
) -> MagicMock:
    recv_data = received or []
    sent_data = sent or []

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        if "sentitems" in path.lower() or "sentItems" in path:
            return {"value": sent_data}
        return {"value": recv_data}

    return MagicMock(side_effect=fake)


def _recv_msg(conv_id: str, mid: str, received_at: str) -> dict:
    return {
        "id": mid,
        "conversationId": conv_id,
        "subject": "Subj",
        "from": {"emailAddress": {"name": "Sender", "address": "sender@example.com"}},
        "toRecipients": [],
        "receivedDateTime": received_at,
        "importance": "normal",
    }


def _sent_msg(conv_id: str, mid: str, sent_at: str) -> dict:
    return {
        "id": mid,
        "conversationId": conv_id,
        "subject": "Subj",
        "from": {"emailAddress": {"name": "Me", "address": "me@example.com"}},
        "toRecipients": [{"emailAddress": {"name": "R", "address": "r@example.com"}}],
        "sentDateTime": sent_at,
        "importance": "normal",
    }


def _contact_request(
    received: list | None = None, sent: list | None = None
) -> MagicMock:
    recv_data = received or []
    sent_data = sent or []

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        if "sentitems" in path.lower() or "sentItems" in path:
            return {"value": sent_data}
        return {"value": recv_data}

    return MagicMock(side_effect=fake)


def _recv_contact_msg(mid: str, email: str, name: str, *, is_read: bool = True) -> dict:
    return {
        "id": mid,
        "subject": "Hi",
        "from": {"emailAddress": {"name": name, "address": email}},
        "receivedDateTime": YESTERDAY_ISO,
        "isRead": is_read,
        "importance": "normal",
        "hasAttachments": False,
        "bodyPreview": "",
    }


def _sent_contact_msg(mid: str, email: str, name: str) -> dict:
    return {
        "id": mid,
        "subject": "Hi",
        "toRecipients": [{"emailAddress": {"name": name, "address": email}}],
        "sentDateTime": YESTERDAY_ISO,
        "importance": "normal",
    }


# ---------------------------------------------------------------------------
# Email collector tests
# ---------------------------------------------------------------------------


def test_email_unread_total_sums_folders() -> None:
    folders = [
        {"displayName": "Inbox", "unreadItemCount": 5, "totalItemCount": 100},
        {"displayName": "Junk", "unreadItemCount": 3, "totalItemCount": 20},
    ]
    signals = collect_email_signals(_email_request(folders=folders), now=NOW)
    assert signals["unread_total"] == 8


def test_email_unread_by_folder_names() -> None:
    folders = [{"displayName": "Inbox", "unreadItemCount": 2, "totalItemCount": 10}]
    signals = collect_email_signals(_email_request(folders=folders), now=NOW)
    assert signals["unread_by_folder"][0]["folder_name"] == "Inbox"


def test_email_needs_response_contains_sender_info() -> None:
    msgs = [_inbox_msg("m1", "alice@example.com", "Alice")]
    signals = collect_email_signals(_email_request(inbox_msgs=msgs), now=NOW)
    assert len(signals["needs_response"]) >= 1
    assert signals["needs_response"][0]["sender_email"] == "alice@example.com"
    assert signals["needs_response"][0]["age_hours"] > 0


def test_email_vip_requires_threshold_of_3() -> None:
    # Two from alice (not VIP), three from bob (VIP).
    alice = _inbox_msg("a1", "alice@example.com", "Alice")
    bobs = [_inbox_msg(f"b{i}", "bob@example.com", "Bob") for i in range(3)]
    msgs = [alice, alice.copy()] + bobs
    signals = collect_email_signals(_email_request(inbox_msgs=msgs), now=NOW)
    vip_senders = {v["sender_email"] for v in signals["vip_unread"]}
    assert "bob@example.com" in vip_senders
    assert "alice@example.com" not in vip_senders


def test_email_sent_last_24h_from_count() -> None:
    signals = collect_email_signals(_email_request(sent_count=7), now=NOW)
    assert signals["sent_last_24h"] == 7


def test_email_received_last_24h_counts_all_inbox() -> None:
    msgs = [
        _inbox_msg(f"m{i}", f"u{i}@ex.com", f"U{i}", is_read=True) for i in range(4)
    ]
    signals = collect_email_signals(_email_request(inbox_msgs=msgs), now=NOW)
    assert signals["received_last_24h"] == 4


def test_email_no_messages_returns_zero_totals() -> None:
    signals = collect_email_signals(_email_request(), now=NOW)
    assert signals["unread_total"] == 0
    assert signals["needs_response"] == []
    assert signals["vip_unread"] == []


def test_fetch_sent_count_paginates_across_nextlink() -> None:
    """_fetch_sent_count must follow @odata.nextLink and sum across pages.

    This is the I2 regression test: the old $count approach silently returned
    len(value) capped at _SENT_PAGE_SIZE.  The new paginating approach must
    count items from ALL pages.
    """
    page1 = [{"id": f"s1_{i}"} for i in range(3)]
    page2 = [{"id": f"s2_{i}"} for i in range(2)]
    calls: list[str] = []

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        if "/mailFolders" in path and "messages" not in path:
            return {"value": []}
        if "sentItems" in path or "sentitems" in path:
            calls.append(path)
            if len(calls) == 1:
                # First page — include nextLink
                return {
                    "value": page1,
                    "@odata.nextLink": (
                        "https://graph.microsoft.com/v1.0"
                        "/me/mailFolders/sentItems/messages?$skip=3"
                    ),
                }
            # Second page — no nextLink
            return {"value": page2}
        return {"value": []}

    req = MagicMock(side_effect=fake)
    signals = collect_email_signals(req, now=NOW)

    # Must be sum of both pages (3 + 2 = 5), not capped at first page (3)
    assert signals["sent_last_24h"] == 5
    # Pagination was actually triggered — at least two sent-path calls
    assert len(calls) >= 2


# ---------------------------------------------------------------------------
# Calendar collector tests
# ---------------------------------------------------------------------------

_T_START = "2026-06-17T09:00:00"
_T_END = "2026-06-17T10:00:00"
_T_OVERLAP_START = "2026-06-17T09:30:00"
_T_OVERLAP_END = "2026-06-17T10:30:00"


def test_calendar_today_events_returned() -> None:
    ev = _cal_event("e1", _T_START, _T_END)
    signals = collect_calendar_signals(_cal_request(today=[ev]), now=NOW)
    assert signals["total_events_today"] == 1
    assert signals["today_events"][0]["subject"] == "Evt e1"


def test_calendar_conflict_on_overlapping_events() -> None:
    ev1 = _cal_event("e1", _T_START, _T_END)
    ev2 = _cal_event("e2", _T_OVERLAP_START, _T_OVERLAP_END)
    signals = collect_calendar_signals(_cal_request(today=[ev1, ev2]), now=NOW)
    assert len(signals["conflicts"]) == 1
    assert signals["conflicts"][0]["overlap_minutes"] == 30


def test_calendar_no_conflict_on_sequential_events() -> None:
    ev1 = _cal_event("e1", "2026-06-17T09:00:00", "2026-06-17T10:00:00")
    ev2 = _cal_event("e2", "2026-06-17T10:00:00", "2026-06-17T11:00:00")
    signals = collect_calendar_signals(_cal_request(today=[ev1, ev2]), now=NOW)
    assert signals["conflicts"] == []


def test_calendar_meeting_hours_today() -> None:
    ev = _cal_event("e1", _T_START, _T_END)
    signals = collect_calendar_signals(_cal_request(today=[ev]), now=NOW)
    assert signals["meeting_hours_today"] == 1.0


def test_calendar_prep_needed_three_attendees() -> None:
    ev = _cal_event("e1", _T_START, _T_END, attendees=3)
    signals = collect_calendar_signals(_cal_request(today=[ev]), now=NOW)
    assert len(signals["prep_needed"]) == 1


def test_calendar_prep_not_needed_for_solo() -> None:
    ev = _cal_event("e1", _T_START, _T_END, attendees=1)
    signals = collect_calendar_signals(_cal_request(today=[ev]), now=NOW)
    assert signals["prep_needed"] == []


def test_calendar_tomorrow_events_returned() -> None:
    tom = _cal_event("t1", "2026-06-18T09:00:00", "2026-06-18T10:00:00")
    signals = collect_calendar_signals(_cal_request(tomorrow=[tom]), now=NOW)
    assert len(signals["tomorrow_events"]) == 1


def test_calendar_all_day_excluded_from_meeting_hours() -> None:
    ev = _cal_event("e1", "2026-06-17T00:00:00", "2026-06-18T00:00:00", is_all_day=True)
    signals = collect_calendar_signals(_cal_request(today=[ev]), now=NOW)
    assert signals["meeting_hours_today"] == 0.0


def test_calendar_no_events_zero_hours() -> None:
    signals = collect_calendar_signals(_cal_request(), now=NOW)
    assert signals["meeting_hours_today"] == 0.0
    assert signals["conflicts"] == []


def test_calendar_timezone_converts_utc_to_local() -> None:
    """UTC datetimes from Graph must be converted to naive local time.

    NOW is 2026-06-17T09:00:00 UTC.  America/Chicago is CDT = UTC-5 in June,
    so the local date is also 2026-06-17 (09:00 UTC = 04:00 CDT).

    The Graph event starts at 14:00 UTC / 09:00 CDT and ends at 15:00 UTC /
    10:00 CDT.  The collector must store the event times as naive local ISO
    strings ("2026-06-17T09:00:00" / "2026-06-17T10:00:00") not as UTC
    ("2026-06-17T14:00:00" / "2026-06-17T15:00:00").

    Free blocks for the day run 08:00–18:00 local.  With one 09:00-10:00 CDT
    event there should be a free block starting at 08:00 and another starting
    at 10:00 local time — NOT at 14:00/15:00 which would be wrong (UTC).

    These assertions would FAIL against the old code that passed times through
    as UTC strings.
    """
    # Graph returns UTC regardless of timezone requested — simulate that.
    ev_utc = _cal_event(
        "tz1", "2026-06-17T14:00:00.0000000", "2026-06-17T15:00:00.0000000"
    )
    signals = collect_calendar_signals(
        _cal_request(today=[ev_utc]),
        now=NOW,
        timezone="America/Chicago",
    )

    # 1. Event time stored as naive CDT (local), not UTC.
    assert signals["today_events"][0]["start"] == "2026-06-17T09:00:00"
    assert signals["today_events"][0]["end"] == "2026-06-17T10:00:00"

    # 2. Meeting duration is tz-invariant: still 1 hour.
    assert signals["meeting_hours_today"] == 1.0

    # 3. Free blocks are anchored to local working hours 08:00–18:00 CDT.
    #    With one event at 09:00–10:00 CDT there must be a pre-event gap
    #    (08:00–09:00) — this would be absent if times were left in UTC.
    free_starts = [b["start"] for b in signals["free_blocks"]]
    assert "2026-06-17T08:00:00" in free_starts, (
        f"Expected 08:00 CDT free block; got {free_starts}"
    )
    # There must be no free block starting at the UTC time 14:00.
    assert "2026-06-17T14:00:00" not in free_starts, (
        "Free block at 14:00 means timezone conversion was not applied"
    )


# ---------------------------------------------------------------------------
# Thread collector tests
# ---------------------------------------------------------------------------


def test_threads_awaiting_my_reply_when_inbound_last() -> None:
    recv = _recv_msg("conv1", "r1", YESTERDAY_ISO)
    signals = collect_thread_signals(_thread_request(received=[recv]), now=NOW)
    assert len(signals["awaiting_my_reply"]) == 1
    assert signals["awaiting_my_reply"][0]["conversation_id"] == "conv1"


def test_threads_awaiting_their_reply_when_outbound_last() -> None:
    sent = _sent_msg("conv2", "s1", YESTERDAY_ISO)
    signals = collect_thread_signals(_thread_request(sent=[sent]), now=NOW)
    assert len(signals["awaiting_their_reply"]) == 1
    assert signals["awaiting_their_reply"][0]["conversation_id"] == "conv2"


def test_threads_direction_by_most_recent() -> None:
    # recv older, sent newer → outbound
    recv = _recv_msg("conv3", "r3", TWO_DAYS_AGO_ISO)
    sent = _sent_msg("conv3", "s3", YESTERDAY_ISO)
    signals = collect_thread_signals(
        _thread_request(received=[recv], sent=[sent]), now=NOW
    )
    assert len(signals["awaiting_their_reply"]) == 1
    assert signals["awaiting_my_reply"] == []


def test_threads_stale_when_exceeds_stale_hours() -> None:
    sent = _sent_msg("conv4", "s4", TWO_DAYS_AGO_ISO)
    signals = collect_thread_signals(
        _thread_request(sent=[sent]), now=NOW, stale_hours=1
    )
    assert len(signals["stale_threads"]) == 1


def test_threads_not_stale_when_recent() -> None:
    sent = _sent_msg("conv5", "s5", YESTERDAY_ISO)
    signals = collect_thread_signals(
        _thread_request(sent=[sent]), now=NOW, stale_hours=72
    )
    assert signals["stale_threads"] == []


def test_threads_message_count_combines_directions() -> None:
    recv = _recv_msg("conv6", "r6", TWO_DAYS_AGO_ISO)
    sent = _sent_msg("conv6", "s6", YESTERDAY_ISO)
    signals = collect_thread_signals(
        _thread_request(received=[recv], sent=[sent]), now=NOW
    )
    assert signals["awaiting_their_reply"][0]["message_count"] == 2


def test_threads_sorted_oldest_first() -> None:
    old = _recv_msg("c7", "r7", TWO_DAYS_AGO_ISO)
    new = _recv_msg("c8", "r8", YESTERDAY_ISO)
    signals = collect_thread_signals(_thread_request(received=[new, old]), now=NOW)
    ages = [t["age_hours"] for t in signals["awaiting_my_reply"]]
    assert ages == sorted(ages, reverse=True)


def test_threads_empty_returns_empty_signals() -> None:
    signals = collect_thread_signals(_thread_request(), now=NOW)
    assert signals["awaiting_my_reply"] == []
    assert signals["awaiting_their_reply"] == []
    assert signals["stale_threads"] == []


# ---------------------------------------------------------------------------
# Contact collector tests
# ---------------------------------------------------------------------------


def test_contacts_top_sorted_by_total_interactions() -> None:
    alice = [_recv_contact_msg(f"ra{i}", "alice@ex.com", "Alice") for i in range(5)]
    bob = [_recv_contact_msg(f"rb{i}", "bob@ex.com", "Bob") for i in range(2)]
    signals = collect_contact_signals(_contact_request(received=alice + bob), now=NOW)
    assert signals["top_contacts"][0]["email"] == "alice@ex.com"
    assert signals["top_contacts"][0]["total_interactions"] == 5


def test_contacts_pending_has_unread() -> None:
    unread = _recv_contact_msg("u1", "unread@ex.com", "Unread", is_read=False)
    read = _recv_contact_msg("r1", "read@ex.com", "Read", is_read=True)
    signals = collect_contact_signals(
        _contact_request(received=[unread, read]), now=NOW
    )
    pending_emails = {c["email"] for c in signals["pending_contacts"]}
    assert "unread@ex.com" in pending_emails
    assert "read@ex.com" not in pending_emails


def test_contacts_unique_senders_count() -> None:
    msgs = [_recv_contact_msg(f"r{i}", f"u{i}@ex.com", f"U{i}") for i in range(3)]
    signals = collect_contact_signals(_contact_request(received=msgs), now=NOW)
    assert signals["total_unique_senders"] == 3


def test_contacts_unique_recipients_count() -> None:
    sent = [_sent_contact_msg(f"s{i}", f"u{i}@ex.com", f"U{i}") for i in range(4)]
    signals = collect_contact_signals(_contact_request(sent=sent), now=NOW)
    assert signals["total_unique_recipients"] == 4


def test_contacts_interaction_combines_sent_and_received() -> None:
    recv = _recv_contact_msg("r1", "shared@ex.com", "Shared")
    sent = _sent_contact_msg("s1", "shared@ex.com", "Shared")
    signals = collect_contact_signals(
        _contact_request(received=[recv], sent=[sent]), now=NOW
    )
    shared = next(c for c in signals["top_contacts"] if c["email"] == "shared@ex.com")
    assert shared["total_interactions"] == 2
    assert shared["received_from"] == 1
    assert shared["sent_to"] == 1


def test_contacts_no_messages_returns_empty() -> None:
    signals = collect_contact_signals(_contact_request(), now=NOW)
    assert signals["top_contacts"] == []
    assert signals["pending_contacts"] == []
    assert signals["total_unique_senders"] == 0
    assert signals["total_unique_recipients"] == 0


# ---------------------------------------------------------------------------
# Pagination tests (I2 + M3)
# ---------------------------------------------------------------------------


def _paged_email_request(page1: list, page2: list) -> MagicMock:
    """Fake request that returns two pages for each inbox call, empty for folders/sent.

    collect_email_signals calls inbox twice (unread_only=True, then unread_only=False).
    Each call to paginate() will: first hit the normal path (returns page1 + nextLink),
    then hit the stripped nextLink path (returns page2).  We track per-"base path"
    whether we've already served page1 so both paginations get all items.
    """
    # Track how many times each logical path has been fetched at the first-page level
    first_page_served: set[str] = set()

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        if "/mailFolders" in path and "messages" not in path:
            return {"value": []}
        if "sentItems" in path or "sentitems" in path:
            return {"value": [], "@odata.count": 0}
        # If this is the stripped nextLink path, return page2 (second page)
        if "$skip=50" in path:
            return {"value": page2}
        # First page for this inbox path
        first_page_served.add(path)
        return {
            "value": page1,
            "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=50",
        }

    return MagicMock(side_effect=fake)


def test_email_inbox_pagination_merges_pages() -> None:
    """paginate() follows @odata.nextLink and merges both pages."""
    page1 = [_inbox_msg(f"p1m{i}", f"u{i}@ex.com", f"U{i}") for i in range(3)]
    page2 = [_inbox_msg(f"p2m{i}", f"v{i}@ex.com", f"V{i}") for i in range(2)]
    req = _paged_email_request(page1, page2)
    signals = collect_email_signals(req, now=NOW)
    # received_last_24h counts all inbox (unread_only=False path); should see all 5
    assert signals["received_last_24h"] == 5


def test_email_inbox_pagination_second_call_uses_stripped_path() -> None:
    """The second call's path must be the base-stripped nextLink."""
    page1 = [_inbox_msg("m1", "a@ex.com", "A")]
    page2 = [_inbox_msg("m2", "b@ex.com", "B")]

    second_call_path: list[str] = []
    calls: list[str] = []

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        if "/mailFolders" in path and "messages" not in path:
            return {"value": []}
        if "sentItems" in path or "sentitems" in path:
            return {"value": [], "@odata.count": 0}
        calls.append(path)
        if len(calls) == 1:
            return {
                "value": page1,
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=50",
            }
        second_call_path.append(path)
        return {"value": page2}

    req = MagicMock(side_effect=fake)
    collect_email_signals(req, now=NOW)

    # At least one second-page call must have been made
    assert second_call_path, "No second page call was made"
    # The path must be the stripped form, not the full URL
    assert second_call_path[0] == "/me/messages?$skip=50"


def _paged_contact_request(page1: list, page2: list) -> MagicMock:
    """Fake request that returns two pages for the inbox path, empty for sent."""
    calls: list[str] = []

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        if "sentitems" in path.lower() or "sentItems" in path:
            return {"value": []}
        calls.append(path)
        if len(calls) == 1:
            return {
                "value": page1,
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=50",
            }
        return {"value": page2}

    return MagicMock(side_effect=fake)


def test_contacts_pagination_merges_pages() -> None:
    """paginate() in contacts._fetch_messages merges items across pages."""
    page1 = [_recv_contact_msg(f"p1r{i}", f"a{i}@ex.com", f"A{i}") for i in range(3)]
    page2 = [_recv_contact_msg(f"p2r{i}", f"b{i}@ex.com", f"B{i}") for i in range(2)]
    req = _paged_contact_request(page1, page2)
    signals = collect_contact_signals(req, now=NOW)
    # All 5 senders (3 from page1, 2 from page2) must be counted
    assert signals["total_unique_senders"] == 5


def test_contacts_pagination_second_call_uses_stripped_path() -> None:
    """The second contacts fetch call's path is the base-stripped nextLink."""
    page1 = [_recv_contact_msg("r1", "a@ex.com", "A")]
    page2 = [_recv_contact_msg("r2", "b@ex.com", "B")]

    second_call_path: list[str] = []
    calls: list[str] = []

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        if "sentitems" in path.lower() or "sentItems" in path:
            return {"value": []}
        calls.append(path)
        if len(calls) == 1:
            return {
                "value": page1,
                "@odata.nextLink": "https://graph.microsoft.com/v1.0/me/messages?$skip=50",
            }
        second_call_path.append(path)
        return {"value": page2}

    req = MagicMock(side_effect=fake)
    collect_contact_signals(req, now=NOW)

    assert second_call_path, "No second page call was made"
    assert second_call_path[0] == "/me/messages?$skip=50"


# ---------------------------------------------------------------------------
# I3: contacts _fetch_messages re-raise test
# ---------------------------------------------------------------------------


def test_contacts_fetch_messages_propagates_exception() -> None:
    """_fetch_messages (and collect_contact_signals) must propagate exceptions."""
    import pytest

    def failing_request(
        method: str, path: str, *, params: dict | None = None, **_: Any
    ) -> dict:
        raise RuntimeError("simulated network failure")

    with pytest.raises(RuntimeError, match="simulated network failure"):
        collect_contact_signals(MagicMock(side_effect=failing_request), now=NOW)
