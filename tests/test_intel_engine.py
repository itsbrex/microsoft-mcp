"""Tests for the intel analyzers and report engine.

Analyzers and engine functions accept dependency-injected ``request``
callables and a fixed ``now`` datetime so results are deterministic.
The fixed ``now`` value is 2026-06-17 09:00 UTC (same as collector tests).
"""

from __future__ import annotations

from datetime import UTC, datetime
from typing import Any
from unittest.mock import MagicMock

from src.microsoft_mcp.intel.analyzers.priority import score_priorities
from src.microsoft_mcp.intel.analyzers.relationships import analyze_relationships
from src.microsoft_mcp.intel.analyzers.schedule import analyze_schedule
from src.microsoft_mcp.intel.engine import (
    generate_briefing,
    generate_contact_report,
    generate_recap,
    generate_signals,
)
from src.microsoft_mcp.intel.types import (
    CalendarEvent,
    CalendarSignals,
    ConflictPair,
    ContactInteraction,
    ContactSignals,
    EmailNeedingResponse,
    EmailSignals,
    ThreadInfo,
    ThreadSignals,
    TimeBlock,
    VipEmail,
)

NOW = datetime(2026, 6, 17, 9, 0, 0, tzinfo=UTC)

# ISO datetimes relative to NOW
_YESTERDAY = "2026-06-16T09:00:00Z"
_4_DAYS_AGO = "2026-06-13T09:00:00Z"
_6_DAYS_AGO = "2026-06-11T09:00:00Z"

# ---------------------------------------------------------------------------
# Factories for canned signal objects
# ---------------------------------------------------------------------------


def _make_email_signals(
    needs_response: list[EmailNeedingResponse] | None = None,
    vip_unread: list[VipEmail] | None = None,
) -> EmailSignals:
    return EmailSignals(
        unread_total=0,
        unread_by_folder=[],
        needs_response=needs_response or [],
        vip_unread=vip_unread or [],
        received_last_24h=5,
        sent_last_24h=3,
    )


def _make_calendar_signals(
    today_events: list[CalendarEvent] | None = None,
    conflicts: list[ConflictPair] | None = None,
    prep_needed: list[CalendarEvent] | None = None,
    free_blocks: list[TimeBlock] | None = None,
    meeting_hours: float = 0.0,
) -> CalendarSignals:
    events = today_events or []
    return CalendarSignals(
        today_events=events,
        tomorrow_events=[],
        conflicts=conflicts or [],
        prep_needed=prep_needed or [],
        free_blocks=free_blocks or [],
        meeting_hours_today=meeting_hours,
        total_events_today=len(events),
    )


def _make_thread_signals(
    awaiting_my_reply: list[ThreadInfo] | None = None,
    awaiting_their_reply: list[ThreadInfo] | None = None,
    stale_threads: list[ThreadInfo] | None = None,
) -> ThreadSignals:
    return ThreadSignals(
        awaiting_my_reply=awaiting_my_reply or [],
        awaiting_their_reply=awaiting_their_reply or [],
        stale_threads=stale_threads or [],
    )


def _make_contact_signals(
    top_contacts: list[ContactInteraction] | None = None,
) -> ContactSignals:
    return ContactSignals(
        top_contacts=top_contacts or [],
        pending_contacts=[],
        total_unique_senders=1,
        total_unique_recipients=1,
    )


def _needs_response(
    age_hours: float = 10.0,
    importance: str = "normal",
    has_attachments: bool = False,
) -> EmailNeedingResponse:
    return EmailNeedingResponse(
        id="e1",
        subject="Hello",
        sender_name="Alice",
        sender_email="alice@example.com",
        received_at=_YESTERDAY,
        age_hours=age_hours,
        importance=importance,
        has_attachments=has_attachments,
        body_preview="",
    )


def _vip_email(interaction_count: int = 3, importance: str = "normal") -> VipEmail:
    return VipEmail(
        id="v1",
        subject="VIP Message",
        sender_name="Bob",
        sender_email="bob@example.com",
        received_at=_YESTERDAY,
        is_read=False,
        importance=importance,
        interaction_count=interaction_count,
    )


def _conflict() -> ConflictPair:
    return ConflictPair(
        event_a_subject="Standup",
        event_a_start="2026-06-17T09:00:00Z",
        event_b_subject="1:1",
        event_b_start="2026-06-17T09:30:00Z",
        overlap_minutes=30,
    )


def _calendar_event(
    eid: str = "evt1",
    has_external: bool = False,
    attendee_count: int = 2,
    is_all_day: bool = False,
    start: str = "2026-06-17T10:00:00Z",
    end: str = "2026-06-17T11:00:00Z",
    response_status: str = "accepted",
) -> CalendarEvent:
    return CalendarEvent(
        id=eid,
        subject=f"Event {eid}",
        start=start,
        end=end,
        location="",
        is_all_day=is_all_day,
        is_online=False,
        organizer_name="Me",
        organizer_email="me@example.com",
        attendee_count=attendee_count,
        has_external_attendees=has_external,
        response_status=response_status,
        show_as="busy",
    )


def _thread_info(
    age_hours: float = 5.0,
    importance: str = "normal",
    conv_id: str = "conv1",
    last_sender_email: str = "alice@example.com",
) -> ThreadInfo:
    return ThreadInfo(
        conversation_id=conv_id,
        subject="Thread Subject",
        last_sender_name="Alice",
        last_sender_email=last_sender_email,
        last_message_at=_YESTERDAY,
        age_hours=age_hours,
        message_count=3,
        direction="inbound",
        importance=importance,
    )


def _contact_interaction(
    email: str = "alice@example.com",
    name: str = "Alice",
    total_interactions: int = 5,
    sent_to: int = 2,
    received_from: int = 3,
    last_interaction: str = _YESTERDAY,
) -> ContactInteraction:
    return ContactInteraction(
        email=email,
        name=name,
        total_interactions=total_interactions,
        sent_to=sent_to,
        received_from=received_from,
        last_interaction=last_interaction,
        has_pending_email=False,
    )


# ---------------------------------------------------------------------------
# Priority analyzer tests
# ---------------------------------------------------------------------------


class TestScorePriorities:
    def test_empty_signals_returns_empty_list(self) -> None:
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        assert result == []

    def test_conflict_scores_90(self) -> None:
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(conflicts=[_conflict()]),
            _make_thread_signals(),
        )
        assert len(result) == 1
        assert result[0]["score"] == 90.0
        assert result[0]["category"] == "conflict"
        assert result[0]["source"] == "calendar"

    def test_needs_response_base_60_plus_age(self) -> None:
        email = _needs_response(age_hours=5.0)
        result = score_priorities(
            _make_email_signals(needs_response=[email]),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        assert len(result) == 1
        # Base 60 + min(5, 10) age bonus = 65
        assert result[0]["score"] == 65.0
        assert result[0]["category"] == "needs_response"

    def test_needs_response_high_importance_adds_20(self) -> None:
        email = _needs_response(age_hours=0.0, importance="high")
        result = score_priorities(
            _make_email_signals(needs_response=[email]),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        # Base 60 + 20 importance = 80
        assert result[0]["score"] == 80.0

    def test_needs_response_with_attachments_adds_10(self) -> None:
        email = _needs_response(age_hours=0.0, has_attachments=True)
        result = score_priorities(
            _make_email_signals(needs_response=[email]),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        # Base 60 + 10 attachment = 70
        assert result[0]["score"] == 70.0

    def test_needs_response_age_capped_at_10(self) -> None:
        email = _needs_response(age_hours=50.0)
        result = score_priorities(
            _make_email_signals(needs_response=[email]),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        # Base 60 + min(50, 10) = 70
        assert result[0]["score"] == 70.0

    def test_vip_email_base_50(self) -> None:
        vip = _vip_email(interaction_count=1)
        result = score_priorities(
            _make_email_signals(vip_unread=[vip]),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        # Base 50 + 1*5 = 55
        assert result[0]["score"] == 55.0
        assert result[0]["category"] == "vip_email"

    def test_vip_email_interaction_count_capped_at_25(self) -> None:
        vip = _vip_email(interaction_count=10)  # 10*5=50 -> capped at 25
        result = score_priorities(
            _make_email_signals(vip_unread=[vip]),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        # Base 50 + 25 (capped) = 75
        assert result[0]["score"] == 75.0

    def test_vip_high_importance_adds_15(self) -> None:
        vip = _vip_email(interaction_count=1, importance="high")
        result = score_priorities(
            _make_email_signals(vip_unread=[vip]),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        # Base 50 + 5 + 15 = 70
        assert result[0]["score"] == 70.0

    def test_prep_needed_base_40(self) -> None:
        event = _calendar_event(has_external=False, attendee_count=2)
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(prep_needed=[event]),
            _make_thread_signals(),
        )
        assert result[0]["score"] == 40.0
        assert result[0]["category"] == "prep_needed"

    def test_prep_needed_external_adds_20(self) -> None:
        event = _calendar_event(has_external=True, attendee_count=2)
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(prep_needed=[event]),
            _make_thread_signals(),
        )
        # 40 + 20 = 60
        assert result[0]["score"] == 60.0

    def test_prep_needed_five_plus_attendees_adds_10(self) -> None:
        event = _calendar_event(has_external=True, attendee_count=5)
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(prep_needed=[event]),
            _make_thread_signals(),
        )
        # 40 + 20 + 10 = 70
        assert result[0]["score"] == 70.0

    def test_awaiting_my_reply_base_55(self) -> None:
        thread = _thread_info(age_hours=0.0)
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(),
            _make_thread_signals(awaiting_my_reply=[thread]),
        )
        assert result[0]["score"] == 55.0

    def test_awaiting_my_reply_high_importance_adds_20(self) -> None:
        thread = _thread_info(age_hours=0.0, importance="high")
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(),
            _make_thread_signals(awaiting_my_reply=[thread]),
        )
        # 55 + 20 = 75
        assert result[0]["score"] == 75.0

    def test_awaiting_my_reply_age_bonus_capped_at_15(self) -> None:
        thread = _thread_info(age_hours=60.0)  # 60/2=30, capped at 15
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(),
            _make_thread_signals(awaiting_my_reply=[thread]),
        )
        # 55 + 15 = 70
        assert result[0]["score"] == 70.0

    def test_stale_thread_not_scored_below_72h(self) -> None:
        thread = _thread_info(age_hours=71.0)
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(),
            _make_thread_signals(awaiting_their_reply=[thread]),
        )
        assert result == []

    def test_stale_thread_scored_above_72h(self) -> None:
        thread = _thread_info(age_hours=80.0)
        result = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(),
            _make_thread_signals(awaiting_their_reply=[thread]),
        )
        assert len(result) == 1
        # Base 30 + min(80/4, 20) = 30 + 20 = 50
        assert result[0]["score"] == 50.0
        assert result[0]["category"] == "stale_thread"

    def test_sorted_by_score_descending(self) -> None:
        conflict = _conflict()  # 90
        email = _needs_response(age_hours=0.0)  # 60
        thread = _thread_info(age_hours=0.0)  # 55
        result = score_priorities(
            _make_email_signals(needs_response=[email]),
            _make_calendar_signals(conflicts=[conflict]),
            _make_thread_signals(awaiting_my_reply=[thread]),
        )
        scores = [item["score"] for item in result]
        assert scores == sorted(scores, reverse=True)
        assert scores[0] == 90.0

    def test_score_capped_at_100(self) -> None:
        email = _needs_response(age_hours=10.0, importance="high", has_attachments=True)
        result = score_priorities(
            _make_email_signals(needs_response=[email]),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        # 60 + 20 + 10 + 10 = 100 (exactly at cap)
        assert result[0]["score"] == 100.0


# ---------------------------------------------------------------------------
# Schedule analyzer tests
# ---------------------------------------------------------------------------


class TestAnalyzeSchedule:
    def test_empty_day(self) -> None:
        cal = _make_calendar_signals(meeting_hours=0.0)
        result = analyze_schedule(cal)
        assert result["meeting_density_pct"] == 0.0
        assert result["back_to_back_count"] == 0
        assert result["external_meeting_count"] == 0
        assert result["longest_free_block_minutes"] == 0
        assert result["focus_time_available"] is False
        assert "Light" in result["summary"]

    def test_density_calculation(self) -> None:
        # 5 hours of meetings in a 10-hour day = 50%
        cal = _make_calendar_signals(meeting_hours=5.0)
        result = analyze_schedule(cal)
        assert result["meeting_density_pct"] == 50.0
        assert "Moderate" in result["summary"]

    def test_density_capped_at_100(self) -> None:
        cal = _make_calendar_signals(meeting_hours=15.0)
        result = analyze_schedule(cal)
        assert result["meeting_density_pct"] == 100.0

    def test_back_to_back_detection(self) -> None:
        # Two events with 5 min gap (< 15 min threshold)
        evt1 = _calendar_event(
            "e1", start="2026-06-17T10:00:00Z", end="2026-06-17T11:00:00Z"
        )
        evt2 = _calendar_event(
            "e2", start="2026-06-17T11:05:00Z", end="2026-06-17T12:00:00Z"
        )
        cal = _make_calendar_signals(today_events=[evt1, evt2], meeting_hours=2.0)
        result = analyze_schedule(cal)
        assert result["back_to_back_count"] == 1

    def test_no_back_to_back_with_large_gap(self) -> None:
        evt1 = _calendar_event(
            "e1", start="2026-06-17T09:00:00Z", end="2026-06-17T10:00:00Z"
        )
        evt2 = _calendar_event(
            "e2", start="2026-06-17T11:00:00Z", end="2026-06-17T12:00:00Z"
        )
        cal = _make_calendar_signals(today_events=[evt1, evt2], meeting_hours=2.0)
        result = analyze_schedule(cal)
        assert result["back_to_back_count"] == 0

    def test_external_meeting_count(self) -> None:
        external = _calendar_event("e1", has_external=True)
        internal = _calendar_event("e2", has_external=False)
        cal = _make_calendar_signals(
            today_events=[external, internal], meeting_hours=2.0
        )
        result = analyze_schedule(cal)
        assert result["external_meeting_count"] == 1

    def test_focus_time_available(self) -> None:
        free = TimeBlock(
            start="2026-06-17T13:00:00Z",
            end="2026-06-17T15:00:00Z",
            duration_minutes=120,
        )
        cal = _make_calendar_signals(free_blocks=[free])
        result = analyze_schedule(cal)
        assert result["focus_time_available"] is True
        assert "focus time available" in result["summary"]

    def test_no_focus_time(self) -> None:
        free = TimeBlock(
            start="2026-06-17T13:00:00Z",
            end="2026-06-17T14:00:00Z",
            duration_minutes=60,
        )
        cal = _make_calendar_signals(free_blocks=[free])
        result = analyze_schedule(cal)
        assert result["focus_time_available"] is False
        assert "no focus blocks" in result["summary"]

    def test_density_labels(self) -> None:
        assert (
            "Packed"
            in analyze_schedule(_make_calendar_signals(meeting_hours=9.0))["summary"]
        )
        assert (
            "Busy"
            in analyze_schedule(_make_calendar_signals(meeting_hours=7.0))["summary"]
        )


# ---------------------------------------------------------------------------
# Relationship analyzer tests
# ---------------------------------------------------------------------------


class TestAnalyzeRelationships:
    def test_empty_contacts(self) -> None:
        contacts = _make_contact_signals()
        result = analyze_relationships(contacts, now=NOW)
        assert result == []

    def test_engagement_score_computed(self) -> None:
        contact = _contact_interaction(
            total_interactions=5,
            sent_to=2,
            received_from=3,
            last_interaction=_YESTERDAY,  # 1 day ago from NOW
        )
        contacts = _make_contact_signals(top_contacts=[contact])
        result = analyze_relationships(contacts, now=NOW)
        assert len(result) == 1
        insight = result[0]
        # Base: min(5*5, 50) = 25, recency: +30 (1 day), bidirectional: +20 = 75
        assert insight["engagement_score"] == 75.0

    def test_trend_rising(self) -> None:
        # recent (<=2 days) and frequent (>=5 interactions)
        contact = _contact_interaction(
            total_interactions=5,
            last_interaction=_YESTERDAY,
        )
        contacts = _make_contact_signals(top_contacts=[contact])
        result = analyze_relationships(contacts, now=NOW)
        assert result[0]["trend"] == "rising"

    def test_trend_cooling(self) -> None:
        # no contact in >= 5 days (6 days ago satisfies the >= 5 threshold)
        contact = _contact_interaction(
            total_interactions=2,
            last_interaction=_6_DAYS_AGO,
        )
        contacts = _make_contact_signals(top_contacts=[contact])
        result = analyze_relationships(contacts, now=NOW)
        assert result[0]["trend"] == "cooling"

    def test_trend_stable(self) -> None:
        # recent but not enough interactions
        contact = _contact_interaction(
            total_interactions=2,
            last_interaction=_YESTERDAY,
        )
        contacts = _make_contact_signals(top_contacts=[contact])
        result = analyze_relationships(contacts, now=NOW)
        assert result[0]["trend"] == "stable"

    def test_response_ratio(self) -> None:
        contact = _contact_interaction(sent_to=2, received_from=4)
        contacts = _make_contact_signals(top_contacts=[contact])
        result = analyze_relationships(contacts, now=NOW)
        assert result[0]["response_ratio"] == 2.0

    def test_sorted_by_engagement_descending(self) -> None:
        high = _contact_interaction(
            email="high@example.com",
            total_interactions=10,
            last_interaction=_YESTERDAY,
        )
        low = _contact_interaction(
            email="low@example.com",
            total_interactions=1,
            last_interaction=_4_DAYS_AGO,
        )
        contacts = _make_contact_signals(top_contacts=[low, high])
        result = analyze_relationships(contacts, now=NOW)
        assert result[0]["email"] == "high@example.com"

    def test_company_field_propagated(self) -> None:
        contact = _contact_interaction()
        contact["company"] = "Acme Corp"
        contacts = _make_contact_signals(top_contacts=[contact])
        result = analyze_relationships(contacts, now=NOW)
        assert result[0].get("company") == "Acme Corp"


# ---------------------------------------------------------------------------
# Signal bucketing thresholds (critical/important/informational)
# ---------------------------------------------------------------------------


class TestSignalBucketing:
    """Verify the three bucketing thresholds used by generate_signals."""

    def test_critical_threshold_at_80(self) -> None:
        # conflict scores exactly 90 -> critical
        conflict = _conflict()
        emails = _make_email_signals()
        cal = _make_calendar_signals(conflicts=[conflict])
        threads = _make_thread_signals()
        items = score_priorities(emails, cal, threads)
        critical = [i for i in items if i["score"] >= 80]
        assert len(critical) == 1
        assert critical[0]["score"] == 90.0

    def test_important_range_50_to_79(self) -> None:
        # awaiting_my_reply base=55 -> important
        thread = _thread_info(age_hours=0.0)
        items = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(),
            _make_thread_signals(awaiting_my_reply=[thread]),
        )
        important = [i for i in items if 50 <= i["score"] < 80]
        assert len(important) == 1
        assert important[0]["score"] == 55.0

    def test_informational_below_50(self) -> None:
        # prep_needed base=40 (no external, <5 attendees) -> informational
        event = _calendar_event(has_external=False, attendee_count=2)
        items = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(prep_needed=[event]),
            _make_thread_signals(),
        )
        info = [i for i in items if i["score"] < 50]
        assert len(info) == 1
        assert info[0]["score"] == 40.0

    def test_score_79_is_important_not_critical(self) -> None:
        # vip with interaction_count=4 (4*5=20, capped=20), importance=high: 50+20+15=85 -> critical
        # Use no-importance vip with interaction_count=5 (5*5=25): 50+25=75 -> important
        vip = _vip_email(interaction_count=5, importance="normal")
        items = score_priorities(
            _make_email_signals(vip_unread=[vip]),
            _make_calendar_signals(),
            _make_thread_signals(),
        )
        assert items[0]["score"] == 75.0
        assert 50 <= items[0]["score"] < 80

    def test_score_49_is_informational(self) -> None:
        # stale thread: base 30 + min(80/4=20, 20) = 50, exactly 50 -> important
        # Use age_hours=72+eps so threshold triggers, but compute carefully
        # age_hours=76: 30 + min(76/4=19, 20) = 49 -> informational
        thread = _thread_info(age_hours=76.0)
        items = score_priorities(
            _make_email_signals(),
            _make_calendar_signals(),
            _make_thread_signals(awaiting_their_reply=[thread]),
        )
        assert len(items) == 1
        assert items[0]["score"] == 49.0
        assert items[0]["score"] < 50


# ---------------------------------------------------------------------------
# Engine tests (generate_* functions)
# ---------------------------------------------------------------------------


def _make_engine_request(
    folders: list | None = None,
    inbox_msgs: list | None = None,
    cal_events: list | None = None,
    sent_count: int = 0,
    contacts: list | None = None,
) -> MagicMock:
    """Build a fake request callable covering all collector paths."""
    folder_data = folders or []
    inbox_data = inbox_msgs or []
    cal_data = cal_events or []
    contact_data = contacts or []

    def fake(method: str, path: str, *, params: dict | None = None, **_: Any) -> dict:
        if "/mailFolders" in path and "messages" not in path:
            return {"value": folder_data}
        if "calendarView" in path:
            return {"value": cal_data}
        if "sentItems" in path or "sentitems" in path:
            return {"value": [], "@odata.count": sent_count}
        if "/me/contacts" in path:
            return {"value": contact_data}
        return {"value": inbox_data}

    return MagicMock(side_effect=fake)


class TestGenerateBriefing:
    def test_returns_briefing_report_keys(self) -> None:
        request = _make_engine_request()
        report = generate_briefing(request, account="user@example.com", now=NOW)
        assert "generated_at" in report
        assert "account" in report
        assert "priority_items" in report
        assert "email_summary" in report
        assert "calendar_summary" in report
        assert "schedule_analysis" in report
        assert "thread_summary" in report

    def test_account_matches(self) -> None:
        request = _make_engine_request()
        report = generate_briefing(request, account="user@example.com", now=NOW)
        assert report["account"] == "user@example.com"

    def test_generated_at_uses_injected_now(self) -> None:
        request = _make_engine_request()
        report = generate_briefing(request, account="user@example.com", now=NOW)
        assert report["generated_at"] == NOW.isoformat()

    def test_priority_items_sorted_descending(self) -> None:
        # Provide a folder with unread to get some signals
        folder = {"displayName": "Inbox", "unreadItemCount": 2, "totalItemCount": 10}
        request = _make_engine_request(folders=[folder])
        report = generate_briefing(request, account="user@example.com", now=NOW)
        scores = [item["score"] for item in report["priority_items"]]
        assert scores == sorted(scores, reverse=True)


class TestGenerateSignals:
    def test_returns_signals_report_keys(self) -> None:
        request = _make_engine_request()
        report = generate_signals(request, account="user@example.com", now=NOW)
        assert "critical" in report
        assert "important" in report
        assert "informational" in report
        assert "total_signals" in report

    def test_total_signals_sum(self) -> None:
        request = _make_engine_request()
        report = generate_signals(request, account="user@example.com", now=NOW)
        assert report["total_signals"] == (
            len(report["critical"])
            + len(report["important"])
            + len(report["informational"])
        )

    def test_critical_items_all_gte_80(self) -> None:
        request = _make_engine_request()
        report = generate_signals(request, account="user@example.com", now=NOW)
        for item in report["critical"]:
            assert item["score"] >= 80

    def test_important_items_in_range(self) -> None:
        request = _make_engine_request()
        report = generate_signals(request, account="user@example.com", now=NOW)
        for item in report["important"]:
            assert 50 <= item["score"] < 80

    def test_informational_items_below_50(self) -> None:
        request = _make_engine_request()
        report = generate_signals(request, account="user@example.com", now=NOW)
        for item in report["informational"]:
            assert item["score"] < 50


class TestGenerateRecap:
    def test_returns_recap_report_keys(self) -> None:
        request = _make_engine_request()
        report = generate_recap(request, account="user@example.com", now=NOW)
        assert "emails_received_today" in report
        assert "emails_sent_today" in report
        assert "emails_still_unread" in report
        assert "meetings_attended" in report
        assert "threads_resolved" in report
        assert "threads_still_pending" in report
        assert "tomorrow_preview" in report

    def test_threads_resolved_zero(self) -> None:
        request = _make_engine_request()
        report = generate_recap(request, account="user@example.com", now=NOW)
        assert report["threads_resolved"] == 0

    def test_meetings_attended_counts_accepted(self) -> None:
        # Calendar event with accepted response
        cal_event_raw = {
            "id": "e1",
            "subject": "Standup",
            "start": {"dateTime": "2026-06-17T09:00:00Z"},
            "end": {"dateTime": "2026-06-17T09:30:00Z"},
            "location": {"displayName": ""},
            "isAllDay": False,
            "isOnlineMeeting": False,
            "organizer": {"emailAddress": {"name": "Me", "address": "me@example.com"}},
            "attendees": [],
            "responseStatus": {"response": "accepted"},
            "showAs": "busy",
        }
        request = _make_engine_request(cal_events=[cal_event_raw])
        report = generate_recap(request, account="user@example.com", now=NOW)
        assert report["meetings_attended"] == 1  # one accepted event on 2026-06-17


class TestGenerateContactReport:
    def test_returns_contact_report_keys(self) -> None:
        request = _make_engine_request()
        report = generate_contact_report(
            request,
            account="user@example.com",
            target_email="alice@example.com",
            now=NOW,
        )
        assert "target_email" in report
        assert "target_name" in report
        assert "relationship" in report
        assert "recent_threads" in report
        assert "pending_items" in report
        assert "recent_emails_from" in report
        assert "recent_emails_to" in report

    def test_target_email_preserved(self) -> None:
        request = _make_engine_request()
        report = generate_contact_report(
            request,
            account="user@example.com",
            target_email="Alice@Example.COM",
            now=NOW,
        )
        assert report["target_email"] == "Alice@Example.COM"

    def test_missing_contact_gets_zero_interactions(self) -> None:
        request = _make_engine_request()
        report = generate_contact_report(
            request,
            account="user@example.com",
            target_email="unknown@example.com",
            now=NOW,
        )
        assert report["recent_emails_from"] == 0
        assert report["recent_emails_to"] == 0

    def test_relationship_insight_present(self) -> None:
        request = _make_engine_request()
        report = generate_contact_report(
            request,
            account="user@example.com",
            target_email="alice@example.com",
            now=NOW,
        )
        rel = report["relationship"]
        assert "engagement_score" in rel
        assert "trend" in rel
