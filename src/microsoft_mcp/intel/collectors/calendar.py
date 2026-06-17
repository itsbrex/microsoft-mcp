"""Calendar signal collector for the intelligence layer.

Gathers calendar events from the Microsoft Graph API calendarView endpoint
and derives intelligence signals: conflicts, free blocks, prep needed, and
meeting load metrics.

The ``request`` callable is dependency-injected so this module has no direct
import of ``graph``, keeping it pure and unit-testable.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import UTC, datetime, timedelta
from typing import Any
from zoneinfo import ZoneInfo

from microsoft_mcp.intel.types import (
    CalendarEvent,
    CalendarSignals,
    ConflictPair,
    TimeBlock,
)

logger = logging.getLogger(__name__)

_WORK_DAY_START_HOUR = 8
_WORK_DAY_END_HOUR = 18
_MIN_FREE_BLOCK_MINUTES = 15
_MAX_EVENTS_PER_DAY = 50


def _parse_event_datetime(dt_str: str) -> datetime | None:
    """Parse naive ISO datetime string from Graph calendarView response."""
    if not dt_str:
        return None
    try:
        return datetime.fromisoformat(dt_str)
    except ValueError:
        logger.warning("Failed to parse event datetime: %s", dt_str)
        return None


def _extract_domain(email: str) -> str:
    if "@" in email:
        return email.rsplit("@", 1)[1].lower()
    return ""


def _convert_event(event: dict[str, Any]) -> CalendarEvent:
    organizer_addr = (event.get("organizer") or {}).get("emailAddress") or {}
    organizer_email = organizer_addr.get("address", "")
    organizer_domain = _extract_domain(organizer_email)
    attendees: list[dict[str, Any]] = event.get("attendees") or []
    has_external = False
    if organizer_domain:
        for att in attendees:
            att_email = (att.get("emailAddress") or {}).get("address", "")
            if att_email and _extract_domain(att_email) != organizer_domain:
                has_external = True
                break
    return CalendarEvent(
        id=event.get("id", ""),
        subject=event.get("subject", ""),
        start=(event.get("start") or {}).get("dateTime", ""),
        end=(event.get("end") or {}).get("dateTime", ""),
        location=(event.get("location") or {}).get("displayName", ""),
        is_all_day=event.get("isAllDay", False),
        is_online=event.get("isOnlineMeeting", False),
        organizer_name=organizer_addr.get("name", ""),
        organizer_email=organizer_email,
        attendee_count=len(attendees),
        has_external_attendees=has_external,
        response_status=(event.get("responseStatus") or {}).get("response", "none"),
        show_as=event.get("showAs", "busy"),
    )


def _fetch_calendar_view(
    request: Callable[..., Any],
    start: datetime,
    end: datetime,
    tz_name: str,
) -> list[CalendarEvent]:
    data = request(
        "GET",
        "/me/calendarView",
        params={
            "startDateTime": start.astimezone(UTC).strftime(
                "%Y-%m-%dT%H:%M:%S.0000000Z"
            ),
            "endDateTime": end.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%S.0000000Z"),
            "$orderby": "start/dateTime",
            "$top": _MAX_EVENTS_PER_DAY,
            "$select": (
                "id,subject,start,end,location,isAllDay,isOnlineMeeting,"
                "organizer,attendees,responseStatus,showAs"
            ),
            "Prefer": f'outlook.timezone="{tz_name}"',
        },
    )
    return [_convert_event(e) for e in data.get("value", [])]


def _detect_conflicts(events: list[CalendarEvent]) -> list[ConflictPair]:
    conflicts: list[ConflictPair] = []
    timed = [e for e in events if not e["is_all_day"]]
    for i in range(len(timed)):
        end_a = _parse_event_datetime(timed[i]["end"])
        if end_a is None:
            continue
        for j in range(i + 1, len(timed)):
            start_b = _parse_event_datetime(timed[j]["start"])
            end_b = _parse_event_datetime(timed[j]["end"])
            if start_b is None or end_b is None:
                continue
            if end_a > start_b:
                overlap_minutes = int(
                    (min(end_a, end_b) - start_b).total_seconds() / 60
                )
                if overlap_minutes > 0:
                    conflicts.append(
                        ConflictPair(
                            event_a_subject=timed[i]["subject"],
                            event_a_start=timed[i]["start"],
                            event_b_subject=timed[j]["subject"],
                            event_b_start=timed[j]["start"],
                            overlap_minutes=overlap_minutes,
                        )
                    )
            else:
                break
    return conflicts


def _find_prep_needed(events: list[CalendarEvent]) -> list[CalendarEvent]:
    return [
        e for e in events if e["has_external_attendees"] or e["attendee_count"] >= 3
    ]


def _find_free_blocks(
    events: list[CalendarEvent], day_start: datetime
) -> list[TimeBlock]:
    """Find free blocks during working hours (08:00–18:00).

    Event datetimes are naive when Graph uses Prefer:outlook.timezone, so
    day_start tzinfo is stripped for comparison.
    """
    work_start = day_start.replace(
        hour=_WORK_DAY_START_HOUR, minute=0, second=0, microsecond=0, tzinfo=None
    )
    work_end = day_start.replace(
        hour=_WORK_DAY_END_HOUR, minute=0, second=0, microsecond=0, tzinfo=None
    )
    busy: list[tuple[datetime, datetime]] = []
    for event in events:
        if event["is_all_day"]:
            continue
        s = _parse_event_datetime(event["start"])
        e = _parse_event_datetime(event["end"])
        if s is None or e is None:
            continue
        cs, ce = max(s, work_start), min(e, work_end)
        if cs < ce:
            busy.append((cs, ce))
    busy.sort(key=lambda p: p[0])
    merged: list[tuple[datetime, datetime]] = []
    for s, e in busy:
        if merged and s <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], e))
        else:
            merged.append((s, e))
    free_blocks: list[TimeBlock] = []
    cursor = work_start
    for bs, be in merged:
        if bs > cursor:
            gap = int((bs - cursor).total_seconds() / 60)
            if gap >= _MIN_FREE_BLOCK_MINUTES:
                free_blocks.append(
                    TimeBlock(
                        start=cursor.isoformat(),
                        end=bs.isoformat(),
                        duration_minutes=gap,
                    )
                )
        cursor = max(cursor, be)
    if cursor < work_end:
        gap = int((work_end - cursor).total_seconds() / 60)
        if gap >= _MIN_FREE_BLOCK_MINUTES:
            free_blocks.append(
                TimeBlock(
                    start=cursor.isoformat(),
                    end=work_end.isoformat(),
                    duration_minutes=gap,
                )
            )
    return free_blocks


def _calculate_meeting_hours(events: list[CalendarEvent]) -> float:
    total = 0.0
    for event in events:
        if event["is_all_day"]:
            continue
        s = _parse_event_datetime(event["start"])
        e = _parse_event_datetime(event["end"])
        if s is not None and e is not None and e > s:
            total += (e - s).total_seconds() / 60
    return total / 60


def collect_calendar_signals(
    request: Callable[..., Any],
    *,
    now: datetime,
    timezone: str = "UTC",
) -> CalendarSignals:
    """Collect calendar intelligence signals.

    Args:
        request: Injected Graph request callable — signature
            ``request(method, path, *, params=None, json=None) -> dict``.
        now: Current datetime (injected for deterministic testing; never call
            ``datetime.now()`` inside this function).
        timezone: IANA timezone string (e.g. ``"America/Chicago"``).

    Returns:
        :class:`CalendarSignals` with today's events, conflicts, free blocks,
        prep-needed events, and meeting load metrics.
    """
    tz = ZoneInfo(timezone)
    now_local = now.astimezone(tz)
    today_start = now_local.replace(hour=0, minute=0, second=0, microsecond=0)
    today_end = today_start + timedelta(days=1)
    tomorrow_start = today_end
    tomorrow_end = tomorrow_start + timedelta(days=1)

    today_events = _fetch_calendar_view(request, today_start, today_end, timezone)
    tomorrow_events = _fetch_calendar_view(
        request, tomorrow_start, tomorrow_end, timezone
    )

    return CalendarSignals(
        today_events=today_events,
        tomorrow_events=tomorrow_events,
        conflicts=_detect_conflicts(today_events),
        prep_needed=_find_prep_needed(today_events),
        free_blocks=_find_free_blocks(today_events, today_start),
        meeting_hours_today=round(_calculate_meeting_hours(today_events), 2),
        total_events_today=len(today_events),
    )
