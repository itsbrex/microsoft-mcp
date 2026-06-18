"""Schedule analyzer for the intelligence layer.

Analyzes calendar signals to produce schedule insights: meeting density,
back-to-back detection, external meeting counts, free block analysis,
busiest hour identification, and a human-readable summary.
"""

from __future__ import annotations

import logging
from collections import Counter

from microsoft_mcp.intel._utils import parse_graph_datetime as _parse_dt
from microsoft_mcp.intel.types import (
    CalendarEvent,
    CalendarSignals,
    ScheduleAnalysis,
)

logger = logging.getLogger(__name__)

# Work day boundaries (8am-6pm = 10 hours)
_WORK_DAY_START_HOUR = 8
_WORK_DAY_END_HOUR = 18
_WORK_DAY_HOURS = _WORK_DAY_END_HOUR - _WORK_DAY_START_HOUR

# Back-to-back threshold: events with less than 15 minutes gap
_BACK_TO_BACK_GAP_MINUTES = 15

# Minimum free block duration (minutes) to qualify as focus time
_FOCUS_TIME_THRESHOLD_MINUTES = 120

# Meeting density thresholds for summary labels
_DENSITY_LIGHT = 30.0
_DENSITY_MODERATE = 60.0
_DENSITY_BUSY = 80.0


def _non_all_day_events(events: list[CalendarEvent]) -> list[CalendarEvent]:
    """Filter out all-day events, returning only timed events."""
    return [e for e in events if not e["is_all_day"]]


def _count_back_to_back(events: list[CalendarEvent]) -> int:
    """Count pairs of consecutive events with < 15 minutes gap.

    Only considers non-all-day events. Events are sorted by start time
    before checking gaps.

    Args:
        events: Today's calendar events.

    Returns:
        Number of back-to-back meeting pairs.
    """
    timed = _non_all_day_events(events)
    if len(timed) < 2:
        return 0

    sorted_events = sorted(timed, key=lambda e: e["start"])

    count = 0
    for i in range(len(sorted_events) - 1):
        end_a = _parse_dt(sorted_events[i]["end"])
        start_b = _parse_dt(sorted_events[i + 1]["start"])
        gap_minutes = (start_b - end_a).total_seconds() / 60.0
        if gap_minutes < _BACK_TO_BACK_GAP_MINUTES:
            count += 1

    return count


def _count_external_meetings(events: list[CalendarEvent]) -> int:
    """Count events that have external attendees.

    Args:
        events: Today's calendar events.

    Returns:
        Number of events with external attendees.
    """
    return sum(1 for e in events if e["has_external_attendees"])


def _find_longest_free_block(calendar: CalendarSignals) -> int:
    """Find the longest free block duration in minutes.

    Args:
        calendar: Collected calendar signals containing free_blocks.

    Returns:
        Duration in minutes of the longest free block, or 0 if none.
    """
    if not calendar["free_blocks"]:
        return 0
    return max(block["duration_minutes"] for block in calendar["free_blocks"])


def _find_busiest_hour(events: list[CalendarEvent]) -> str | None:
    """Find the 1-hour work day slot with the most overlapping events.

    Divides the work day (8am-6pm) into 1-hour slots and counts how many
    non-all-day events overlap each slot. An event overlaps a slot if its
    time range intersects the slot interval.

    Args:
        events: Today's calendar events.

    Returns:
        The busiest hour as "HH:00-HH:00" (e.g., "10:00-11:00"),
        or None if no non-all-day events exist or there is a tie for
        the maximum.
    """
    timed = _non_all_day_events(events)
    if not timed:
        return None

    slot_counts: Counter[int] = Counter()

    for hour in range(_WORK_DAY_START_HOUR, _WORK_DAY_END_HOUR):
        for event in timed:
            event_start = _parse_dt(event["start"])
            event_end = _parse_dt(event["end"])

            slot_start = event_start.replace(
                hour=hour,
                minute=0,
                second=0,
                microsecond=0,
            )
            slot_end = event_start.replace(
                hour=hour + 1,
                minute=0,
                second=0,
                microsecond=0,
            )

            if event_start < slot_end and event_end > slot_start:
                slot_counts[hour] += 1

    if not slot_counts:
        return None

    max_count = max(slot_counts.values())
    busiest_slots = [h for h, c in slot_counts.items() if c == max_count]
    if len(busiest_slots) > 1:
        return None

    busiest = busiest_slots[0]
    return f"{busiest:02d}:00-{busiest + 1:02d}:00"


def _density_label(pct: float) -> str:
    """Return a human-readable label for the meeting density percentage.

    Thresholds:
        <30% = "Light", 30-60% = "Moderate", 60-80% = "Busy", >80% = "Packed"
    """
    if pct < _DENSITY_LIGHT:
        return "Light"
    if pct < _DENSITY_MODERATE:
        return "Moderate"
    if pct < _DENSITY_BUSY:
        return "Busy"
    return "Packed"


def _build_summary(
    total_events: int,
    density_pct: float,
    back_to_back: int,
    external_count: int,
    longest_free_minutes: int,
    focus_available: bool,
) -> str:
    """Build a human-readable schedule summary.

    Examples:
        "Light day: 2 meetings (20%), 4h focus time available"
        "Busy day: 8 meetings (75%), 2 back-to-back, no focus blocks"
        "Packed day: 12 meetings (95%), 5 back-to-back, 3 external"

    Args:
        total_events: Total number of events today.
        density_pct: Meeting density as a percentage.
        back_to_back: Number of back-to-back meeting pairs.
        external_count: Number of meetings with external attendees.
        longest_free_minutes: Duration of the longest free block in minutes.
        focus_available: Whether a 2+ hour free block exists.

    Returns:
        Human-readable summary string.
    """
    label = _density_label(density_pct)
    meeting_word = "meeting" if total_events == 1 else "meetings"
    parts = [f"{label} day: {total_events} {meeting_word} ({density_pct:.0f}%)"]

    if back_to_back > 0:
        parts.append(f"{back_to_back} back-to-back")

    if external_count > 0:
        parts.append(f"{external_count} external")

    if focus_available:
        focus_hours = longest_free_minutes / 60.0
        if focus_hours == int(focus_hours):
            parts.append(f"{int(focus_hours)}h focus time available")
        else:
            parts.append(f"{focus_hours:.1f}h focus time available")
    else:
        parts.append("no focus blocks")

    return ", ".join(parts)


def analyze_schedule(calendar: CalendarSignals) -> ScheduleAnalysis:
    """Analyze today's schedule for insights.

    Computes meeting density, detects back-to-back meetings, counts
    external meetings, identifies the longest free block and busiest hour,
    and generates a human-readable summary.

    Args:
        calendar: Collected calendar signals.

    Returns:
        ScheduleAnalysis with meeting density, back-to-back detection,
        free block analysis, busiest hour, and summary.
    """
    today_events = calendar["today_events"]

    density_pct = min(
        (calendar["meeting_hours_today"] / _WORK_DAY_HOURS) * 100.0,
        100.0,
    )

    back_to_back = _count_back_to_back(today_events)
    external_count = _count_external_meetings(today_events)
    longest_free = _find_longest_free_block(calendar)
    focus_available = longest_free >= _FOCUS_TIME_THRESHOLD_MINUTES
    busiest = _find_busiest_hour(today_events)

    summary = _build_summary(
        total_events=calendar["total_events_today"],
        density_pct=density_pct,
        back_to_back=back_to_back,
        external_count=external_count,
        longest_free_minutes=longest_free,
        focus_available=focus_available,
    )

    logger.debug(
        "Schedule analysis: density=%.1f%%, b2b=%d, external=%d, longest_free=%dm, focus=%s",
        density_pct,
        back_to_back,
        external_count,
        longest_free,
        focus_available,
    )

    result = ScheduleAnalysis(
        meeting_density_pct=round(density_pct, 1),
        back_to_back_count=back_to_back,
        external_meeting_count=external_count,
        longest_free_block_minutes=longest_free,
        focus_time_available=focus_available,
        summary=summary,
    )

    if busiest is not None:
        result["busiest_hour"] = busiest

    return result
