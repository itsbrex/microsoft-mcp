"""Priority scoring analyzer.

Takes collected email, calendar, and thread signals and scores each
item by urgency/importance, producing a ranked list of PriorityItems
for use in briefings and alert reports.
"""

from __future__ import annotations

from microsoft_mcp.intel.types import (
    CalendarEvent,
    CalendarSignals,
    ConflictPair,
    EmailNeedingResponse,
    EmailSignals,
    PriorityItem,
    ThreadInfo,
    ThreadSignals,
    VipEmail,
)

_MAX_SCORE = 100.0
_STALE_THREAD_THRESHOLD_HOURS = 72.0


def score_priorities(
    emails: EmailSignals,
    calendar: CalendarSignals,
    threads: ThreadSignals,
) -> list[PriorityItem]:
    """Score and rank all signals by priority.

    Collects items from email, calendar, and thread signals, assigns a
    numeric urgency score (0.0--100.0) to each, and returns them sorted
    by score descending.

    Args:
        emails: Collected email signals.
        calendar: Collected calendar signals.
        threads: Collected thread signals.

    Returns:
        List of PriorityItem sorted by score descending.
    """
    items: list[PriorityItem] = []

    # -- Email signals ---------------------------------------------------------
    for email in emails["needs_response"]:
        items.append(_score_needs_response(email))

    for vip in emails["vip_unread"]:
        items.append(_score_vip_email(vip))

    # -- Calendar signals ------------------------------------------------------
    for conflict in calendar["conflicts"]:
        items.append(_score_conflict(conflict))

    for event in calendar["prep_needed"]:
        items.append(_score_prep_needed(event))

    # -- Thread signals --------------------------------------------------------
    for thread in threads["awaiting_my_reply"]:
        items.append(_score_awaiting_my_reply(thread))

    for thread in threads["awaiting_their_reply"]:
        if thread["age_hours"] > _STALE_THREAD_THRESHOLD_HOURS:
            items.append(_score_stale_thread(thread))

    items.sort(key=lambda item: item["score"], reverse=True)
    return items


# =============================================================================
# Individual scoring functions
# =============================================================================


def _cap(score: float) -> float:
    """Clamp a score to [0.0, 100.0]."""
    return min(max(score, 0.0), _MAX_SCORE)


def _score_needs_response(email: EmailNeedingResponse) -> PriorityItem:
    """Score an email needing a response.

    Base 60, +20 if high importance, +10 if has attachments,
    +min(age_hours, 10) so older unanswered emails rank higher.
    """
    score = 60.0
    if email["importance"] == "high":
        score += 20.0
    if email["has_attachments"]:
        score += 10.0
    score += min(email["age_hours"], 10.0)

    return PriorityItem(
        score=_cap(score),
        source="email",
        category="needs_response",
        title=f"Reply needed: {email['subject']}",
        description=(f"{email['sender_name']} sent this {email['age_hours']:.1f}h ago"),
        age_hours=email["age_hours"],
        sender=email["sender_name"],
        action_hint=f"Reply to {email['sender_name']}",
    )


def _score_vip_email(vip: VipEmail) -> PriorityItem:
    """Score an unread email from a high-interaction contact.

    Base 50, +5 per interaction_count (capped at +25),
    +15 if high importance.
    """
    score = 50.0
    score += min(vip["interaction_count"] * 5.0, 25.0)
    if vip["importance"] == "high":
        score += 15.0

    return PriorityItem(
        score=_cap(score),
        source="email",
        category="vip_email",
        title=f"VIP: {vip['subject']}",
        description=(
            f"Unread email from {vip['sender_name']} "
            f"({vip['interaction_count']} recent interactions)"
        ),
        sender=vip["sender_name"],
        action_hint=f"Read email from {vip['sender_name']}",
    )


def _score_conflict(conflict: ConflictPair) -> PriorityItem:
    """Score a calendar conflict (always critical, score 90)."""
    return PriorityItem(
        score=90.0,
        source="calendar",
        category="conflict",
        title="Schedule conflict",
        description=(
            f"{conflict['event_a_subject']} overlaps with {conflict['event_b_subject']}"
        ),
        action_hint="Resolve conflict",
    )


def _score_prep_needed(event: CalendarEvent) -> PriorityItem:
    """Score a meeting needing preparation.

    Base 40, +20 if has external attendees, +10 if 5+ attendees.
    """
    score = 40.0
    if event["has_external_attendees"]:
        score += 20.0
    if event["attendee_count"] >= 5:
        score += 10.0

    return PriorityItem(
        score=_cap(score),
        source="calendar",
        category="prep_needed",
        title=f"Prep needed: {event['subject']}",
        description=(
            f"{event['subject']} at {event['start']} with {event['attendee_count']} attendees"
        ),
        action_hint=f"Prepare for {event['subject']}",
    )


def _score_awaiting_my_reply(thread: ThreadInfo) -> PriorityItem:
    """Score a thread awaiting the user's reply.

    Base 55, +20 if high importance, +min(age_hours / 2, 15).
    """
    score = 55.0
    if thread["importance"] == "high":
        score += 20.0
    score += min(thread["age_hours"] / 2.0, 15.0)

    return PriorityItem(
        score=_cap(score),
        source="thread",
        category="needs_response",
        title=f"Thread reply needed: {thread['subject']}",
        description=(
            f"{thread['last_sender_name']} is waiting for a reply ({thread['age_hours']:.1f}h)"
        ),
        age_hours=thread["age_hours"],
        sender=thread["last_sender_name"],
        action_hint=f"Reply to thread from {thread['last_sender_name']}",
    )


def _score_stale_thread(thread: ThreadInfo) -> PriorityItem:
    """Score a stale thread (awaiting their reply, age > 72h).

    Base 30, +min(age_hours / 4, 20).
    """
    score = 30.0
    score += min(thread["age_hours"] / 4.0, 20.0)

    return PriorityItem(
        score=_cap(score),
        source="thread",
        category="stale_thread",
        title=f"Follow up: {thread['subject']}",
        description=(
            f"No response from {thread['last_sender_name']} in {thread['age_hours']:.1f}h"
        ),
        age_hours=thread["age_hours"],
        sender=thread["last_sender_name"],
        action_hint=f"Follow up with {thread['last_sender_name']}",
    )
