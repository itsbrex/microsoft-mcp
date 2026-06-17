"""Type definitions for the intelligence layer.

All TypedDict definitions for signals produced by collectors.
Field names are kept aligned with the outlook-creds reference so that
Task 8.4/8.5 analyzers and MCP tools can consume them without adaptation.
"""

from __future__ import annotations

from typing import NotRequired, TypedDict


# ============================================================================
# Collector Output Types
# ============================================================================


class FolderUnreadCount(TypedDict):
    """Unread email count for a mail folder."""

    folder_name: str
    unread_count: int
    total_count: int


class EmailNeedingResponse(TypedDict):
    """An email that likely needs the user's response."""

    id: str
    subject: str
    sender_name: str
    sender_email: str
    received_at: str  # ISO datetime
    age_hours: float
    importance: str  # low, normal, high
    has_attachments: bool
    body_preview: str


class VipEmail(TypedDict):
    """An email from a high-interaction contact."""

    id: str
    subject: str
    sender_name: str
    sender_email: str
    received_at: str
    is_read: bool
    importance: str
    interaction_count: int  # How often user interacts with this sender


class EmailSignals(TypedDict):
    """Collected email intelligence signals."""

    unread_total: int
    unread_by_folder: list[FolderUnreadCount]
    needs_response: list[EmailNeedingResponse]
    vip_unread: list[VipEmail]
    received_last_24h: int
    sent_last_24h: int


class CalendarEvent(TypedDict):
    """A calendar event with intelligence metadata."""

    id: str
    subject: str
    start: str  # ISO datetime
    end: str  # ISO datetime
    location: str
    is_all_day: bool
    is_online: bool
    organizer_name: str
    organizer_email: str
    attendee_count: int
    has_external_attendees: bool
    response_status: str  # none, organizer, tentativelyAccepted, accepted, declined
    show_as: str  # free, tentative, busy, oof, workingElsewhere


class TimeBlock(TypedDict):
    """A block of free time between events."""

    start: str  # ISO datetime
    end: str  # ISO datetime
    duration_minutes: int


class ConflictPair(TypedDict):
    """Two events that overlap in time."""

    event_a_subject: str
    event_a_start: str
    event_b_subject: str
    event_b_start: str
    overlap_minutes: int


class CalendarSignals(TypedDict):
    """Collected calendar intelligence signals."""

    today_events: list[CalendarEvent]
    tomorrow_events: list[CalendarEvent]
    conflicts: list[ConflictPair]
    prep_needed: list[CalendarEvent]  # External meetings needing prep
    free_blocks: list[TimeBlock]
    meeting_hours_today: float
    total_events_today: int


class ContactInteraction(TypedDict):
    """Contact with recent interaction statistics."""

    email: str
    name: str
    company: NotRequired[str]
    job_title: NotRequired[str]
    total_interactions: int
    sent_to: int
    received_from: int
    last_interaction: str  # ISO datetime
    has_pending_email: bool  # Unreplied email from them


class ContactSignals(TypedDict):
    """Collected contact interaction signals."""

    top_contacts: list[ContactInteraction]  # By interaction count
    pending_contacts: list[ContactInteraction]  # Unreplied emails
    total_unique_senders: int
    total_unique_recipients: int


class ThreadInfo(TypedDict):
    """A conversation thread requiring attention."""

    conversation_id: str
    subject: str
    last_sender_name: str
    last_sender_email: str
    last_message_at: str  # ISO datetime
    age_hours: float
    message_count: int
    direction: str  # "inbound" (they sent last) or "outbound" (I sent last)
    importance: str


class ThreadSignals(TypedDict):
    """Collected thread tracking signals."""

    awaiting_my_reply: list[ThreadInfo]  # They sent last, I haven't replied
    awaiting_their_reply: list[ThreadInfo]  # I sent last, no response
    stale_threads: list[ThreadInfo]  # Old threads with no activity
