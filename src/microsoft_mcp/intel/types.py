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


# ============================================================================
# Analyzer Output Types
# ============================================================================


class PriorityItem(TypedDict):
    """A scored, prioritized item for the briefing."""

    score: float  # 0.0 to 100.0
    source: str  # "email", "calendar", "thread"
    category: str  # "needs_response", "conflict", "vip_email", "stale_thread", etc.
    title: str
    description: str
    age_hours: NotRequired[float]
    sender: NotRequired[str]
    action_hint: NotRequired[str]  # Suggested action


class ScheduleAnalysis(TypedDict):
    """Analyzed schedule insights."""

    meeting_density_pct: float  # % of work day in meetings
    back_to_back_count: int  # Number of back-to-back meeting pairs
    external_meeting_count: int
    longest_free_block_minutes: int
    focus_time_available: bool  # Has a 2+ hour free block
    busiest_hour: NotRequired[str]  # e.g., "10:00-11:00"
    summary: str  # Human-readable schedule summary


class RelationshipInsight(TypedDict):
    """Analyzed relationship status for a contact."""

    email: str
    name: str
    company: NotRequired[str]
    engagement_score: float  # 0.0 to 100.0
    trend: str  # "rising", "stable", "cooling"
    last_interaction: str  # ISO datetime
    days_since_contact: int
    sent_to: int
    received_from: int
    response_ratio: float  # ratio of received/sent, >1 means they contact you more


# ============================================================================
# Report Types (Engine Output)
# ============================================================================


class BriefingReport(TypedDict):
    """Complete morning briefing output."""

    generated_at: str  # ISO datetime
    account: str
    priority_items: list[PriorityItem]  # Top items sorted by score
    email_summary: EmailSignals
    calendar_summary: CalendarSignals
    schedule_analysis: ScheduleAnalysis
    thread_summary: ThreadSignals


class SignalsReport(TypedDict):
    """Actionable signals and alerts."""

    generated_at: str
    account: str
    critical: list[PriorityItem]  # Score >= 80
    important: list[PriorityItem]  # Score >= 50
    informational: list[PriorityItem]  # Score < 50
    total_signals: int


class ContactReport(TypedDict):
    """Person-level intelligence report."""

    generated_at: str
    account: str
    target_email: str
    target_name: str
    company: NotRequired[str]
    job_title: NotRequired[str]
    relationship: RelationshipInsight
    recent_threads: list[ThreadInfo]
    recent_emails_from: int
    recent_emails_to: int
    pending_items: list[EmailNeedingResponse]


class RecapReport(TypedDict):
    """End-of-day summary."""

    generated_at: str
    account: str
    emails_received_today: int
    emails_sent_today: int
    emails_still_unread: int
    meetings_attended: int
    threads_resolved: int  # Threads that got replies today
    threads_still_pending: list[ThreadInfo]
    tomorrow_preview: list[CalendarEvent]
