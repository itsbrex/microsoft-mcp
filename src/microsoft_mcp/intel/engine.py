"""Engine orchestrator for the intelligence layer.

Ties together collectors and analyzers to produce final reports:
briefing, signals, contact, and recap.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import datetime
from typing import Any

from microsoft_mcp.intel.analyzers.priority import score_priorities
from microsoft_mcp.intel.analyzers.relationships import analyze_relationships
from microsoft_mcp.intel.analyzers.schedule import analyze_schedule
from microsoft_mcp.intel.collectors.calendar import collect_calendar_signals
from microsoft_mcp.intel.collectors.contacts import collect_contact_signals
from microsoft_mcp.intel.collectors.email import collect_email_signals
from microsoft_mcp.intel.collectors.threads import collect_thread_signals
from microsoft_mcp.intel.types import (
    BriefingReport,
    CalendarSignals,
    ContactInteraction,
    ContactReport,
    EmailSignals,
    RecapReport,
    RelationshipInsight,
    SignalsReport,
    ThreadInfo,
    ThreadSignals,
)

logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Shared collection helper
# ---------------------------------------------------------------------------


def _collect_all_signals(
    request: Callable[..., Any],
    *,
    now: datetime,
    timezone: str,
    email_hours: int = 24,
    thread_hours: int = 48,
) -> tuple[EmailSignals, CalendarSignals, ThreadSignals]:
    """Run all three collectors and return their results."""
    email_signals = collect_email_signals(request, now=now, lookback_hours=email_hours)
    calendar_signals = collect_calendar_signals(request, now=now, timezone=timezone)
    thread_signals = collect_thread_signals(
        request, now=now, lookback_hours=thread_hours
    )
    return email_signals, calendar_signals, thread_signals


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def generate_briefing(
    request: Callable[..., Any],
    *,
    account: str,
    timezone: str = "UTC",
    now: datetime,
) -> BriefingReport:
    """Generate a comprehensive morning briefing.

    Runs all collectors, scores priorities, and analyzes the schedule
    to produce a single report covering email, calendar, and thread
    activity.

    Args:
        request: Injected Graph request callable.
        account: Account email address.
        timezone: User's IANA timezone (e.g., "America/Chicago").
        now: Current UTC datetime (injected for deterministic testing).

    Returns:
        Complete BriefingReport.
    """
    logger.info("Generating briefing for %s (tz=%s)", account, timezone)

    email_signals, calendar_signals, thread_signals = _collect_all_signals(
        request,
        now=now,
        timezone=timezone,
    )

    priority_items = score_priorities(email_signals, calendar_signals, thread_signals)
    schedule_analysis = analyze_schedule(calendar_signals)

    logger.info(
        "Briefing complete: %d priority items, %d emails needing response, %d events today",
        len(priority_items),
        len(email_signals["needs_response"]),
        calendar_signals["total_events_today"],
    )

    return BriefingReport(
        generated_at=now.isoformat(),
        account=account,
        priority_items=priority_items,
        email_summary=email_signals,
        calendar_summary=calendar_signals,
        schedule_analysis=schedule_analysis,
        thread_summary=thread_signals,
    )


def generate_signals(
    request: Callable[..., Any],
    *,
    account: str,
    timezone: str = "UTC",
    now: datetime,
) -> SignalsReport:
    """Generate actionable signals/alerts.

    Same collection as briefing but categorized by urgency level:
    critical (score >= 80), important (50--79), informational (< 50).

    Args:
        request: Injected Graph request callable.
        account: Account email address.
        timezone: User's IANA timezone.
        now: Current UTC datetime (injected for deterministic testing).

    Returns:
        SignalsReport with items bucketed by urgency.
    """
    logger.info("Generating signals for %s (tz=%s)", account, timezone)

    email_signals, calendar_signals, thread_signals = _collect_all_signals(
        request,
        now=now,
        timezone=timezone,
    )

    priority_items = score_priorities(email_signals, calendar_signals, thread_signals)

    critical = [item for item in priority_items if item["score"] >= 80]
    important = [item for item in priority_items if 50 <= item["score"] < 80]
    informational = [item for item in priority_items if item["score"] < 50]

    logger.info(
        "Signals complete: %d critical, %d important, %d informational",
        len(critical),
        len(important),
        len(informational),
    )

    return SignalsReport(
        generated_at=now.isoformat(),
        account=account,
        critical=critical,
        important=important,
        informational=informational,
        total_signals=len(priority_items),
    )


def generate_contact_report(
    request: Callable[..., Any],
    *,
    account: str,
    target_email: str,
    now: datetime,
    lookback_days: int = 30,
) -> ContactReport:
    """Generate intelligence report for a specific contact.

    Combines contact interaction data, relationship analysis, thread
    tracking, and email response needs into a single person-level report.

    Args:
        request: Injected Graph request callable.
        account: Account email address.
        target_email: Email of the contact to report on.
        now: Current UTC datetime (injected for deterministic testing).
        lookback_days: Look-back window in days (default 30).

    Returns:
        ContactReport with relationship insight, threads, and pending items.
    """
    target_lower = target_email.lower()
    logger.info(
        "Generating contact report for %s (account=%s, days=%d)",
        target_email,
        account,
        lookback_days,
    )

    # 1. Collect contact signals and find the target
    contact_signals = collect_contact_signals(
        request, now=now, lookback_days=lookback_days
    )
    target_contact = _find_contact_in_signals(
        contact_signals.get("top_contacts", []), target_lower
    )

    if target_contact is None:
        logger.debug(
            "Target %s not found in top contacts; building minimal interaction",
            target_email,
        )
        target_contact = ContactInteraction(
            email=target_lower,
            name=target_email,
            total_interactions=0,
            sent_to=0,
            received_from=0,
            last_interaction=now.isoformat(),
            has_pending_email=False,
        )

    # 2. Analyze relationships and find the target insight
    relationship_insights = analyze_relationships(contact_signals, now=now)
    relationship = _find_relationship_insight(relationship_insights, target_lower)

    if relationship is None:
        relationship = RelationshipInsight(
            email=target_lower,
            name=target_contact["name"],
            engagement_score=0.0,
            trend="stable",
            last_interaction=target_contact["last_interaction"],
            days_since_contact=lookback_days,
            sent_to=target_contact["sent_to"],
            received_from=target_contact["received_from"],
            response_ratio=0.0,
        )

    # 3. Collect threads and filter for target
    hours = lookback_days * 24
    thread_signals = collect_thread_signals(request, now=now, lookback_hours=hours)
    recent_threads = _filter_threads_by_email(thread_signals, target_lower)

    # 4. Collect email signals and filter needs_response from target
    email_signals = collect_email_signals(request, now=now, lookback_hours=hours)
    pending_items = [
        email
        for email in email_signals["needs_response"]
        if email["sender_email"].lower() == target_lower
    ]

    # 5. Look up contact details via Graph API
    display_name, company_name, job_title = _lookup_contact_details(
        request, target_email
    )

    target_name = display_name or target_contact["name"] or target_email

    logger.info(
        "Contact report complete for %s: %d threads, %d pending items",
        target_email,
        len(recent_threads),
        len(pending_items),
    )

    report = ContactReport(
        generated_at=now.isoformat(),
        account=account,
        target_email=target_email,
        target_name=target_name,
        relationship=relationship,
        recent_threads=recent_threads,
        recent_emails_from=target_contact["received_from"],
        recent_emails_to=target_contact["sent_to"],
        pending_items=pending_items,
    )

    if company_name:
        report["company"] = company_name
    if job_title:
        report["job_title"] = job_title

    return report


def generate_recap(
    request: Callable[..., Any],
    *,
    account: str,
    timezone: str = "UTC",
    now: datetime,
) -> RecapReport:
    """Generate end-of-day recap.

    Summarizes today's activity (emails received/sent, meetings attended,
    unread count) and previews tomorrow's schedule.

    Args:
        request: Injected Graph request callable.
        account: Account email address.
        timezone: User's IANA timezone.
        now: Current UTC datetime (injected for deterministic testing).

    Returns:
        RecapReport with today's summary and tomorrow's preview.
    """
    logger.info("Generating recap for %s (tz=%s)", account, timezone)

    email_signals, calendar_signals, thread_signals = _collect_all_signals(
        request,
        now=now,
        timezone=timezone,
        thread_hours=24,
    )

    attended_statuses = {"organizer", "accepted"}
    meetings_attended = sum(
        1
        for event in calendar_signals["today_events"]
        if event["response_status"] in attended_statuses
    )

    threads_still_pending = thread_signals["awaiting_my_reply"]
    threads_resolved = 0

    logger.info(
        "Recap complete: %d received, %d sent, %d unread, %d meetings attended",
        email_signals["received_last_24h"],
        email_signals["sent_last_24h"],
        email_signals["unread_total"],
        meetings_attended,
    )

    return RecapReport(
        generated_at=now.isoformat(),
        account=account,
        emails_received_today=email_signals["received_last_24h"],
        emails_sent_today=email_signals["sent_last_24h"],
        emails_still_unread=email_signals["unread_total"],
        meetings_attended=meetings_attended,
        threads_resolved=threads_resolved,
        threads_still_pending=threads_still_pending,
        tomorrow_preview=calendar_signals["tomorrow_events"],
    )


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _find_contact_in_signals(
    top_contacts: list[ContactInteraction],
    target_lower: str,
) -> ContactInteraction | None:
    """Find a contact by email in the top contacts list."""
    for contact in top_contacts:
        if contact["email"].lower() == target_lower:
            return contact
    return None


def _find_relationship_insight(
    insights: list[RelationshipInsight],
    target_lower: str,
) -> RelationshipInsight | None:
    """Find a relationship insight by email."""
    for insight in insights:
        if insight["email"].lower() == target_lower:
            return insight
    return None


def _filter_threads_by_email(
    thread_signals: ThreadSignals,
    target_lower: str,
) -> list[ThreadInfo]:
    """Filter threads that involve the target email address.

    Checks the ``last_sender_email`` field across all thread categories
    (awaiting_my_reply, awaiting_their_reply, stale_threads).

    Args:
        thread_signals: ThreadSignals dict.
        target_lower: Lowercased target email address.

    Returns:
        Deduplicated list of ThreadInfo involving the target.
    """
    seen_conv_ids: set[str] = set()
    result: list[ThreadInfo] = []

    all_threads: list[ThreadInfo] = (
        thread_signals.get("awaiting_my_reply", [])
        + thread_signals.get("awaiting_their_reply", [])
        + thread_signals.get("stale_threads", [])
    )

    for thread in all_threads:
        if thread["conversation_id"] in seen_conv_ids:
            continue
        if thread["last_sender_email"].lower() == target_lower:
            result.append(thread)
            seen_conv_ids.add(thread["conversation_id"])

    return result


def _lookup_contact_details(
    request: Callable[..., Any],
    target_email: str,
) -> tuple[str, str, str]:
    """Look up a contact's details via Graph API.

    Queries the user's contacts for a matching email address and
    returns display name, company name, and job title.

    Args:
        request: Injected Graph request callable.
        target_email: Email address to look up.

    Returns:
        Tuple of (displayName, companyName, jobTitle). Empty strings
        for fields that are not found or if the lookup fails.
    """
    try:
        safe_email = target_email.replace("'", "''")
        result = request(
            "GET",
            "/me/contacts",
            params={
                "$filter": f"emailAddresses/any(a:a/address eq '{safe_email}')",
                "$select": "displayName,companyName,jobTitle",
            },
        )
        contacts = (result or {}).get("value", [])
        if contacts:
            contact = contacts[0]
            return (
                contact.get("displayName", ""),
                contact.get("companyName", ""),
                contact.get("jobTitle", ""),
            )
    except Exception:
        logger.warning(
            "Failed to look up contact details for %s",
            target_email,
            exc_info=True,
        )

    return ("", "", "")
