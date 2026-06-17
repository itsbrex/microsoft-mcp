"""Contact interaction signal collector.

Analyzes recent email activity (sent and received) to build a picture of who
the user interacts with most, and which contacts have pending (unread) emails.

The ``request`` callable is dependency-injected so this module has no direct
import of ``graph``, keeping it pure and unit-testable.
"""

from __future__ import annotations

import logging
from collections.abc import Callable
from datetime import datetime, timedelta
from typing import Any

from microsoft_mcp.intel.types import ContactInteraction, ContactSignals

logger = logging.getLogger(__name__)

# Maximum top contacts returned.
_TOP_CONTACTS_LIMIT = 20


class _ReceivedInfo:
    """Accumulator for emails received from a single sender."""

    __slots__ = ("count", "has_unread", "last_date", "name")

    def __init__(self, name: str, received_at: str, *, is_unread: bool) -> None:
        self.name = name
        self.count = 1
        self.last_date = received_at
        self.has_unread = is_unread

    def update(self, name: str, received_at: str, *, is_unread: bool) -> None:
        self.count += 1
        if name:
            self.name = name
        if received_at > self.last_date:
            self.last_date = received_at
        if is_unread:
            self.has_unread = True


class _SentInfo:
    """Accumulator for emails sent to a single recipient."""

    __slots__ = ("count", "last_date")

    def __init__(self, sent_at: str) -> None:
        self.count = 1
        self.last_date = sent_at

    def update(self, sent_at: str) -> None:
        self.count += 1
        if sent_at > self.last_date:
            self.last_date = sent_at


def _extract_sender(msg: dict[str, Any]) -> tuple[str, str] | None:
    """Extract ``(email, name)`` from a message's ``from`` field.

    Returns ``None`` when the sender cannot be determined.
    """
    from_field = msg.get("from")
    if not from_field:
        return None
    email_address = from_field.get("emailAddress")
    if not email_address:
        return None
    address = email_address.get("address", "")
    if not address:
        return None
    return address.lower(), email_address.get("name", "")


def _extract_recipients(msg: dict[str, Any]) -> list[tuple[str, str]]:
    """Extract ``[(email, name), ...]`` from ``toRecipients``."""
    results: list[tuple[str, str]] = []
    for recipient in msg.get("toRecipients") or []:
        ea = recipient.get("emailAddress")
        if not ea:
            continue
        address = ea.get("address", "")
        if address:
            results.append((address.lower(), ea.get("name", "")))
    return results


def _fetch_messages(
    request: Callable[..., Any],
    path: str,
    filter_field: str,
    since_iso: str,
    select: str,
) -> list[dict[str, Any]]:
    """Fetch messages from a mail folder with a date filter."""
    try:
        data = request(
            "GET",
            path,
            params={
                "$filter": f"{filter_field} ge {since_iso}",
                "$orderby": f"{filter_field} desc",
                "$top": 200,
                "$select": select,
            },
        )
        return data.get("value", [])
    except Exception:
        logger.exception("Failed to fetch messages from %s", path)
        return []


def collect_contact_signals(
    request: Callable[..., Any],
    *,
    now: datetime,
    lookback_days: int = 7,
) -> ContactSignals:
    """Collect contact interaction signals from recent email activity.

    Analyzes both received inbox messages and sent items over the look-back
    window to determine who the user communicates with most and which contacts
    have unread (pending) emails.

    Args:
        request: Injected Graph request callable — signature
            ``request(method, path, *, params=None, json=None) -> dict``.
        now: Current UTC datetime (injected for deterministic testing; never
            call ``datetime.now()`` inside this function).
        lookback_days: Look-back window in days (default 7).

    Returns:
        :class:`ContactSignals` with top contacts, pending contacts, and
        unique sender/recipient counts.
    """
    since = now - timedelta(days=lookback_days)
    since_iso = since.strftime("%Y-%m-%dT%H:%M:%SZ")

    received_select = (
        "id,subject,from,receivedDateTime,isRead,importance,hasAttachments,bodyPreview"
    )
    sent_select = "id,subject,toRecipients,sentDateTime,importance"

    received_msgs = _fetch_messages(
        request, "/me/messages", "receivedDateTime", since_iso, received_select
    )
    sent_msgs = _fetch_messages(
        request,
        "/me/mailFolders/sentitems/messages",
        "sentDateTime",
        since_iso,
        sent_select,
    )

    # Aggregate received emails per sender.
    received: dict[str, _ReceivedInfo] = {}
    for msg in received_msgs:
        sender = _extract_sender(msg)
        if sender is None:
            continue
        email, name = sender
        is_unread = not msg.get("isRead", True)
        if email in received:
            received[email].update(
                name, msg.get("receivedDateTime", ""), is_unread=is_unread
            )
        else:
            received[email] = _ReceivedInfo(
                name, msg.get("receivedDateTime", ""), is_unread=is_unread
            )

    # Aggregate sent emails per recipient.
    sent: dict[str, _SentInfo] = {}
    sent_names: dict[str, str] = {}
    for msg in sent_msgs:
        for email, name in _extract_recipients(msg):
            sent_at = msg.get("sentDateTime", "")
            if email in sent:
                sent[email].update(sent_at)
            else:
                sent[email] = _SentInfo(sent_at)
            if name:
                sent_names[email] = name

    # Merge into ContactInteraction entries.
    interactions: list[ContactInteraction] = []
    for email in set(received) | set(sent):
        recv_info = received.get(email)
        sent_info = sent.get(email)

        received_from = recv_info.count if recv_info else 0
        sent_to = sent_info.count if sent_info else 0

        name = ""
        if recv_info and recv_info.name:
            name = recv_info.name
        elif email in sent_names:
            name = sent_names[email]

        last_dates: list[str] = []
        if recv_info and recv_info.last_date:
            last_dates.append(recv_info.last_date)
        if sent_info and sent_info.last_date:
            last_dates.append(sent_info.last_date)
        last_interaction = max(last_dates) if last_dates else since_iso

        has_pending = recv_info.has_unread if recv_info else False

        interactions.append(
            ContactInteraction(
                email=email,
                name=name,
                total_interactions=sent_to + received_from,
                sent_to=sent_to,
                received_from=received_from,
                last_interaction=last_interaction,
                has_pending_email=has_pending,
            )
        )

    top_contacts = sorted(
        interactions, key=lambda c: c["total_interactions"], reverse=True
    )[:_TOP_CONTACTS_LIMIT]

    pending_contacts = sorted(
        [c for c in interactions if c["has_pending_email"]],
        key=lambda c: c["last_interaction"],
        reverse=True,
    )

    return ContactSignals(
        top_contacts=top_contacts,
        pending_contacts=pending_contacts,
        total_unique_senders=len(received),
        total_unique_recipients=len(sent),
    )
