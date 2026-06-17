"""Thread tracker collector for the intelligence layer.

Tracks conversation threads needing attention by analyzing received and sent
messages to determine reply status and staleness.

The ``request`` callable is dependency-injected so this module has no direct
import of ``graph``, keeping it pure and unit-testable.
"""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime, timedelta
from typing import Any

from microsoft_mcp.intel._utils import parse_graph_datetime as _parse_dt
from microsoft_mcp.intel.types import ThreadInfo, ThreadSignals


def _fetch_received_messages(
    request: Callable[..., Any],
    since_iso: str,
) -> list[dict[str, Any]]:
    """Fetch received messages within the look-back window."""
    data = request(
        "GET",
        "/me/messages",
        params={
            "$filter": f"receivedDateTime ge {since_iso}",
            "$orderby": "receivedDateTime desc",
            "$top": 200,
            "$select": (
                "id,conversationId,subject,from,toRecipients,"
                "receivedDateTime,importance"
            ),
        },
    )
    return data.get("value", [])


def _fetch_sent_messages(
    request: Callable[..., Any],
    since_iso: str,
) -> list[dict[str, Any]]:
    """Fetch sent messages within the look-back window."""
    data = request(
        "GET",
        "/me/mailFolders/sentitems/messages",
        params={
            "$filter": f"sentDateTime ge {since_iso}",
            "$orderby": "sentDateTime desc",
            "$top": 200,
            "$select": (
                "id,conversationId,subject,from,toRecipients,sentDateTime,importance"
            ),
        },
    )
    return data.get("value", [])


def _build_thread_info(
    conversation_id: str,
    last_msg: dict[str, Any],
    last_msg_dt: datetime,
    message_count: int,
    direction: str,
    now: datetime,
) -> ThreadInfo:
    """Build a :class:`ThreadInfo` from the most recent message in a thread."""
    if direction == "inbound":
        addr = (last_msg.get("from") or {}).get("emailAddress") or {}
        sender_name = addr.get("name", "")
        sender_email = addr.get("address", "")
    else:
        recipients = last_msg.get("toRecipients") or []
        if recipients:
            addr = recipients[0].get("emailAddress") or {}
            sender_name = addr.get("name", "")
            sender_email = addr.get("address", "")
        else:
            sender_name = ""
            sender_email = ""

    age_hours = (now - last_msg_dt).total_seconds() / 3600.0
    return ThreadInfo(
        conversation_id=conversation_id,
        subject=last_msg.get("subject", "(no subject)"),
        last_sender_name=sender_name,
        last_sender_email=sender_email,
        last_message_at=last_msg_dt.isoformat(),
        age_hours=round(age_hours, 2),
        message_count=message_count,
        direction=direction,
        importance=last_msg.get("importance", "normal"),
    )


def collect_thread_signals(
    request: Callable[..., Any],
    *,
    now: datetime,
    lookback_hours: int = 48,
    stale_hours: int = 72,
) -> ThreadSignals:
    """Track conversation threads needing attention.

    Fetches recent received and sent messages, groups them by
    ``conversationId``, and determines which threads are awaiting the user's
    reply, which are awaiting the other party's reply, and which outbound
    threads have gone stale.

    Args:
        request: Injected Graph request callable — signature
            ``request(method, path, *, params=None, json=None) -> dict``.
        now: Current UTC datetime (injected for deterministic testing; never
            call ``datetime.now()`` inside this function).
        lookback_hours: Look-back window for active threads (default 48).
        stale_hours: Age threshold in hours for marking outbound threads as
            stale (default 72).

    Returns:
        :class:`ThreadSignals` with ``awaiting_my_reply``,
        ``awaiting_their_reply``, and ``stale_threads`` lists, each sorted by
        ``age_hours`` descending (oldest / most urgent first).
    """
    # Use the longer window so stale detection captures threads older than
    # lookback_hours but within stale_hours.
    lookback = max(lookback_hours, stale_hours)
    since = now - timedelta(hours=lookback)
    since_iso = since.strftime("%Y-%m-%dT%H:%M:%SZ")

    received_msgs = _fetch_received_messages(request, since_iso)
    sent_msgs = _fetch_sent_messages(request, since_iso)

    # Group by conversationId — keep the most recent per direction.
    received_by_conv: dict[str, tuple[dict[str, Any], datetime]] = {}
    received_count: dict[str, int] = {}
    for msg in received_msgs:
        conv_id = msg.get("conversationId")
        if not conv_id:
            continue
        msg_dt = _parse_dt(msg["receivedDateTime"])
        received_count[conv_id] = received_count.get(conv_id, 0) + 1
        existing = received_by_conv.get(conv_id)
        if existing is None or msg_dt > existing[1]:
            received_by_conv[conv_id] = (msg, msg_dt)

    sent_by_conv: dict[str, tuple[dict[str, Any], datetime]] = {}
    sent_count: dict[str, int] = {}
    for msg in sent_msgs:
        conv_id = msg.get("conversationId")
        if not conv_id:
            continue
        msg_dt = _parse_dt(msg["sentDateTime"])
        sent_count[conv_id] = sent_count.get(conv_id, 0) + 1
        existing = sent_by_conv.get(conv_id)
        if existing is None or msg_dt > existing[1]:
            sent_by_conv[conv_id] = (msg, msg_dt)

    all_conv_ids = set(received_by_conv) | set(sent_by_conv)
    awaiting_my_reply: list[ThreadInfo] = []
    awaiting_their_reply: list[ThreadInfo] = []
    stale_threads: list[ThreadInfo] = []

    for conv_id in all_conv_ids:
        recv = received_by_conv.get(conv_id)
        sent = sent_by_conv.get(conv_id)
        total = received_count.get(conv_id, 0) + sent_count.get(conv_id, 0)

        if recv and sent:
            if recv[1] >= sent[1]:
                direction, last_msg, last_dt = "inbound", recv[0], recv[1]
            else:
                direction, last_msg, last_dt = "outbound", sent[0], sent[1]
        elif recv:
            direction, last_msg, last_dt = "inbound", recv[0], recv[1]
        elif sent:
            direction, last_msg, last_dt = "outbound", sent[0], sent[1]
        else:
            continue

        info = _build_thread_info(conv_id, last_msg, last_dt, total, direction, now)

        if direction == "inbound":
            awaiting_my_reply.append(info)
        else:
            awaiting_their_reply.append(info)
            if info["age_hours"] > stale_hours:
                stale_threads.append(info)

    awaiting_my_reply.sort(key=lambda t: t["age_hours"], reverse=True)
    awaiting_their_reply.sort(key=lambda t: t["age_hours"], reverse=True)
    stale_threads.sort(key=lambda t: t["age_hours"], reverse=True)

    return ThreadSignals(
        awaiting_my_reply=awaiting_my_reply,
        awaiting_their_reply=awaiting_their_reply,
        stale_threads=stale_threads,
    )
