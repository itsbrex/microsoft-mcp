"""Bounce/NDR (non-delivery report) classifier, DSN parser, and folder scanner.

Pure module — no global imports of graph or auth; all Graph interaction uses
an injected ``request`` callable so the module remains unit-testable.

Pattern catalogs consolidated from outlook-creds api/bounces.py so detection
logic does not drift across scripts.
"""

from __future__ import annotations

import csv
import pathlib
import re
from collections.abc import Iterator
from html import unescape
from typing import Any

# ============================================================================
# Pattern catalogs
# ============================================================================

# 23 subject keywords (used to surface bounce candidates)
SUBJECT_KEYWORDS: list[str] = [
    # Core delivery failures
    "Undeliverable",
    "Delivery Failure",
    "Delivery Warning",
    "Returned mail",
    "Mail Delivery Failed",
    "Mail Delivery Subsystem",
    "Delivery Status Notification",
    # Failure indicators
    "Failed",
    "Rejected",
    "Bounced",
    "Not delivered",
    "Could not be delivered",
    # System messages
    "System Administrator",
    "Automatic reply",
    "Out of Office",
    # Specific error types
    "Invalid recipient",
    "User unknown",
    "Mailbox full",
    "Mailbox unavailable",
    "Message too large",
    "Relay access denied",
    "Connection timeout",
    "Connection refused",
]

# Sender substrings (case-insensitive)
SENDER_PATTERNS: list[str] = [
    "postmaster",
    "mailer-daemon",
    "no-reply",
    "noreply",
    "do-not-reply",
    "administrator",
    "system",
    "mail delivery",
]

# Body content substrings (case-insensitive)
BODY_PATTERNS: list[str] = [
    # MX/DNS issues
    "MX record",
    "MX records or is invalid",
    "Domain has no MX records",
    "DNS lookup failed",
    "No route to host",
    # Recipient issues
    "recipient email address is possibly incorrect",
    "Invalid recipient",
    "User unknown",
    "Address not found",
    "No such user",
    "Unknown user",
    # Delivery system messages
    "Email Delivery Failure",
    "Message could not be delivered",
    "Delivery has failed",
    "permanent error",
    "permanent failure",
    # Connection/network issues
    "Communications error",
    "Connection timed out",
    "Read timed out",
    "Network unreachable",
    "Connection refused",
    # Mailbox issues
    "Mailbox full",
    "Mailbox unavailable",
    "Quota exceeded",
    "Insufficient storage",
    # Rejection/blocking
    "Relay access denied",
    "Sender address rejected",
    "Recipient address rejected",
    "Message rejected",
    "Blocked by recipient",
    "Spam detected",
    "Policy violation",
]

# Reason classifiers (regex, priority order — first match wins)
BOUNCE_REASONS: list[tuple[str, str]] = [
    # SMTP status codes (from DSN Diagnostic-Code)
    (r"550 5\.1\.10", "Recipient Not Found (550 5.1.10)"),
    (r"550 5\.1\.1\b", "Invalid Recipient"),
    (r"550 5\.4\.1", "No Answer from Host (550 5.4.1)"),
    (r"550 5\.7\.1", "Delivery Not Authorized (550 5.7.1)"),
    (r"550 5\.2\.1", "Mailbox Disabled (550 5.2.1)"),
    (r"552 5\.2\.2", "Mailbox Full (552 5.2.2)"),
    (r"5\.4\.14", "Mail Loop Detected (5.4.14)"),
    (r"[Hh]op count exceeded", "Mail Loop Detected"),
    (r"RESOLVER\.ADR\.RecipientNotFound", "Recipient Not Found"),
    (r"RESOLVER\.ADR\.BadPrimary", "Bad Primary Address"),
    # Connection/network
    (r"Communications error", "Communications Error"),
    (r"Read timed out", "Read Timeout"),
    (r"Connection timed out", "Connection Timeout"),
    (r"Connection refused", "Connection Refused"),
    (r"Network unreachable", "Network Unreachable"),
    # DNS/MX
    (r"MX records? (?:or )?is invalid", "Invalid MX Records"),
    (r"Domain has no MX records", "No MX Records"),
    (r"DNS lookup failed", "DNS Lookup Failed"),
    (r"Temporary error looking up MX", "Temporary MX Lookup Error"),
    # Recipient
    (r"Recipient email address is possibly incorrect", "Invalid Recipient Address"),
    (r"User unknown", "User Unknown"),
    (r"No such user", "No Such User"),
    (r"Address not found", "Address Not Found"),
    (r"not found by SMTP address lookup", "SMTP Address Not Found"),
    # Mailbox
    (r"Mailbox unavailable", "Mailbox Unavailable"),
    (r"Mailbox full", "Mailbox Full"),
    (r"Mailbox disabled", "Mailbox Disabled"),
    (r"Quota exceeded", "Quota Exceeded"),
    (r"Insufficient storage", "Insufficient Storage"),
    # Size/content
    (r"Message too large", "Message Too Large"),
    # Relay/policy
    (r"Relay access denied", "Relay Access Denied"),
    (r"Sender address rejected", "Sender Rejected"),
    (r"Recipient address rejected", "Recipient Rejected"),
    (r"Blocked by recipient", "Blocked by Recipient"),
    (r"Spam filter", "Spam Filter"),
    (r"Content filter", "Content Filter"),
    (r"Policy rejection", "Policy Rejection"),
    (r"Delivery not authorized", "Delivery Not Authorized"),
]

# Subjects considered strong bounce indicators regardless of sender
STRONG_SUBJECT_INDICATORS: tuple[str, ...] = (
    "[Postmaster]",
    "Undeliverable:",
    "Undeliverable mail:",
    "Undelivered Mail",
    "Delivery Failure",
    "Delivery Warning",
    "Delivery Status Notification",
    "couldn't be delivered",
    "Returned mail:",
)

# Subjects that should NOT count as bounces even if sender matches
EXCLUDED_SUBJECT_PREFIXES: tuple[str, ...] = (
    "Automatic reply:",
    "Messages on hold",
)

# Fields the bounce scanner needs from a Graph message
BOUNCE_MESSAGE_FIELDS = (
    "id,subject,from,sender,receivedDateTime,body,bodyPreview,hasAttachments"
)

# CSV output columns (order determines column order in CSV)
_CSV_FIELDNAMES = [
    "first_name",
    "last_name",
    "email",
    "reason",
    "date",
    "iso_date",
    "subject",
    "sender",
    "body",
    "message_id",
    "has_attachments",
]

# Graph base URLs stripped from @odata.nextLink before passing to request()
_GRAPH_BASES = (
    "https://graph.microsoft.com/v1.0",
    "https://graph.microsoft.com/beta",
)


def _strip_graph_base(url: str) -> str:
    """Turn an absolute Graph @odata.nextLink into a path that request() accepts."""
    for base in _GRAPH_BASES:
        if url.startswith(base):
            return url[len(base) :]
    return url


# ============================================================================
# Pure helpers
# ============================================================================

_EMAIL_RE = re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")


def extract_email_from_text(text: str | None) -> str | None:
    """Pull the first email address out of free text. Returns None if none found."""
    if not text:
        return None
    match = _EMAIL_RE.search(text)
    return match.group(0) if match else None


def parse_name_from_email(email: str | None) -> tuple[str, str]:
    """Split first/last name out of the local part. Best-effort.

    Handles ``first.last``, ``first_last``, ``first-last``, and CamelCase.
    """
    if not email or "@" not in email:
        return ("", "")

    local = email.split("@", 1)[0]
    local = re.sub(r"\d+", "", local)
    parts = re.split(r"[._-]", local)

    if len(parts) >= 2:
        return (parts[0].capitalize(), parts[-1].capitalize())
    if len(parts) == 1 and parts[0]:
        camel = re.findall(r"[A-Z]?[a-z]+|[A-Z]+(?=[A-Z][a-z]|\b)", parts[0])
        if len(camel) >= 2:
            return (camel[0].capitalize(), camel[-1].capitalize())
        return (parts[0].capitalize(), "")
    return ("", "")


def parse_dsn_content(text: str) -> dict[str, str | None]:
    """Extract recipient/status/action/diagnostic out of a DSN or EML payload.

    Looks at X-Failed-Recipients, Final-Recipient, Original-Recipient, and
    To: in priority order.

    Returns a dict with keys: ``final_recipient``, ``action``, ``status``,
    ``diagnostic_code``, ``display_name``.
    """
    result: dict[str, str | None] = {
        "final_recipient": None,
        "action": None,
        "status": None,
        "diagnostic_code": None,
        "display_name": None,
    }

    m = re.search(r"X-Failed-Recipients:\s*(.+?)[\r\n]", text, re.IGNORECASE)
    if m:
        result["final_recipient"] = m.group(1).strip()

    if not result["final_recipient"]:
        m = re.search(
            r"Final-Recipient:\s*(?:rfc822;|rfc/822;|RFC822;\s*)?\s*<?([^>\s\r\n]+@[^>\s\r\n]+)>?",
            text,
            re.IGNORECASE,
        )
        if m:
            result["final_recipient"] = m.group(1).strip("<>")

    if not result["final_recipient"]:
        m = re.search(
            r"Original-Recipient:\s*(?:rfc822;|rfc/822;|RFC822;\s*)?\s*<?([^>\s\r\n]+@[^>\s\r\n]+)>?",
            text,
            re.IGNORECASE,
        )
        if m:
            result["final_recipient"] = m.group(1).strip("<>")

    if not result["final_recipient"]:
        to_matches = re.findall(
            r"^To:\s*(?:[^<\r\n]*<)?([^>\s\r\n]+@[^>\s\r\n]+)>?",
            text,
            re.MULTILINE | re.IGNORECASE,
        )
        for raw in to_matches:
            addr = raw.strip("<>")
            result["final_recipient"] = addr
            break

    m = re.search(r"Action:\s*(.+?)[\r\n]", text, re.IGNORECASE)
    if m:
        result["action"] = m.group(1).strip()

    m = re.search(r"Status:\s*(\d+\.\d+\.\d+)", text, re.IGNORECASE)
    if m:
        result["status"] = m.group(1).strip()

    m = re.search(
        r"Diagnostic-Code:\s*(.+?)(?=\r?\n[A-Z]|\r?\n\r?\n|$)",
        text,
        re.IGNORECASE | re.DOTALL,
    )
    if m:
        result["diagnostic_code"] = " ".join(m.group(1).split())

    m = re.search(r"X-Display-Name:\s*(.+?)[\r\n]", text, re.IGNORECASE)
    if m:
        result["display_name"] = m.group(1).strip()

    return result


def determine_bounce_reason(subject: str, body: str) -> str:
    """Classify bounce reason by scanning subject + body against BOUNCE_REASONS.

    Note: parameter order is ``(subject, body)`` — the brief order, which is
    the reverse of the reference implementation.
    """
    combined = f"{subject}\n{body}"
    for pattern, reason in BOUNCE_REASONS:
        if re.search(pattern, combined, re.IGNORECASE):
            return reason

    subject_lower = subject.lower()
    if "failure" in subject_lower:
        return "Delivery Failure (Unspecified)"
    if "warning" in subject_lower:
        return "Delivery Warning (Unspecified)"
    if "undeliverable" in subject_lower:
        return "Undeliverable"
    return "Unknown"


def is_bounce_message(
    subject: str,
    sender_email: str,
    body: str | None = None,
    *,
    use_body: bool = True,
) -> bool:
    """Decide whether a message is a bounce/NDR.

    Logic:

    1. Exclude auto-replies and quarantine "Messages on hold" notifications.
    2. Match if sender contains postmaster/mailer-daemon/DAEMON and subject
       contains a delivery-failure phrase.
    3. Match if subject alone contains a strong indicator (catches forwarded
       bounces like ``FW: [Postmaster]``).
    4. Optionally match on body content patterns (only when body is provided).
    """
    if not subject:
        subject = ""
    if not sender_email:
        sender_email = ""

    for prefix in EXCLUDED_SUBJECT_PREFIXES:
        if prefix in subject:
            return False

    sender_lower = sender_email.lower()
    sender_is_postmaster = any(
        s in sender_lower for s in ("postmaster", "mailer-daemon", "daemon")
    )

    if sender_is_postmaster:
        delivery_phrases = (
            "Delivery Failure",
            "Delivery Warning",
            "Undeliverable",
            "Returned mail",
            "couldn't be delivered",
            "could not be delivered",
            "Delivery Status Notification",
            "Undelivered Mail",
            "Mail Delivery",
            "Delivery has failed",
        )
        if any(p in subject for p in delivery_phrases):
            return True

    if any(p in subject for p in STRONG_SUBJECT_INDICATORS):
        return True

    if use_body and body:
        body_lower = body.lower()
        if any(p.lower() in body_lower for p in BODY_PATTERNS):
            return True

    return False


def _message_body_text(message: dict[str, Any]) -> str:
    """Extract plain text body content from a Graph-style message dict.

    Strips HTML tags if contentType is HTML.
    """
    body = message.get("body") or {}
    content = body.get("content") or message.get("bodyPreview") or ""
    if (body.get("contentType") or "").lower() == "html":
        content = unescape(re.sub(r"<[^>]+>", " ", content))
        content = re.sub(r"\s+", " ", content).strip()
    return content


def _sender_email_from_message(message: dict[str, Any]) -> str:
    """Pull the sender's email address from a Graph-style message dict."""
    for key in ("from", "sender"):
        block = message.get(key) or {}
        addr = (block.get("emailAddress") or {}).get("address")
        if addr:
            return addr
    return ""


# Public aliases — the brief and CLI reference these without the underscore prefix
def message_body_text(message: dict[str, Any]) -> str:
    """Public alias for ``_message_body_text``.

    Extract plain text body content from a Graph-style message dict.
    Strips HTML tags if contentType is HTML.
    """
    return _message_body_text(message)


def sender_email_from_message(message: dict[str, Any]) -> str:
    """Public alias for ``_sender_email_from_message``.

    Pull the sender's email address from a Graph-style message dict.
    """
    return _sender_email_from_message(message)


def classify_bounce_message(msg: dict[str, Any]) -> dict[str, Any] | None:
    """Build a structured bounce record from a Graph-style message dict.

    Returns ``None`` if the message is not a bounce. Does not call Graph.

    Returned dict keys: ``first_name``, ``last_name``, ``email``, ``reason``,
    ``date``, ``iso_date``, ``subject``, ``sender``, ``body``,
    ``message_id``, ``has_attachments``.
    """
    subject = msg.get("subject") or ""
    sender = _sender_email_from_message(msg)
    body = _message_body_text(msg)

    if not is_bounce_message(subject, sender, body):
        return None

    recipient = extract_email_from_text(body) or extract_email_from_text(subject)
    if recipient:
        first_name, last_name = parse_name_from_email(recipient)
    else:
        first_name, last_name = ("Unknown", "Recipient")
        recipient = "(not extracted)"

    reason = determine_bounce_reason(subject, body)
    received = msg.get("receivedDateTime") or ""

    return {
        "first_name": first_name,
        "last_name": last_name,
        "email": recipient,
        "reason": reason,
        "date": received,
        "iso_date": received,
        "subject": subject,
        "sender": sender,
        "body": body.replace("\n", " ").replace("\r", " ")[:2000],
        "message_id": msg.get("id"),
        "has_attachments": bool(msg.get("hasAttachments")),
    }


# ============================================================================
# Graph-backed folder scanner (injected request, no global graph import)
# ============================================================================


def iter_folder_messages(
    request: Any,
    folder_id: str,
    *,
    limit: int | None = None,
    page_size: int = 50,
    select: str = BOUNCE_MESSAGE_FIELDS,
    order_by: str = "receivedDateTime desc",
) -> Iterator[dict[str, Any]]:
    """Yield messages from a Graph mail folder, paginating via @odata.nextLink.

    ``request`` is the injected Graph request callable
    (``graph.request`` in production, a fake in tests).

    Stops once ``limit`` messages have been yielded (``None`` = no cap).
    ``@odata.nextLink`` is an absolute URL; we strip the Graph base prefix
    before passing it to ``request``, which expects a path.
    """
    path = f"/me/mailFolders/{folder_id}/messages"
    params: dict[str, Any] | None = {
        "$top": min(page_size, limit) if limit is not None else page_size,
        "$select": select,
        "$orderby": order_by,
    }
    yielded = 0

    while True:
        data = (
            request("GET", path, params=params)
            if params is not None
            else request("GET", path)
        )
        if not data:
            break
        for msg in data.get("value", []):
            yield msg
            yielded += 1
            if limit is not None and yielded >= limit:
                return
        next_link = data.get("@odata.nextLink")
        if not next_link:
            break
        path = _strip_graph_base(next_link)
        params = None  # nextLink already encodes all query params


def scan_folder(
    request: Any,
    folder_id: str,
    *,
    limit: int | None = None,
) -> list[dict[str, Any]]:
    """Scan a mail folder and return classified bounce records.

    Iterates up to ``limit`` messages (``None`` = no cap) via
    ``iter_folder_messages``, checks each with ``is_bounce_message``, and
    classifies matching messages with ``classify_bounce_message``.

    Returns a list of bounce dicts (``None`` results from
    ``classify_bounce_message`` are silently skipped).
    """
    results: list[dict[str, Any]] = []
    for msg in iter_folder_messages(request, folder_id, limit=limit):
        subject = msg.get("subject") or ""
        sender = sender_email_from_message(msg)
        body = message_body_text(msg)
        if is_bounce_message(subject, sender, body):
            record = classify_bounce_message(msg)
            if record is not None:
                results.append(record)
    return results


def write_csv(rows: list[dict[str, Any]], path: str | pathlib.Path) -> None:
    """Write bounce records to a CSV file.

    Columns follow ``_CSV_FIELDNAMES`` order.  If ``rows`` is empty, a
    header-only CSV is written.  The file is always UTF-8 with proper
    newline handling.
    """
    out_path = pathlib.Path(path)
    with out_path.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=_CSV_FIELDNAMES, extrasaction="ignore")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)
