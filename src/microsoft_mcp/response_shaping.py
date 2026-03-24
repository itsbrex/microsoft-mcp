from __future__ import annotations

import re
from enum import Enum
from dataclasses import dataclass
from typing import Any
from urllib.parse import quote


class ResponseProfile(str, Enum):
    RAW = "raw"
    DETAIL = "detail"
    SUMMARY = "summary"

    @classmethod
    def default_for_operation(cls, operation: str) -> "ResponseProfile":
        return cls.SUMMARY if operation in {"list", "search"} else cls.DETAIL


@dataclass(frozen=True)
class BudgetHints:
    include_body: bool
    max_items: int

    @classmethod
    def for_operation(cls, tool_name: str) -> "BudgetHints":
        return cls(include_body=False, max_items=25)


# ---------------------------------------------------------------------------
# Global Graph payload cleanup
# ---------------------------------------------------------------------------

ODATA_KEYS = {
    "@odata.context",
    "@odata.etag",
    "@odata.type",
    "@odata.id",
    "@odata.count",
}
NOISE_KEYS = {
    "changeKey",
    "parentFolderId",
    "calendar@odata.associationLink",
    "calendar@odata.navigationLink",
}
_EMPTY = (None, "", [], {})


def cleanup_graph_payload(value: object) -> object:
    if isinstance(value, dict):
        cleaned: dict = {}
        for key, child in value.items():
            if key in ODATA_KEYS or key in NOISE_KEYS:
                continue
            next_value = cleanup_graph_payload(child)
            if next_value in _EMPTY:
                continue
            cleaned[key] = next_value
        return cleaned
    if isinstance(value, list):
        return [
            item
            for item in (cleanup_graph_payload(v) for v in value)
            if item not in _EMPTY
        ]
    return value


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

_EXCHANGE_DN_RE = re.compile(r"^/o=", re.IGNORECASE)
_SAFE_LINK_RE = re.compile(
    r"https?://\S*safelinks\.protection\.outlook\.com/?\?[^\"'\s]*url=([^&\"'\s]+)",
    re.IGNORECASE,
)
_MIMECAST_RE = re.compile(
    r"https?://url\..*?mimecast\.com/[^\"'\s]*",
    re.IGNORECASE,
)
_SECURITY_BANNER_RE = re.compile(
    r"(\[?EXTERNAL[:\]]\s*|CAUTION:.*?(?:\n|$)|This email originated from outside.*?(?:\n|$))",
    re.IGNORECASE,
)


def flatten_email_address(obj: dict[str, Any]) -> str:
    ea = obj.get("emailAddress", obj)
    name = ea.get("name", "")
    address = ea.get("address", "")
    if name and address:
        return f"{name} <{address}>"
    return address or name or ""


def compact_location(loc: dict[str, Any] | None) -> str | None:
    if not loc:
        return None
    return loc.get("displayName") or None


def _extract_teams_join_url(event: dict[str, Any]) -> str | None:
    om = event.get("onlineMeeting")
    if om and om.get("joinUrl"):
        return om["joinUrl"]
    body = event.get("body", {})
    content = body.get("content", "")
    m = re.search(
        r'href="(https://teams\.microsoft\.com/l/meetup-join/[^"]+)"', content
    )
    return m.group(1) if m else None


def extract_teams_meeting_info(event: dict[str, Any]) -> dict[str, Any] | None:
    join_url = _extract_teams_join_url(event)
    if not join_url:
        return None
    info: dict[str, Any] = {"join_url": join_url}
    om = event.get("onlineMeeting", {})
    if om.get("conferenceId"):
        info["meeting_id"] = om["conferenceId"]
    if om.get("tollNumber"):
        info["dial_in"] = om["tollNumber"]
    return info


def _clean_body_text(text: str) -> str:
    text = _SECURITY_BANNER_RE.sub("", text)
    return text.strip()


def _classify_email_address(addr: str) -> tuple[str, bool]:
    if _EXCHANGE_DN_RE.match(addr):
        return addr, False
    return addr, True


# ---------------------------------------------------------------------------
# Email shapers
# ---------------------------------------------------------------------------


def shape_email_summary(raw: dict[str, Any]) -> dict[str, Any]:
    shaped: dict[str, Any] = {"id": raw["id"], "subject": raw.get("subject", "")}

    if "from" in raw:
        shaped["from"] = flatten_email_address(raw["from"])

    if "toRecipients" in raw:
        shaped["to"] = [flatten_email_address(r) for r in raw["toRecipients"]]

    if raw.get("receivedDateTime"):
        shaped["received"] = raw["receivedDateTime"]

    if raw.get("isRead") is not None:
        shaped["is_read"] = raw["isRead"]

    if raw.get("hasAttachments"):
        shaped["has_attachments"] = True

    if raw.get("bodyPreview"):
        shaped["snippet"] = _clean_body_text(raw["bodyPreview"])[:200]

    conv_id = raw.get("conversationId")
    if conv_id:
        shaped["conversation_url"] = (
            f"https://outlook.office.com/mail/deeplink/readconv/{quote(conv_id, safe='')}"
        )

    return shaped


def shape_email_detail(raw: dict[str, Any]) -> dict[str, Any]:
    shaped = shape_email_summary(raw)

    if "ccRecipients" in raw:
        shaped["cc"] = [flatten_email_address(r) for r in raw["ccRecipients"]]

    body = raw.get("body")
    if body:
        content = body.get("content", "")
        shaped["body"] = {
            "contentType": body.get("contentType", "text"),
            "content": _clean_body_text(content),
        }

    if raw.get("webLink"):
        shaped["web_url"] = raw["webLink"]

    return shaped


# ---------------------------------------------------------------------------
# Event shapers
# ---------------------------------------------------------------------------


def shape_event_summary(raw: dict[str, Any]) -> dict[str, Any]:
    shaped: dict[str, Any] = {"id": raw["id"], "subject": raw.get("subject", "")}

    for key in ("start", "end"):
        if key in raw:
            shaped[key] = raw[key]

    shaped["location"] = compact_location(raw.get("location"))

    if "organizer" in raw:
        shaped["organizer"] = flatten_email_address(raw["organizer"])

    if raw.get("isAllDay"):
        shaped["is_all_day"] = True

    if raw.get("seriesMasterId"):
        shaped["series_master_id"] = raw["seriesMasterId"]

    meeting = extract_teams_meeting_info(raw)
    if meeting:
        shaped["meeting"] = meeting

    return shaped


def shape_event_detail(raw: dict[str, Any]) -> dict[str, Any]:
    shaped = shape_event_summary(raw)

    if "attendees" in raw:
        shaped["attendees"] = [
            {
                "name": flatten_email_address(a),
                "status": a.get("status", {}).get("response", "none"),
            }
            for a in raw["attendees"]
        ]

    body = raw.get("body")
    if body:
        shaped["body"] = {
            "contentType": body.get("contentType", "text"),
            "content": body.get("content", ""),
        }

    if raw.get("recurrence"):
        shaped["recurrence"] = raw["recurrence"]

    if raw.get("categories"):
        shaped["categories"] = raw["categories"]

    return shaped


# ---------------------------------------------------------------------------
# Contact shapers
# ---------------------------------------------------------------------------


def _partition_emails(
    raw_addresses: list[dict[str, Any]],
) -> tuple[list[str], list[str]]:
    smtp: list[str] = []
    unresolved: list[str] = []
    for entry in raw_addresses:
        addr = entry.get("address", "")
        if not addr:
            continue
        addr_str, is_smtp = _classify_email_address(addr)
        if is_smtp:
            smtp.append(addr_str)
        else:
            unresolved.append(addr_str)
    return smtp, unresolved


def shape_contact_summary(raw: dict[str, Any]) -> dict[str, Any]:
    shaped: dict[str, Any] = {"id": raw["id"]}

    for key in ("displayName", "jobTitle", "companyName"):
        if raw.get(key):
            shaped[key] = raw[key]

    smtp, unresolved = _partition_emails(raw.get("emailAddresses", []))
    if smtp:
        shaped["email_addresses"] = smtp
    if unresolved:
        shaped["unresolved_addresses"] = unresolved

    for key in ("businessPhones", "mobilePhone"):
        if raw.get(key):
            shaped[key] = raw[key]

    return shaped


def shape_contact_detail(raw: dict[str, Any]) -> dict[str, Any]:
    shaped = shape_contact_summary(raw)

    for key in (
        "givenName",
        "surname",
        "nickname",
        "birthday",
        "department",
        "officeLocation",
        "businessAddress",
        "homeAddress",
        "personalNotes",
    ):
        if raw.get(key):
            shaped[key] = raw[key]

    return shaped


# ---------------------------------------------------------------------------
# Teams message shapers
# ---------------------------------------------------------------------------


def shape_message_summary(raw: dict[str, Any]) -> dict[str, Any]:
    shaped: dict[str, Any] = {"id": raw["id"]}

    sender = raw.get("from", {})
    user = sender.get("user", {})
    shaped["from"] = user.get("displayName", user.get("id", "unknown"))

    if raw.get("createdDateTime"):
        shaped["created"] = raw["createdDateTime"]

    body = raw.get("body", {})
    content = body.get("content", "")
    # Strip HTML tags for snippet
    snippet = re.sub(r"<[^>]+>", "", content).strip()
    shaped["snippet"] = snippet[:200]

    return shaped


def shape_message_detail(raw: dict[str, Any]) -> dict[str, Any]:
    shaped = shape_message_summary(raw)

    body = raw.get("body", {})
    shaped["body"] = {
        "contentType": body.get("contentType", "text"),
        "content": body.get("content", ""),
    }

    if raw.get("attachments"):
        shaped["attachments"] = [
            {
                "id": a.get("id"),
                "name": a.get("name"),
                "contentType": a.get("contentType"),
            }
            for a in raw["attachments"]
        ]

    return shaped
