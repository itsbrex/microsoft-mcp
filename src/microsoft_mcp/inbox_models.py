"""Normalized inbox item model shared across email, calendar, and Teams."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class InboxItem:
    id: str = ""
    kind: str = ""  # email, event, chatMessage, channelMessage
    source_tool: str = ""
    title: str = ""
    snippet: str = ""
    participants: list[str] = field(default_factory=list)
    cc: list[str] = field(default_factory=list)
    bcc: list[str] = field(default_factory=list)
    when: str | None = None
    state: str = ""  # e.g. unread, tentative, accepted
    score: float = 0.0
    reason: str = ""
    action_hints: list[str] = field(default_factory=list)
    web_url: str = ""

    # Ranking signals
    unread: bool = False
    mentioned: bool = False
    flagged: bool = False
    is_newsletter: bool = False
    is_bounce: bool = False
    has_attachments: bool = False
    direct_to: bool = False  # active user in toRecipients (not just cc/bcc)
    on_cc: bool = False
    on_bcc: bool = False
    starts_in_minutes: float | None = None

    def to_dict(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "id": self.id,
            "kind": self.kind,
            "title": self.title,
            "score": self.score,
        }
        if self.snippet:
            d["snippet"] = self.snippet
        if self.participants:
            d["participants"] = self.participants
            d["sender"] = self.participants[0]
        if self.cc:
            d["cc"] = self.cc
        if self.bcc:
            d["bcc"] = self.bcc
        if self.when:
            d["when"] = self.when
        if self.state:
            d["state"] = self.state
        if self.reason:
            d["reason"] = self.reason
        if self.action_hints:
            d["action_hints"] = self.action_hints
        if self.web_url:
            d["web_url"] = self.web_url
        if self.has_attachments:
            d["has_attachments"] = True
        if self.is_bounce:
            d["is_bounce"] = True
        if self.direct_to:
            d["direct_to"] = True
        if self.on_cc:
            d["on_cc"] = True
        if self.on_bcc:
            d["on_bcc"] = True
        if self.unread:
            d["unread"] = True
        if self.starts_in_minutes is not None:
            d["starts_in_minutes"] = self.starts_in_minutes
        return d
