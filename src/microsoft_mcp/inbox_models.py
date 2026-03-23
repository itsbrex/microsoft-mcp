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
        return d
