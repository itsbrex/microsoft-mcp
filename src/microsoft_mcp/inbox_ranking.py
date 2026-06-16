"""Ranking heuristics for inbox items."""

from __future__ import annotations

from .inbox_models import InboxItem


def _event_proximity_score(minutes: float) -> float:
    """Score event by proximity. Smooth decay across a 7-day window.

    - <= 15 min  -> 25
    - <= 60 min  -> 15
    - <= 120 min -> 8
    - <= 24 h    -> linear 7 -> 3
    - <= 7 days  -> linear 3 -> 0
    - > 7 days   -> 0
    """
    if minutes <= 15:
        return 25.0
    if minutes <= 60:
        return 15.0
    if minutes <= 120:
        return 8.0
    if minutes <= 1440:  # 24h
        span = 1440 - 120
        frac = (minutes - 120) / span
        return 7.0 - (7.0 - 3.0) * frac
    if minutes <= 7 * 1440:  # 7d
        span = 7 * 1440 - 1440
        frac = (minutes - 1440) / span
        return 3.0 - 3.0 * frac
    return 0.0


def _compute_score_and_reason(item: InboxItem) -> tuple[float, str]:
    score = 0.0
    reasons: list[str] = []

    if item.unread:
        score += 10.0
        reasons.append("unread")

    if item.mentioned:
        score += 15.0
        reasons.append("mentioned")

    if item.flagged:
        score += 8.0
        reasons.append("flagged")

    if item.direct_to:
        score += 5.0
        reasons.append("direct-to")
    elif item.on_cc:
        score -= 5.0
        reasons.append("cc only")
    elif item.on_bcc:
        score -= 8.0
        reasons.append("bcc only")

    if item.has_attachments:
        score += 2.0
        reasons.append("has attachments")

    if item.starts_in_minutes is not None:
        prox = _event_proximity_score(item.starts_in_minutes)
        if prox > 0:
            score += prox
            if item.starts_in_minutes <= 15:
                reasons.append(f"starts in {int(item.starts_in_minutes)}m")
            elif item.starts_in_minutes <= 60:
                reasons.append(f"starts in {int(item.starts_in_minutes)}m")
            elif item.starts_in_minutes <= 1440:
                reasons.append(f"in {int(item.starts_in_minutes / 60)}h")
            else:
                reasons.append(f"in {int(item.starts_in_minutes / 1440)}d")

    if item.is_newsletter:
        score -= 20.0
        reasons.append("newsletter")

    if item.is_bounce:
        # Strong suppression — DSN/NDR messages are pure noise and should sink
        # below newsletters. Net of an unread direct-to bounce with attachments
        # is roughly 17 - 30 = -13, well below all human mail.
        score -= 30.0
        reasons.append("bounce")

    reason = ", ".join(reasons)
    return score, reason


def _compute_score(item: InboxItem) -> float:
    """Return the score only. Kept for callers that don't need the reason."""
    score, _ = _compute_score_and_reason(item)
    return score


def rank_items(items: list[InboxItem]) -> list[InboxItem]:
    """Rank inbox items by priority signals and return sorted list."""
    for item in items:
        item.score, item.reason = _compute_score_and_reason(item)

    return sorted(items, key=lambda i: i.score, reverse=True)
