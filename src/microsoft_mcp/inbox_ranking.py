"""Ranking heuristics for inbox items."""

from __future__ import annotations

from .inbox_models import InboxItem


def _compute_score(item: InboxItem) -> float:
    score = 0.0

    # Unread items get a base boost
    if item.unread:
        score += 10.0

    # Mentions are high-signal
    if item.mentioned:
        score += 15.0

    # Flagged / important
    if item.flagged:
        score += 8.0

    # Meeting start proximity: closer = higher priority
    if item.starts_in_minutes is not None:
        if item.starts_in_minutes <= 15:
            score += 25.0
        elif item.starts_in_minutes <= 60:
            score += 15.0
        elif item.starts_in_minutes <= 120:
            score += 5.0

    # Newsletter suppression
    if item.is_newsletter:
        score -= 20.0

    return score


def rank_items(items: list[InboxItem]) -> list[InboxItem]:
    """Rank inbox items by priority signals and return sorted list."""
    for item in items:
        item.score = _compute_score(item)

    return sorted(items, key=lambda i: i.score, reverse=True)
