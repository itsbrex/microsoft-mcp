"""Relationship analyzer for the intelligence layer.

Analyzes contact interaction patterns to produce relationship insights:
engagement scoring, trend detection, and response-ratio calculation
for each contact in the signals snapshot.
"""

from __future__ import annotations

from datetime import datetime

from microsoft_mcp.intel._utils import parse_graph_datetime as _parse_iso_datetime
from microsoft_mcp.intel.types import (
    ContactInteraction,
    ContactSignals,
    RelationshipInsight,
)

_MAX_SCORE = 100.0

# Engagement score component caps
_INTERACTION_MULTIPLIER = 5
_INTERACTION_CAP = 50
_RECENCY_1_DAY_BONUS = 30
_RECENCY_3_DAY_BONUS = 20
_RECENCY_7_DAY_BONUS = 10
_BIDIRECTIONAL_BONUS = 20

# Trend detection thresholds (days)
_RISING_MAX_DAYS = 2
_RISING_MIN_INTERACTIONS = 5
_COOLING_MIN_DAYS = 5


def _cap(score: float) -> float:
    """Clamp a score to [0.0, 100.0]."""
    return min(max(score, 0.0), _MAX_SCORE)


def _compute_engagement_score(
    contact: ContactInteraction,
    days_since: int,
) -> float:
    """Compute engagement score (0--100) for a contact.

    Components:
        - Base: min(total_interactions * 5, 50)
        - Recency bonus: +30 (1 day), +20 (3 days), +10 (7 days), +0 (older)
        - Bidirectional bonus: +20 if both sent_to > 0 AND received_from > 0

    Args:
        contact: Contact interaction data.
        days_since: Days since last interaction.

    Returns:
        Clamped engagement score between 0.0 and 100.0.
    """
    base = min(
        contact["total_interactions"] * _INTERACTION_MULTIPLIER, _INTERACTION_CAP
    )

    if days_since <= 1:
        recency = _RECENCY_1_DAY_BONUS
    elif days_since <= 3:
        recency = _RECENCY_3_DAY_BONUS
    elif days_since <= 7:
        recency = _RECENCY_7_DAY_BONUS
    else:
        recency = 0

    bidirectional = (
        _BIDIRECTIONAL_BONUS
        if contact["sent_to"] > 0 and contact["received_from"] > 0
        else 0
    )

    return _cap(base + recency + bidirectional)


def _detect_trend(days_since: int, total_interactions: int) -> str:
    """Detect relationship trend from a single snapshot.

    Heuristic based on recency and interaction volume:
        - "rising": recent (<=2 days) and frequent (>=5 interactions)
        - "cooling": no contact in >=5 days
        - "stable": everything else

    Args:
        days_since: Days since last interaction.
        total_interactions: Total interaction count in the window.

    Returns:
        One of "rising", "stable", or "cooling".
    """
    if (
        days_since <= _RISING_MAX_DAYS
        and total_interactions >= _RISING_MIN_INTERACTIONS
    ):
        return "rising"
    if days_since >= _COOLING_MIN_DAYS:
        return "cooling"
    return "stable"


def _analyze_contact(
    contact: ContactInteraction,
    now: datetime,
) -> RelationshipInsight:
    """Produce a RelationshipInsight for a single contact.

    Args:
        contact: Contact interaction data from the collector.
        now: Current UTC time for recency calculations.

    Returns:
        RelationshipInsight with engagement score, trend, and response ratio.
    """
    last_dt = _parse_iso_datetime(contact["last_interaction"])
    days_since = (now - last_dt).days

    engagement = _compute_engagement_score(contact, days_since)
    trend = _detect_trend(days_since, contact["total_interactions"])
    response_ratio = contact["received_from"] / max(contact["sent_to"], 1)

    insight = RelationshipInsight(
        email=contact["email"],
        name=contact["name"],
        engagement_score=round(engagement, 1),
        trend=trend,
        last_interaction=contact["last_interaction"],
        days_since_contact=days_since,
        sent_to=contact["sent_to"],
        received_from=contact["received_from"],
        response_ratio=round(response_ratio, 2),
    )

    if "company" in contact:
        insight["company"] = contact["company"]

    return insight


def analyze_relationships(
    contacts: ContactSignals,
    *,
    now: datetime,
) -> list[RelationshipInsight]:
    """Analyze contact relationships for insights.

    Processes each contact in the top_contacts list, computing engagement
    scores, trend detection, and response ratios. Returns results sorted
    by engagement score descending.

    Args:
        contacts: Collected contact interaction signals.
        now: Current UTC datetime (injected for deterministic testing).

    Returns:
        List of RelationshipInsight sorted by engagement_score descending.
    """
    insights = [_analyze_contact(contact, now) for contact in contacts["top_contacts"]]
    insights.sort(key=lambda i: i["engagement_score"], reverse=True)
    return insights
