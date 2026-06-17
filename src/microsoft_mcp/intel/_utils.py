"""Shared utilities for the intelligence layer."""

from __future__ import annotations

from datetime import datetime


def parse_graph_datetime(dt_string: str) -> datetime:
    """Parse an ISO datetime string from the Microsoft Graph API.

    Handles both ``Z`` suffix and ``+00:00`` offset formats.  Python 3.10's
    ``fromisoformat`` does not accept a bare ``Z``, so we normalise first.

    Args:
        dt_string: ISO datetime string (e.g. ``"2026-02-23T10:00:00Z"``).

    Returns:
        Parsed ``datetime`` (timezone-aware if the input includes an offset,
        naive otherwise — which is the case when the Graph API ``Prefer:
        outlook.timezone`` header strips the offset).
    """
    return datetime.fromisoformat(dt_string.replace("Z", "+00:00"))
