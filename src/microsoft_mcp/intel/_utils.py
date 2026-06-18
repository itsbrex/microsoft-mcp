"""Shared utilities for the intelligence layer."""

from __future__ import annotations

from collections.abc import Callable
from datetime import datetime
from typing import Any

_GRAPH_BASES = (
    "https://graph.microsoft.com/v1.0",
    "https://graph.microsoft.com/beta",
)


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


def _strip_graph_base(url: str) -> str:
    """Turn an absolute Graph @odata.nextLink into a path graph.request accepts."""
    for base in _GRAPH_BASES:
        if url.startswith(base):
            return url[len(base) :]
    return url


def paginate(
    request: Callable[..., Any],
    path: str,
    params: dict[str, Any] | None = None,
    *,
    limit: int | None = None,
) -> list[dict[str, Any]]:
    """Collect all ``value`` items across @odata.nextLink pages via the injected request.

    Mirrors graph.request_paginated but uses the dependency-injected ``request``
    so the intel package stays pure and unit-testable.

    Args:
        request: Injected Graph request callable.
        path: Initial path to request (e.g. ``"/me/messages"``).
        params: Optional query parameters for the first request.
        limit: Optional maximum number of items to return.

    Returns:
        Flat list of all ``value`` items from all pages.
    """
    items: list[dict[str, Any]] = []
    next_link: str | None = None
    while True:
        if next_link:
            result = request("GET", _strip_graph_base(next_link))
        else:
            result = request("GET", path, params=params)
        if not result:
            break
        for item in result.get("value", []):
            items.append(item)
            if limit is not None and len(items) >= limit:
                return items
        next_link = result.get("@odata.nextLink")
        if not next_link:
            break
    return items
