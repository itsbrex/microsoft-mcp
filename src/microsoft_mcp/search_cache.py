"""In-memory TTL cache for degraded search fallback.

Populated by list tools (list_emails, list_events, list_chat_messages,
list_channel_messages). When Graph Search returns 403/404, unified_search
falls back to searching this cache instead of raising.
"""

from __future__ import annotations

import time
import threading
from typing import Any


class SearchCache:
    """Thread-safe in-memory cache of normalized items, keyed by kind."""

    def __init__(self, ttl_seconds: int = 300) -> None:
        self._ttl = ttl_seconds
        self._lock = threading.Lock()
        # kind -> (timestamp, list[item])
        self._store: dict[str, tuple[float, list[dict[str, Any]]]] = {}

    def store(self, kind: str, items: list[dict[str, Any]]) -> None:
        with self._lock:
            self._store[kind] = (time.time(), list(items))

    def search(
        self,
        query: str,
        kinds: list[str] | None = None,
    ) -> list[dict[str, Any]]:
        now = time.time()
        query_lower = query.lower()
        results: list[dict[str, Any]] = []

        with self._lock:
            for kind, (ts, items) in self._store.items():
                if now - ts > self._ttl:
                    continue
                if kinds and kind not in kinds:
                    continue
                for item in items:
                    text = f"{item.get('title', '')} {item.get('snippet', '')}".lower()
                    if query_lower in text:
                        results.append(item)

        return results

    def freshness_info(self) -> dict[str, Any]:
        now = time.time()
        info: dict[str, Any] = {}
        with self._lock:
            for kind, (ts, items) in self._store.items():
                age = now - ts
                info[kind] = {
                    "stored_at": ts,
                    "age_seconds": round(age, 1),
                    "count": len(items),
                    "expired": age > self._ttl,
                }
        return info


# Module-level singleton
_global_cache = SearchCache()


def get_global_cache() -> SearchCache:
    return _global_cache
