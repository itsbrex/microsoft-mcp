from microsoft_mcp import tools as tools_mod
from microsoft_mcp.search_cache import get_global_cache


def _reset_cache():
    cache = get_global_cache()
    with cache._lock:
        cache._store.clear()


def test_list_emails_populates_cache(monkeypatch):
    _reset_cache()

    def fake_paginated(path, params=None, limit=None, auth=None):
        yield {
            "id": "m-1",
            "subject": "Project alpha kickoff",
            "bodyPreview": "Details inside",
            "from": {"emailAddress": {"address": "a@example.com", "name": "A"}},
            "receivedDateTime": "2026-04-23T00:00:00Z",
            "isRead": False,
            "conversationId": "c-1",
        }

    monkeypatch.setattr("microsoft_mcp.graph.request_paginated", fake_paginated)

    tools_mod.list_emails.fn(folder="inbox", limit=5)

    hits = get_global_cache().search("alpha", kinds=["message"])
    assert any(h.get("id") == "m-1" for h in hits)


def test_cache_hit_survives_between_calls(monkeypatch):
    _reset_cache()

    def fake_paginated(path, params=None, limit=None, auth=None):
        yield {"id": "m-2", "subject": "Budget review", "bodyPreview": "Q3 budget"}

    monkeypatch.setattr("microsoft_mcp.graph.request_paginated", fake_paginated)
    tools_mod.list_emails.fn(folder="inbox", limit=5)

    hits = get_global_cache().search("budget", kinds=["message"])
    assert len(hits) >= 1
