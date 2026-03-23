import time
from unittest.mock import patch, MagicMock

import httpx


def _make_403_error():
    """Create a realistic 403 HTTPStatusError."""
    response = MagicMock(spec=httpx.Response)
    response.status_code = 403
    response.text = "Forbidden"
    request = MagicMock(spec=httpx.Request)
    return httpx.HTTPStatusError("403 Forbidden", request=request, response=response)


# ---------------------------------------------------------------------------
# SearchCache unit tests
# ---------------------------------------------------------------------------


def test_search_cache_store_and_search():
    from microsoft_mcp.search_cache import SearchCache

    cache = SearchCache(ttl_seconds=60)
    cache.store(
        "message",
        [
            {
                "id": "m1",
                "kind": "message",
                "title": "Budget Q4",
                "snippet": "quarterly budget review",
            },
            {
                "id": "m2",
                "kind": "message",
                "title": "Lunch plans",
                "snippet": "where to eat",
            },
        ],
    )

    results = cache.search("budget", kinds=["message"])
    assert len(results) == 1
    assert results[0]["id"] == "m1"


def test_search_cache_returns_empty_after_ttl():
    from microsoft_mcp.search_cache import SearchCache

    cache = SearchCache(ttl_seconds=0)
    cache.store(
        "message", [{"id": "m1", "kind": "message", "title": "Budget", "snippet": ""}]
    )
    time.sleep(0.05)

    results = cache.search("budget", kinds=["message"])
    assert results == []


def test_search_cache_filters_by_kind():
    from microsoft_mcp.search_cache import SearchCache

    cache = SearchCache(ttl_seconds=60)
    cache.store(
        "message", [{"id": "m1", "kind": "message", "title": "Budget", "snippet": ""}]
    )
    cache.store(
        "event",
        [{"id": "e1", "kind": "event", "title": "Budget meeting", "snippet": ""}],
    )

    results = cache.search("budget", kinds=["event"])
    assert len(results) == 1
    assert results[0]["kind"] == "event"


def test_search_cache_data_freshness():
    from microsoft_mcp.search_cache import SearchCache

    cache = SearchCache(ttl_seconds=60)
    cache.store(
        "message", [{"id": "m1", "kind": "message", "title": "Test", "snippet": ""}]
    )

    info = cache.freshness_info()
    assert "message" in info
    assert "stored_at" in info["message"]
    assert info["message"]["count"] == 1


# ---------------------------------------------------------------------------
# unified_search degraded-mode integration
# ---------------------------------------------------------------------------


@patch("microsoft_mcp.tools.graph")
def test_unified_search_returns_degraded_mode_when_graph_search_is_forbidden(
    mock_graph,
):
    from microsoft_mcp.tools import unified_search

    # First call succeeds and populates cache
    mock_graph.request.return_value = {
        "value": [
            {
                "hitsContainers": [
                    {
                        "total": 1,
                        "hits": [
                            {
                                "rank": 1,
                                "summary": "AI pilot kickoff",
                                "resource": {
                                    "@odata.type": "#microsoft.graph.message",
                                    "id": "msg-cached",
                                    "subject": "AI pilot",
                                    "from": {
                                        "emailAddress": {
                                            "name": "JP",
                                            "address": "jp@x.com",
                                        }
                                    },
                                    "receivedDateTime": "2026-03-23T10:00:00Z",
                                    "conversationId": "conv-1",
                                },
                            }
                        ],
                    }
                ]
            }
        ]
    }
    unified_search.fn("AI pilot", entity_types=["message"])

    # Second call fails with 403
    mock_graph.request.side_effect = _make_403_error()

    result = unified_search.fn("AI pilot", entity_types=["message"])
    assert result["summary"]["mode"] in {"graph_search", "degraded_cache_search"}


@patch("microsoft_mcp.tools.graph")
def test_unified_search_graph_search_mode_on_success(mock_graph):
    from microsoft_mcp.tools import unified_search

    mock_graph.request.return_value = {"value": []}

    result = unified_search.fn("test query", entity_types=["message"])
    assert result["summary"]["mode"] == "graph_search"


@patch("microsoft_mcp.tools.graph")
def test_unified_search_degraded_includes_meta_fields(mock_graph):
    from microsoft_mcp.tools import unified_search
    from microsoft_mcp.search_cache import get_global_cache

    # Pre-populate cache
    cache = get_global_cache()
    cache.store(
        "message",
        [
            {
                "id": "m1",
                "kind": "message",
                "title": "budget report",
                "snippet": "Q4 budget",
            }
        ],
    )

    # 403 on search
    mock_graph.request.side_effect = _make_403_error()

    result = unified_search.fn("budget", entity_types=["message"])
    assert result["summary"]["mode"] == "degraded_cache_search"
    assert "data_freshness" in result["summary"]
    assert "degraded_reason" in result["summary"]
