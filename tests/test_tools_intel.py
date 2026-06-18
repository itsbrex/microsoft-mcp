"""Tests for the four intel MCP tools in tools.py.

Each tool delegates to the intel engine with dependency-injected ``graph.request``
and a clock-injected ``now``.  We mock the engine functions so tests are
deterministic and require no live Graph credentials.

Patching strategy: tools.py is imported as ``src.microsoft_mcp.tools``, so
``_intel_engine`` is the ``src.microsoft_mcp.intel.engine`` module object.
We patch via ``"src.microsoft_mcp.intel.engine.<fn>"`` which affects the same
object that tools.py holds a reference to.
"""

from __future__ import annotations

from datetime import UTC, datetime
from unittest.mock import MagicMock, patch

import pytest

from src.microsoft_mcp import tools
from src.microsoft_mcp.intel.types import (
    BriefingReport,
    CalendarSignals,
    ContactReport,
    EmailSignals,
    PriorityItem,
    RecapReport,
    RelationshipInsight,
    SignalsReport,
    ThreadSignals,
)

# ---------------------------------------------------------------------------
# Canned fixtures
# ---------------------------------------------------------------------------

NOW = datetime(2026, 6, 18, 9, 0, 0, tzinfo=UTC)

_PRIORITY_ITEM: PriorityItem = {
    "score": 90.0,
    "source": "email",
    "category": "needs_response",
    "title": "Urgent email",
    "description": "Reply needed",
    "action_hint": "reply",
}

_PRIORITY_ITEM_MED: PriorityItem = {
    "score": 60.0,
    "source": "calendar",
    "category": "conflict",
    "title": "Calendar conflict",
    "description": "Two meetings overlap",
}

_PRIORITY_ITEM_LOW: PriorityItem = {
    "score": 30.0,
    "source": "thread",
    "category": "stale_thread",
    "title": "Stale thread",
    "description": "No reply in 7 days",
}


def _make_email_signals() -> EmailSignals:
    return EmailSignals(
        unread_total=5,
        unread_by_folder=[],
        needs_response=[],
        vip_unread=[],
        received_last_24h=10,
        sent_last_24h=3,
    )


def _make_calendar_signals() -> CalendarSignals:
    return CalendarSignals(
        today_events=[],
        tomorrow_events=[],
        conflicts=[],
        prep_needed=[],
        free_blocks=[],
        meeting_hours_today=2.5,
        total_events_today=3,
    )


def _make_thread_signals() -> ThreadSignals:
    return ThreadSignals(
        awaiting_my_reply=[],
        awaiting_their_reply=[],
        stale_threads=[],
    )


def _make_briefing_report(
    items: list[PriorityItem] | None = None,
) -> BriefingReport:
    return BriefingReport(
        generated_at=NOW.isoformat(),
        account="test@example.com",
        priority_items=items
        or [_PRIORITY_ITEM, _PRIORITY_ITEM_MED, _PRIORITY_ITEM_LOW],
        email_summary=_make_email_signals(),
        calendar_summary=_make_calendar_signals(),
        schedule_analysis={
            "meeting_density_pct": 31.0,
            "back_to_back_count": 1,
            "external_meeting_count": 0,
            "longest_free_block_minutes": 90,
            "focus_time_available": False,
            "summary": "Moderate day",
        },
        thread_summary=_make_thread_signals(),
    )


def _make_signals_report() -> SignalsReport:
    return SignalsReport(
        generated_at=NOW.isoformat(),
        account="test@example.com",
        critical=[_PRIORITY_ITEM],
        important=[_PRIORITY_ITEM_MED],
        informational=[_PRIORITY_ITEM_LOW],
        total_signals=3,
    )


def _make_contact_report() -> ContactReport:
    rel: RelationshipInsight = {
        "email": "alice@example.com",
        "name": "Alice",
        "engagement_score": 75.0,
        "trend": "rising",
        "last_interaction": NOW.isoformat(),
        "days_since_contact": 2,
        "sent_to": 5,
        "received_from": 8,
        "response_ratio": 1.6,
    }
    return ContactReport(
        generated_at=NOW.isoformat(),
        account="test@example.com",
        target_email="alice@example.com",
        target_name="Alice",
        relationship=rel,
        recent_threads=[],
        recent_emails_from=8,
        recent_emails_to=5,
        pending_items=[],
    )


def _make_recap_report() -> RecapReport:
    return RecapReport(
        generated_at=NOW.isoformat(),
        account="test@example.com",
        emails_received_today=12,
        emails_sent_today=4,
        emails_still_unread=3,
        meetings_attended=2,
        threads_resolved=0,
        threads_still_pending=[],
        tomorrow_preview=[],
    )


# ---------------------------------------------------------------------------
# Tests for generate_morning_briefing
# ---------------------------------------------------------------------------


class TestGenerateMorningBriefing:
    def test_calls_engine_with_graph_request(self, monkeypatch):
        report = _make_briefing_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_briefing", mock_fn):
            result = tools.generate_morning_briefing.fn(timezone="UTC", limit=10)

        mock_fn.assert_called_once()
        call_kwargs = mock_fn.call_args
        from src.microsoft_mcp import graph

        assert call_kwargs.args[0] is graph.request
        assert call_kwargs.kwargs["account"] == "user@example.com"
        assert call_kwargs.kwargs["timezone"] == "UTC"
        assert "now" in call_kwargs.kwargs
        assert isinstance(result, dict)

    def test_truncates_priority_items_to_limit(self, monkeypatch):
        items = [_PRIORITY_ITEM, _PRIORITY_ITEM_MED, _PRIORITY_ITEM_LOW]
        report = _make_briefing_report(items=items)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch(
            "src.microsoft_mcp.intel.engine.generate_briefing", return_value=report
        ):
            result = tools.generate_morning_briefing.fn(timezone="UTC", limit=2)

        assert len(result["priority_items"]) == 2
        assert result["priority_items"][0]["score"] == 90.0
        assert result["priority_items"][1]["score"] == 60.0

    def test_limit_zero_returns_empty_priority_items(self, monkeypatch):
        report = _make_briefing_report()
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch(
            "src.microsoft_mcp.intel.engine.generate_briefing", return_value=report
        ):
            result = tools.generate_morning_briefing.fn(timezone="UTC", limit=0)

        assert result["priority_items"] == []

    def test_limit_larger_than_items_returns_all(self, monkeypatch):
        report = _make_briefing_report()
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch(
            "src.microsoft_mcp.intel.engine.generate_briefing", return_value=report
        ):
            result = tools.generate_morning_briefing.fn(timezone="UTC", limit=100)

        assert len(result["priority_items"]) == 3

    def test_account_falls_back_to_default(self, monkeypatch):
        report = _make_briefing_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.delenv("MICROSOFT_MCP_ACCOUNT_ID", raising=False)
        with patch("src.microsoft_mcp.intel.engine.generate_briefing", mock_fn):
            tools.generate_morning_briefing.fn(timezone="UTC", limit=10)

        assert mock_fn.call_args.kwargs["account"] == "default"

    def test_injects_now_as_datetime(self, monkeypatch):
        report = _make_briefing_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_briefing", mock_fn):
            tools.generate_morning_briefing.fn(timezone="UTC", limit=10)

        now_arg = mock_fn.call_args.kwargs["now"]
        assert isinstance(now_arg, datetime)
        assert now_arg.tzinfo is not None


# ---------------------------------------------------------------------------
# Tests for get_priority_signals
# ---------------------------------------------------------------------------


class TestGetPrioritySignals:
    def test_calls_engine_with_graph_request(self, monkeypatch):
        report = _make_signals_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_signals", mock_fn):
            result = tools.get_priority_signals.fn(timezone="UTC", level="all")

        from src.microsoft_mcp import graph

        assert mock_fn.call_args.args[0] is graph.request
        assert isinstance(result, dict)
        assert "critical" in result

    def test_level_all_returns_full_report(self, monkeypatch):
        report = _make_signals_report()
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch(
            "src.microsoft_mcp.intel.engine.generate_signals", return_value=report
        ):
            result = tools.get_priority_signals.fn(timezone="UTC", level="all")

        assert len(result["critical"]) == 1
        assert len(result["important"]) == 1
        assert len(result["informational"]) == 1

    def test_level_critical_empties_other_buckets(self, monkeypatch):
        report = _make_signals_report()
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch(
            "src.microsoft_mcp.intel.engine.generate_signals", return_value=report
        ):
            result = tools.get_priority_signals.fn(timezone="UTC", level="critical")

        assert len(result["critical"]) == 1
        assert result["important"] == []
        assert result["informational"] == []

    def test_level_important_empties_other_buckets(self, monkeypatch):
        report = _make_signals_report()
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch(
            "src.microsoft_mcp.intel.engine.generate_signals", return_value=report
        ):
            result = tools.get_priority_signals.fn(timezone="UTC", level="important")

        assert result["critical"] == []
        assert len(result["important"]) == 1
        assert result["informational"] == []

    def test_level_informational_empties_other_buckets(self, monkeypatch):
        report = _make_signals_report()
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch(
            "src.microsoft_mcp.intel.engine.generate_signals", return_value=report
        ):
            result = tools.get_priority_signals.fn(
                timezone="UTC", level="informational"
            )

        assert result["critical"] == []
        assert result["important"] == []
        assert len(result["informational"]) == 1

    def test_invalid_level_raises_value_error(self, monkeypatch):
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with pytest.raises(ValueError, match="Invalid level"):
            tools.get_priority_signals.fn(timezone="UTC", level="urgent")

    def test_invalid_level_checked_before_engine_call(self, monkeypatch):
        mock_fn = MagicMock()
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_signals", mock_fn):
            with pytest.raises(ValueError):
                tools.get_priority_signals.fn(timezone="UTC", level="bad")
        mock_fn.assert_not_called()


# ---------------------------------------------------------------------------
# Tests for get_contact_intelligence
# ---------------------------------------------------------------------------


class TestGetContactIntelligence:
    def test_calls_engine_with_graph_request(self, monkeypatch):
        report = _make_contact_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_contact_report", mock_fn):
            result = tools.get_contact_intelligence.fn(
                target_email="alice@example.com", days=30
            )

        from src.microsoft_mcp import graph

        assert mock_fn.call_args.args[0] is graph.request
        assert mock_fn.call_args.kwargs["target_email"] == "alice@example.com"
        assert mock_fn.call_args.kwargs["lookback_days"] == 30
        assert mock_fn.call_args.kwargs["account"] == "user@example.com"
        assert isinstance(result, dict)

    def test_custom_days_forwarded_to_engine(self, monkeypatch):
        report = _make_contact_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_contact_report", mock_fn):
            tools.get_contact_intelligence.fn(target_email="bob@example.com", days=7)

        assert mock_fn.call_args.kwargs["lookback_days"] == 7

    def test_now_uses_utc(self, monkeypatch):
        report = _make_contact_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_contact_report", mock_fn):
            tools.get_contact_intelligence.fn(target_email="alice@example.com")

        now_arg = mock_fn.call_args.kwargs["now"]
        assert isinstance(now_arg, datetime)
        # UTC timezone — offset is 0
        assert now_arg.utcoffset().total_seconds() == 0  # type: ignore[union-attr]

    def test_returns_dict(self, monkeypatch):
        report = _make_contact_report()
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch(
            "src.microsoft_mcp.intel.engine.generate_contact_report",
            return_value=report,
        ):
            result = tools.get_contact_intelligence.fn(target_email="alice@example.com")

        assert result["target_email"] == "alice@example.com"
        assert "relationship" in result


# ---------------------------------------------------------------------------
# Tests for get_end_of_day_recap
# ---------------------------------------------------------------------------


class TestGetEndOfDayRecap:
    def test_calls_engine_with_graph_request(self, monkeypatch):
        report = _make_recap_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_recap", mock_fn):
            result = tools.get_end_of_day_recap.fn(timezone="UTC")

        from src.microsoft_mcp import graph

        assert mock_fn.call_args.args[0] is graph.request
        assert mock_fn.call_args.kwargs["account"] == "user@example.com"
        assert mock_fn.call_args.kwargs["timezone"] == "UTC"
        assert isinstance(result, dict)

    def test_returns_recap_keys(self, monkeypatch):
        report = _make_recap_report()
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch(
            "src.microsoft_mcp.intel.engine.generate_recap", return_value=report
        ):
            result = tools.get_end_of_day_recap.fn(timezone="UTC")

        assert result["emails_received_today"] == 12
        assert result["emails_sent_today"] == 4
        assert result["meetings_attended"] == 2

    def test_injects_now_as_aware_datetime(self, monkeypatch):
        report = _make_recap_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_recap", mock_fn):
            tools.get_end_of_day_recap.fn(timezone="UTC")

        now_arg = mock_fn.call_args.kwargs["now"]
        assert isinstance(now_arg, datetime)
        assert now_arg.tzinfo is not None

    def test_custom_timezone_forwarded(self, monkeypatch):
        report = _make_recap_report()
        mock_fn = MagicMock(return_value=report)
        monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "user@example.com")
        with patch("src.microsoft_mcp.intel.engine.generate_recap", mock_fn):
            tools.get_end_of_day_recap.fn(timezone="America/Chicago")

        assert mock_fn.call_args.kwargs["timezone"] == "America/Chicago"
