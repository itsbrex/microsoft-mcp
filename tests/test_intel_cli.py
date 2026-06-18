"""Tests for intel_cli.py subcommand dispatch and JSON output.

The CLI makes live Graph calls, so we mock both _bootstrap_graph and the
engine functions.  Each test verifies the correct engine function is called
and that --json emits valid JSON to stdout.
"""

from __future__ import annotations

import json
import sys
from datetime import UTC, datetime
from io import StringIO
from typing import Any
from unittest.mock import MagicMock, patch

import pytest

from src.microsoft_mcp import intel_cli
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

NOW = datetime(2026, 6, 18, 9, 0, 0, tzinfo=UTC)

# ---------------------------------------------------------------------------
# Canned fixtures (minimal, re-used across tests)
# ---------------------------------------------------------------------------

_ITEM: PriorityItem = {
    "score": 85.0,
    "source": "email",
    "category": "needs_response",
    "title": "Important email",
    "description": "Please reply",
}

_ITEM_MED: PriorityItem = {
    "score": 55.0,
    "source": "calendar",
    "category": "conflict",
    "title": "Calendar conflict",
    "description": "Overlap",
}

_ITEM_LOW: PriorityItem = {
    "score": 20.0,
    "source": "thread",
    "category": "stale_thread",
    "title": "Old thread",
    "description": "Stale",
}


def _email_signals() -> EmailSignals:
    return EmailSignals(
        unread_total=2,
        unread_by_folder=[],
        needs_response=[],
        vip_unread=[],
        received_last_24h=5,
        sent_last_24h=2,
    )


def _cal_signals() -> CalendarSignals:
    return CalendarSignals(
        today_events=[],
        tomorrow_events=[],
        conflicts=[],
        prep_needed=[],
        free_blocks=[],
        meeting_hours_today=1.0,
        total_events_today=2,
    )


def _thread_signals() -> ThreadSignals:
    return ThreadSignals(
        awaiting_my_reply=[],
        awaiting_their_reply=[],
        stale_threads=[],
    )


def _briefing_report() -> BriefingReport:
    return BriefingReport(
        generated_at=NOW.isoformat(),
        account="test@example.com",
        priority_items=[_ITEM, _ITEM_MED, _ITEM_LOW],
        email_summary=_email_signals(),
        calendar_summary=_cal_signals(),
        schedule_analysis={
            "meeting_density_pct": 12.5,
            "back_to_back_count": 0,
            "external_meeting_count": 0,
            "longest_free_block_minutes": 120,
            "focus_time_available": True,
            "summary": "Light day",
        },
        thread_summary=_thread_signals(),
    )


def _signals_report() -> SignalsReport:
    return SignalsReport(
        generated_at=NOW.isoformat(),
        account="test@example.com",
        critical=[_ITEM],
        important=[_ITEM_MED],
        informational=[_ITEM_LOW],
        total_signals=3,
    )


def _contact_report() -> ContactReport:
    rel: RelationshipInsight = {
        "email": "alice@example.com",
        "name": "Alice",
        "engagement_score": 60.0,
        "trend": "stable",
        "last_interaction": NOW.isoformat(),
        "days_since_contact": 5,
        "sent_to": 3,
        "received_from": 4,
        "response_ratio": 1.3,
    }
    return ContactReport(
        generated_at=NOW.isoformat(),
        account="test@example.com",
        target_email="alice@example.com",
        target_name="Alice",
        relationship=rel,
        recent_threads=[],
        recent_emails_from=4,
        recent_emails_to=3,
        pending_items=[],
    )


def _recap_report() -> RecapReport:
    return RecapReport(
        generated_at=NOW.isoformat(),
        account="test@example.com",
        emails_received_today=8,
        emails_sent_today=3,
        emails_still_unread=1,
        meetings_attended=2,
        threads_resolved=0,
        threads_still_pending=[],
        tomorrow_preview=[],
    )


# ---------------------------------------------------------------------------
# Helper: run main() capturing stdout
# ---------------------------------------------------------------------------


def _run(argv: list[str], monkeypatch, mock_graph: Any = None) -> tuple[int, str]:
    """Run intel_cli.main(argv), capture stdout, return (exit_code, output)."""
    fake_graph = mock_graph or MagicMock()
    fake_graph.request = MagicMock()

    buf = StringIO()
    with patch.object(intel_cli, "_bootstrap_graph", return_value=fake_graph):
        with patch("sys.stdout", buf):
            code = intel_cli.main(argv)
    return code, buf.getvalue()


# ---------------------------------------------------------------------------
# briefing subcommand
# ---------------------------------------------------------------------------


class TestBriefingSubcommand:
    def test_dispatches_to_generate_briefing(self, monkeypatch):
        report = _briefing_report()
        mock_engine = MagicMock(return_value=report)
        with patch("microsoft_mcp.intel.engine.generate_briefing", mock_engine):
            code, _ = _run(["briefing", "--json"], monkeypatch)
        assert code == 0
        mock_engine.assert_called_once()

    def test_json_flag_emits_valid_json(self, monkeypatch):
        report = _briefing_report()
        with patch("microsoft_mcp.intel.engine.generate_briefing", return_value=report):
            code, out = _run(["briefing", "--json"], monkeypatch)
        assert code == 0
        parsed = json.loads(out)
        assert "priority_items" in parsed
        assert "generated_at" in parsed

    def test_limit_arg_truncates_priority_items(self, monkeypatch):
        report = _briefing_report()
        with patch("microsoft_mcp.intel.engine.generate_briefing", return_value=report):
            code, out = _run(["briefing", "--json", "--limit", "1"], monkeypatch)
        assert code == 0
        parsed = json.loads(out)
        assert len(parsed["priority_items"]) == 1

    def test_default_limit_is_10(self, monkeypatch):
        # Build report with 15 items
        items = [_ITEM] * 15
        report = _briefing_report()
        report = dict(report)
        report["priority_items"] = items

        with patch("microsoft_mcp.intel.engine.generate_briefing", return_value=report):
            code, out = _run(["briefing", "--json"], monkeypatch)
        assert code == 0
        parsed = json.loads(out)
        assert len(parsed["priority_items"]) == 10

    def test_timezone_arg_forwarded(self, monkeypatch):
        report = _briefing_report()
        mock_engine = MagicMock(return_value=report)
        with patch("microsoft_mcp.intel.engine.generate_briefing", mock_engine):
            _run(["briefing", "--json", "--timezone", "America/Chicago"], monkeypatch)
        assert mock_engine.call_args.kwargs["timezone"] == "America/Chicago"

    def test_human_output_without_json_flag(self, monkeypatch):
        report = _briefing_report()
        with patch("microsoft_mcp.intel.engine.generate_briefing", return_value=report):
            code, out = _run(["briefing"], monkeypatch)
        assert code == 0
        # Not valid JSON — human-readable text
        with pytest.raises((json.JSONDecodeError, ValueError)):
            json.loads(out)
        assert "Briefing" in out or "priority" in out.lower() or "item" in out.lower()


# ---------------------------------------------------------------------------
# signals subcommand
# ---------------------------------------------------------------------------


class TestSignalsSubcommand:
    def test_json_flag_emits_valid_json(self, monkeypatch):
        report = _signals_report()
        with patch("microsoft_mcp.intel.engine.generate_signals", return_value=report):
            code, out = _run(["signals", "--json"], monkeypatch)
        assert code == 0
        parsed = json.loads(out)
        assert "critical" in parsed
        assert "important" in parsed
        assert "informational" in parsed

    def test_level_critical_filters_output(self, monkeypatch):
        report = _signals_report()
        with patch("microsoft_mcp.intel.engine.generate_signals", return_value=report):
            code, out = _run(["signals", "--json", "--level", "critical"], monkeypatch)
        assert code == 0
        parsed = json.loads(out)
        assert len(parsed["critical"]) == 1
        assert parsed["important"] == []
        assert parsed["informational"] == []

    def test_level_important_filters_output(self, monkeypatch):
        report = _signals_report()
        with patch("microsoft_mcp.intel.engine.generate_signals", return_value=report):
            code, out = _run(["signals", "--json", "--level", "important"], monkeypatch)
        assert code == 0
        parsed = json.loads(out)
        assert parsed["critical"] == []
        assert len(parsed["important"]) == 1
        assert parsed["informational"] == []

    def test_level_all_returns_all_buckets(self, monkeypatch):
        report = _signals_report()
        with patch("microsoft_mcp.intel.engine.generate_signals", return_value=report):
            code, out = _run(["signals", "--json", "--level", "all"], monkeypatch)
        assert code == 0
        parsed = json.loads(out)
        assert len(parsed["critical"]) == 1
        assert len(parsed["important"]) == 1
        assert len(parsed["informational"]) == 1

    def test_invalid_level_exits_nonzero(self, monkeypatch):
        # argparse will reject unknown choices before engine is called
        with pytest.raises(SystemExit) as exc_info:
            intel_cli.main(["signals", "--level", "bogus"])
        assert exc_info.value.code != 0


# ---------------------------------------------------------------------------
# contact subcommand
# ---------------------------------------------------------------------------


class TestContactSubcommand:
    def test_json_flag_emits_valid_json(self, monkeypatch):
        report = _contact_report()
        with patch(
            "microsoft_mcp.intel.engine.generate_contact_report", return_value=report
        ):
            code, out = _run(["contact", "alice@example.com", "--json"], monkeypatch)
        assert code == 0
        parsed = json.loads(out)
        assert parsed["target_email"] == "alice@example.com"

    def test_email_arg_forwarded_to_engine(self, monkeypatch):
        report = _contact_report()
        mock_engine = MagicMock(return_value=report)
        with patch("microsoft_mcp.intel.engine.generate_contact_report", mock_engine):
            _run(["contact", "bob@example.com", "--json"], monkeypatch)
        assert mock_engine.call_args.kwargs["target_email"] == "bob@example.com"

    def test_days_arg_forwarded_to_engine(self, monkeypatch):
        report = _contact_report()
        mock_engine = MagicMock(return_value=report)
        with patch("microsoft_mcp.intel.engine.generate_contact_report", mock_engine):
            _run(["contact", "alice@example.com", "--days", "7", "--json"], monkeypatch)
        assert mock_engine.call_args.kwargs["lookback_days"] == 7

    def test_default_days_is_30(self, monkeypatch):
        report = _contact_report()
        mock_engine = MagicMock(return_value=report)
        with patch("microsoft_mcp.intel.engine.generate_contact_report", mock_engine):
            _run(["contact", "alice@example.com", "--json"], monkeypatch)
        assert mock_engine.call_args.kwargs["lookback_days"] == 30

    def test_human_output_without_json_flag(self, monkeypatch):
        report = _contact_report()
        with patch(
            "microsoft_mcp.intel.engine.generate_contact_report", return_value=report
        ):
            code, out = _run(["contact", "alice@example.com"], monkeypatch)
        assert code == 0
        assert "Alice" in out or "alice@example.com" in out


# ---------------------------------------------------------------------------
# recap subcommand
# ---------------------------------------------------------------------------


class TestRecapSubcommand:
    def test_json_flag_emits_valid_json(self, monkeypatch):
        report = _recap_report()
        with patch("microsoft_mcp.intel.engine.generate_recap", return_value=report):
            code, out = _run(["recap", "--json"], monkeypatch)
        assert code == 0
        parsed = json.loads(out)
        assert "emails_received_today" in parsed
        assert parsed["emails_received_today"] == 8

    def test_timezone_arg_forwarded(self, monkeypatch):
        report = _recap_report()
        mock_engine = MagicMock(return_value=report)
        with patch("microsoft_mcp.intel.engine.generate_recap", mock_engine):
            _run(["recap", "--json", "--timezone", "Europe/London"], monkeypatch)
        assert mock_engine.call_args.kwargs["timezone"] == "Europe/London"

    def test_human_output_without_json_flag(self, monkeypatch):
        report = _recap_report()
        with patch("microsoft_mcp.intel.engine.generate_recap", return_value=report):
            code, out = _run(["recap"], monkeypatch)
        assert code == 0
        # Should mention recap or today
        assert "Recap" in out or "today" in out.lower() or "email" in out.lower()

    def test_exit_code_zero_on_success(self, monkeypatch):
        report = _recap_report()
        with patch("microsoft_mcp.intel.engine.generate_recap", return_value=report):
            code, _ = _run(["recap", "--json"], monkeypatch)
        assert code == 0


# ---------------------------------------------------------------------------
# CLI structure tests
# ---------------------------------------------------------------------------


class TestCliStructure:
    def test_no_subcommand_exits_nonzero(self):
        with pytest.raises(SystemExit) as exc_info:
            intel_cli.main([])
        assert exc_info.value.code != 0

    def test_unknown_subcommand_exits_nonzero(self):
        with pytest.raises(SystemExit) as exc_info:
            intel_cli.main(["unknown"])
        assert exc_info.value.code != 0

    def test_cli_main_calls_sys_exit(self, monkeypatch):
        report = _recap_report()
        exited_with: list[Any] = []
        monkeypatch.setattr(sys, "argv", ["microsoft-mcp-intel", "recap", "--json"])
        with patch("microsoft_mcp.intel.engine.generate_recap", return_value=report):
            with patch.object(
                intel_cli,
                "_bootstrap_graph",
                return_value=MagicMock(request=MagicMock()),
            ):
                with patch("sys.exit", side_effect=lambda c: exited_with.append(c)):
                    with patch("sys.stdout", StringIO()):
                        try:
                            intel_cli.cli_main()
                        except Exception:
                            pass
        assert exited_with == [0]
