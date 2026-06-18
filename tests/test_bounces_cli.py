"""Tests for bounces_cli.py — scan + patterns subcommands.

Import convention follows the repo pattern:
    from src.microsoft_mcp import <module>

Monkeypatch targets use ``microsoft_mcp.*`` (not ``src.microsoft_mcp.*``) because
bounces_cli._cmd_scan does ``from microsoft_mcp import bounces`` at call time,
so the live module references live in the ``microsoft_mcp`` namespace.
"""

from __future__ import annotations

import json
import sys
from unittest import mock

import pytest

from src.microsoft_mcp import bounces, bounces_cli

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_SAMPLE_ROW = {
    "first_name": "Bob",
    "last_name": "Jones",
    "email": "bob.jones@company.com",
    "reason": "Invalid Recipient",
    "date": "2025-03-15T10:30:00Z",
    "iso_date": "2025-03-15T10:30:00Z",
    "subject": "Undeliverable: Hello bob",
    "sender": "postmaster@mailserver.com",
    "body": "Your message could not be delivered.",
    "message_id": "AABounce001",
    "has_attachments": False,
}


def _make_fake_graph():
    """Return a minimal fake graph module with a no-op request callable."""
    fake_graph = mock.MagicMock()
    fake_graph.request.return_value = {"value": []}
    return fake_graph


# ---------------------------------------------------------------------------
# scan subcommand
# ---------------------------------------------------------------------------


class TestCmdScan:
    def test_scan_dispatches_and_returns_zero(self, monkeypatch):
        """scan subcommand calls scan_folder and exits 0."""
        fake_graph = _make_fake_graph()
        monkeypatch.setattr(
            "microsoft_mcp.bounces_cli._bootstrap_graph", lambda: fake_graph
        )
        monkeypatch.setattr(
            "microsoft_mcp.bounces.scan_folder", lambda req, folder, limit=None: []
        )
        rc = bounces_cli.main(["scan", "--folder", "inbox", "--limit", "10"])
        assert rc == 0

    def test_scan_json_flag_emits_valid_json(self, monkeypatch, capsys):
        """--json flag produces parseable JSON with count and rows."""
        fake_graph = _make_fake_graph()
        monkeypatch.setattr(
            "microsoft_mcp.bounces_cli._bootstrap_graph", lambda: fake_graph
        )
        monkeypatch.setattr(
            "microsoft_mcp.bounces.scan_folder",
            lambda req, folder, limit=None: [_SAMPLE_ROW],
        )
        rc = bounces_cli.main(["scan", "--json"])
        assert rc == 0
        captured = capsys.readouterr()
        data = json.loads(captured.out)
        assert data["count"] == 1
        assert len(data["rows"]) == 1
        assert data["rows"][0]["email"] == "bob.jones@company.com"

    def test_scan_output_triggers_write_csv(self, monkeypatch, tmp_path):
        """--output CSV_PATH calls write_csv with the given path."""
        fake_graph = _make_fake_graph()
        monkeypatch.setattr(
            "microsoft_mcp.bounces_cli._bootstrap_graph", lambda: fake_graph
        )
        monkeypatch.setattr(
            "microsoft_mcp.bounces.scan_folder",
            lambda req, folder, limit=None: [_SAMPLE_ROW],
        )
        csv_path = tmp_path / "out.csv"
        write_csv_calls: list = []

        def fake_write_csv(rows, path):
            write_csv_calls.append((rows, str(path)))

        monkeypatch.setattr("microsoft_mcp.bounces.write_csv", fake_write_csv)
        rc = bounces_cli.main(["scan", "--output", str(csv_path)])
        assert rc == 0
        assert len(write_csv_calls) == 1
        assert write_csv_calls[0][0] == [_SAMPLE_ROW]

    def test_scan_human_output_shows_count(self, monkeypatch, capsys):
        """Human (non-JSON) output prints bounce count."""
        fake_graph = _make_fake_graph()
        monkeypatch.setattr(
            "microsoft_mcp.bounces_cli._bootstrap_graph", lambda: fake_graph
        )
        monkeypatch.setattr(
            "microsoft_mcp.bounces.scan_folder",
            lambda req, folder, limit=None: [_SAMPLE_ROW],
        )
        rc = bounces_cli.main(["scan"])
        assert rc == 0
        captured = capsys.readouterr()
        assert "1" in captured.out  # count appears somewhere

    def test_scan_default_folder_is_inbox(self, monkeypatch):
        """When --folder is omitted, 'inbox' is passed to scan_folder."""
        fake_graph = _make_fake_graph()
        monkeypatch.setattr(
            "microsoft_mcp.bounces_cli._bootstrap_graph", lambda: fake_graph
        )
        folders_used: list[str] = []

        def fake_scan(req, folder, limit=None):
            folders_used.append(folder)
            return []

        monkeypatch.setattr("microsoft_mcp.bounces.scan_folder", fake_scan)
        bounces_cli.main(["scan"])
        assert folders_used == ["inbox"]


# ---------------------------------------------------------------------------
# patterns subcommand
# ---------------------------------------------------------------------------


class TestCmdPatterns:
    def test_patterns_exits_zero(self):
        """patterns subcommand exits 0 without a graph bootstrap."""
        rc = bounces_cli.main(["patterns"])
        assert rc == 0

    def test_patterns_json_flag_emits_valid_json(self, capsys):
        """--json flag produces parseable JSON containing all catalog keys."""
        rc = bounces_cli.main(["patterns", "--json"])
        assert rc == 0
        captured = capsys.readouterr()
        data = json.loads(captured.out)
        assert "SUBJECT_KEYWORDS" in data
        assert "SENDER_PATTERNS" in data
        assert "BODY_PATTERNS" in data
        assert "BOUNCE_REASONS" in data
        assert "STRONG_SUBJECT_INDICATORS" in data
        assert "EXCLUDED_SUBJECT_PREFIXES" in data

    def test_patterns_json_reasons_is_list_of_dicts(self, capsys):
        """BOUNCE_REASONS in JSON output is a list of {pattern, reason} dicts."""
        bounces_cli.main(["patterns", "--json"])
        captured = capsys.readouterr()
        data = json.loads(captured.out)
        assert isinstance(data["BOUNCE_REASONS"], list)
        assert all(
            "pattern" in item and "reason" in item for item in data["BOUNCE_REASONS"]
        )

    def test_patterns_human_output_shows_keywords(self, capsys):
        """Human output prints at least one SUBJECT_KEYWORD."""
        bounces_cli.main(["patterns"])
        captured = capsys.readouterr()
        # At least one keyword should appear in the output
        assert any(kw in captured.out for kw in bounces.SUBJECT_KEYWORDS)

    def test_patterns_does_not_call_bootstrap(self, monkeypatch):
        """patterns never calls _bootstrap_graph (read-only subcommand)."""
        bootstrap_called = [False]

        def fake_bootstrap():
            bootstrap_called[0] = True
            return _make_fake_graph()

        monkeypatch.setattr(
            "microsoft_mcp.bounces_cli._bootstrap_graph", fake_bootstrap
        )
        bounces_cli.main(["patterns"])
        assert not bootstrap_called[0]


# ---------------------------------------------------------------------------
# Exit codes
# ---------------------------------------------------------------------------


class TestExitCodes:
    def test_missing_subcommand_exits_nonzero(self):
        """No subcommand → argparse exits with code 2."""
        with pytest.raises(SystemExit) as exc_info:
            bounces_cli.main([])
        assert exc_info.value.code != 0

    def test_unknown_subcommand_exits_nonzero(self):
        """Unknown subcommand → argparse exits with code 2."""
        with pytest.raises(SystemExit) as exc_info:
            bounces_cli.main(["nonexistent"])
        assert exc_info.value.code != 0

    def test_cli_main_calls_sys_exit(self, monkeypatch):
        """cli_main() wraps main() in sys.exit()."""
        monkeypatch.setattr("microsoft_mcp.bounces_cli.main", lambda argv: 0)
        monkeypatch.setattr(sys, "argv", ["microsoft-mcp-bounces", "patterns"])
        with pytest.raises(SystemExit) as exc_info:
            bounces_cli.cli_main()
        assert exc_info.value.code == 0
