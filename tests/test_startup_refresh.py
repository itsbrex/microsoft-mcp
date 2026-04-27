"""Tests for the _maybe_refresh_on_startup() hook in server.py."""

import shutil
import subprocess
from pathlib import Path
from unittest.mock import patch

import pytest

from microsoft_mcp.server import _maybe_refresh_on_startup

REPO_ROOT = Path(__file__).resolve().parent.parent
CONSOLE_SCRIPT = shutil.which("microsoft-mcp")

# ---------------------------------------------------------------------------
# Fake result lists
# ---------------------------------------------------------------------------

_FAKE_RESULTS_MIXED = [
    {
        "identifier": "a@x.com",
        "status": "valid",
        "expires_at": "2099-01-01T00:00:00Z",
        "error": None,
    },
    {
        "identifier": "b@x.com",
        "status": "valid",
        "expires_at": "2099-01-01T00:00:00Z",
        "error": None,
    },
    {
        "identifier": "c@x.com",
        "status": "refreshed",
        "expires_at": "2099-01-01T01:00:00Z",
        "error": None,
    },
    {
        "identifier": "d@x.com",
        "status": "failed",
        "expires_at": None,
        "error": "connection refused",
    },
]


# ---------------------------------------------------------------------------
# Test cases
# ---------------------------------------------------------------------------


class TestMaybeRefreshOnStartup:
    def test_skips_for_azure_auth_method(self, monkeypatch):
        """Hook must not call refresh_all_accounts when auth method is azure."""
        monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", "azure")

        with patch("microsoft_mcp.auth_msal.refresh_all_accounts") as mock_refresh:
            _maybe_refresh_on_startup()

        mock_refresh.assert_not_called()

    def test_skips_when_auth_method_unset(self, monkeypatch):
        """Hook must not call refresh_all_accounts when MICROSOFT_MCP_AUTH_METHOD is not set (defaults to azure)."""
        monkeypatch.delenv("MICROSOFT_MCP_AUTH_METHOD", raising=False)

        with patch("microsoft_mcp.auth_msal.refresh_all_accounts") as mock_refresh:
            _maybe_refresh_on_startup()

        mock_refresh.assert_not_called()

    def test_skips_when_explicitly_opted_out(self, monkeypatch):
        """Hook must not call refresh_all_accounts when MICROSOFT_MCP_REFRESH_ON_STARTUP=0."""
        monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", "msal")
        monkeypatch.setenv("MICROSOFT_MCP_REFRESH_ON_STARTUP", "0")

        with patch("microsoft_mcp.auth_msal.refresh_all_accounts") as mock_refresh:
            _maybe_refresh_on_startup()

        mock_refresh.assert_not_called()

    def test_default_on_for_msal_calls_refresh(self, monkeypatch):
        """Hook must call refresh_all_accounts (once, no args) when MSAL and refresh not opted out."""
        monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", "msal")
        monkeypatch.delenv("MICROSOFT_MCP_REFRESH_ON_STARTUP", raising=False)

        fake_results = [
            {
                "identifier": "a@x.com",
                "status": "valid",
                "expires_at": "2099-01-01T00:00:00Z",
                "error": None,
            },
            {
                "identifier": "b@x.com",
                "status": "refreshed",
                "expires_at": "2099-01-01T01:00:00Z",
                "error": None,
            },
        ]

        # The hook does `from microsoft_mcp.auth_msal import refresh_all_accounts`
        # inside the try block each call, so patching the module attribute is sufficient.
        with patch(
            "microsoft_mcp.auth_msal.refresh_all_accounts", return_value=fake_results
        ) as mock_refresh:
            _maybe_refresh_on_startup()

        mock_refresh.assert_called_once_with()

    def test_empty_results_stay_silent(self, monkeypatch, capsys):
        """Fresh install (no saved tokens) must not print startup banner or summary."""
        monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", "msal")
        monkeypatch.delenv("MICROSOFT_MCP_REFRESH_ON_STARTUP", raising=False)

        with patch("microsoft_mcp.auth_msal.refresh_all_accounts", return_value=[]):
            _maybe_refresh_on_startup()

        captured = capsys.readouterr()
        assert captured.err == "", (
            f"Expected silent stderr on empty results, got: {captured.err!r}"
        )

    def test_refresh_failure_does_not_raise(self, monkeypatch, caplog):
        """A RuntimeError from refresh_all_accounts must not escape the hook."""
        monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", "msal")
        monkeypatch.delenv("MICROSOFT_MCP_REFRESH_ON_STARTUP", raising=False)

        with patch(
            "microsoft_mcp.auth_msal.refresh_all_accounts",
            side_effect=RuntimeError("boom"),
        ):
            import logging

            with caplog.at_level(logging.WARNING, logger="microsoft_mcp.server"):
                # Must not raise
                _maybe_refresh_on_startup()

        assert any("boom" in record.message for record in caplog.records), (
            "Expected warning containing 'boom' to be logged"
        )

    def test_summary_line_counts_correctly(self, monkeypatch, capsys):
        """Summary line must reflect the correct counts of valid/refreshed/failed."""
        monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", "msal")
        monkeypatch.delenv("MICROSOFT_MCP_REFRESH_ON_STARTUP", raising=False)

        with patch(
            "microsoft_mcp.auth_msal.refresh_all_accounts",
            return_value=_FAKE_RESULTS_MIXED,
        ):
            _maybe_refresh_on_startup()

        captured = capsys.readouterr()
        assert "2 valid, 1 refreshed, 1 failed" in captured.err, (
            f"Expected summary '2 valid, 1 refreshed, 1 failed' in stderr.\nGot: {captured.err!r}"
        )

    @pytest.mark.skipif(
        CONSOLE_SCRIPT is None, reason="microsoft-mcp console script not on PATH"
    )
    def test_subprocess_startup_hook_does_not_crash(self, tmp_path):
        """Server process with MSAL+empty tokens dir must not traceback; must reach CLIENT_ID guard."""
        assert CONSOLE_SCRIPT is not None  # guarded by skipif above
        env = {
            "PATH": "/usr/bin:/bin",
            "HOME": "/tmp",
            "MICROSOFT_MCP_AUTH_METHOD": "msal",
            "MICROSOFT_MCP_REFRESH_ON_STARTUP": "1",
            "MICROSOFT_MCP_TOKENS_DIR": str(tmp_path),
        }
        proc = subprocess.run(
            [CONSOLE_SCRIPT],
            env=env,
            capture_output=True,
            text=True,
            timeout=15,
        )

        assert "Traceback" not in proc.stderr, (
            f"Unexpected traceback in stderr:\n{proc.stderr}"
        )
        assert "MICROSOFT_MCP_CLIENT_ID" in proc.stderr, (
            f"Expected CLIENT_ID guard message in stderr. Got:\n{proc.stderr}"
        )
