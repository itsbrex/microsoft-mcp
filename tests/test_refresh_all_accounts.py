"""Tests for refresh_all_accounts library function and MCP tool.

Test fixtures use a single non-production email (``broach@cresa.email``).
``refresh_all_accounts`` has known auth issues when more than one account is
saved on disk at once, so the supported pattern — and the only pattern
exercised here — is single-account.
"""

import json
from datetime import datetime, timedelta, timezone
from unittest.mock import patch

import pytest

from microsoft_mcp.auth_msal import MSALRefreshTokenAuth, refresh_all_accounts

# Canonical single-account email used across all test fixtures in this file.
TEST_EMAIL = "broach@cresa.email"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _write_token_file(tmp_path, identifier: str, expires_delta: timedelta) -> dict:
    """Write a minimal access-token JSON file and return its data."""
    expires_at = datetime.now(timezone.utc) + expires_delta
    data = {
        "email": identifier,
        "access_token": f"fake-access-token-{identifier}",
        "token_type": "Bearer",
        "expires_in": int(expires_delta.total_seconds()),
        "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "refreshed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "scopes": "https://graph.microsoft.com/.default",
        "api_type": "graph",
    }
    token_file = tmp_path / f"{identifier}_access_token.json"
    token_file.write_text(json.dumps(data))
    return data


def _write_refresh_token(
    tmp_path, identifier: str, value: str = "fake-refresh-token"
) -> None:
    """Write a refresh-token text file."""
    (tmp_path / f"{identifier}_refresh_only.txt").write_text(value)


# ---------------------------------------------------------------------------
# Library function tests
# ---------------------------------------------------------------------------


class TestRefreshAllAccountsLibrary:
    def test_empty_tokens_dir_returns_empty_list(self, tmp_path):
        """Directory exists but has no token files — should return []."""
        result = refresh_all_accounts(tokens_dir=tmp_path)
        assert result == []

    def test_nonexistent_tokens_dir_returns_empty_list(self, tmp_path):
        """Directory doesn't exist — should return [] without error."""
        missing = tmp_path / "does_not_exist"
        result = refresh_all_accounts(tokens_dir=missing)
        assert result == []

    def test_valid_token_reports_valid_status(self, tmp_path):
        """Token with future expiry — status should be 'valid', no network call."""
        email = TEST_EMAIL
        data = _write_token_file(tmp_path, email, timedelta(hours=1))

        result = refresh_all_accounts(tokens_dir=tmp_path)

        assert len(result) == 1
        entry = result[0]
        assert entry["identifier"] == email
        assert entry["status"] == "valid"
        assert entry["expires_at"] == data["expires_at"]
        assert entry["error"] is None

    def test_expired_token_triggers_refresh(self, tmp_path):
        """Expired token — _acquire_token_data is called and returns refreshed status."""
        email = TEST_EMAIL
        _write_token_file(tmp_path, email, timedelta(seconds=-10))
        _write_refresh_token(tmp_path, email)

        mock_token_response = {
            "access_token": "new-access-token",
            "refresh_token": "new-refresh-token",
            "expires_in": 3600,
            "scope": "https://graph.microsoft.com/.default",
        }

        with patch.object(
            MSALRefreshTokenAuth,
            "_refresh_access_token",
            return_value=mock_token_response,
        ):
            result = refresh_all_accounts(tokens_dir=tmp_path)

        assert len(result) == 1
        entry = result[0]
        assert entry["identifier"] == email
        assert entry["status"] == "refreshed"
        assert entry["expires_at"] is not None
        assert entry["error"] is None

    def test_refresh_failure_reports_failed_status(self, tmp_path):
        """Failed refresh on the saved account — status should be 'failed' with error."""
        email = TEST_EMAIL

        _write_token_file(tmp_path, email, timedelta(seconds=-5))
        _write_refresh_token(tmp_path, email)

        with patch.object(
            MSALRefreshTokenAuth,
            "_refresh_access_token",
            side_effect=RuntimeError("nope"),
        ):
            result = refresh_all_accounts(tokens_dir=tmp_path)

        assert len(result) == 1
        entry = result[0]
        assert entry["identifier"] == email
        assert entry["status"] == "failed"
        assert "nope" in (entry["error"] or "")

    def test_refresh_all_accounts_does_not_clear_cache_on_failure(self, tmp_path):
        """refresh_all_accounts must NOT delete token files when a per-account
        refresh fails — the refresh token must survive so the user can retry
        or re-authenticate manually (regression guard for Bug A)."""
        email = TEST_EMAIL

        # Write expired access-token JSON and the precious refresh token.
        _write_token_file(tmp_path, email, timedelta(seconds=-60))
        _write_refresh_token(tmp_path, email, value="precious-refresh-token")

        access_json = tmp_path / f"{email}_access_token.json"
        refresh_txt = tmp_path / f"{email}_refresh_only.txt"

        assert access_json.exists(), "precondition: access_token.json must exist"
        assert refresh_txt.exists(), "precondition: refresh_only.txt must exist"

        with patch.object(
            MSALRefreshTokenAuth,
            "_refresh_access_token",
            side_effect=RuntimeError("simulated network failure"),
        ):
            result = refresh_all_accounts(tokens_dir=tmp_path)

        assert len(result) == 1
        entry = result[0]
        assert entry["status"] == "failed"
        assert "simulated network failure" in (entry["error"] or "")

        # Both files MUST still exist — clear_cache must NOT have been called.
        assert access_json.exists(), (
            "access_token.json was deleted by refresh_all_accounts on failure — "
            "clear_cache() must not be called from this path"
        )
        assert refresh_txt.exists(), (
            "refresh_only.txt was deleted by refresh_all_accounts on failure — "
            "the refresh token must be preserved for recovery"
        )

    def test_acquire_token_data_clears_cache_then_reauths_on_refresh_failure(
        self, tmp_path, monkeypatch
    ):
        """_acquire_token_data clears cache on refresh failure, then calls authenticate().

        Evicting the corrupted token ensures the interactive re-auth path starts
        clean. authenticate() is mocked here so no real device-code flow runs.
        """
        monkeypatch.delenv("MICROSOFT_MCP_NONINTERACTIVE", raising=False)
        email = TEST_EMAIL

        _write_token_file(tmp_path, email, timedelta(seconds=-60))
        _write_refresh_token(tmp_path, email, value="stale-refresh-token")

        access_json = tmp_path / f"{email}_access_token.json"
        refresh_txt = tmp_path / f"{email}_refresh_only.txt"

        auth = MSALRefreshTokenAuth(
            tokens_dir=tmp_path,
            account_identifier=email,
        )

        # authenticate() raises so we can inspect file state after clear_cache()
        # without needing to fake the post-auth token reload.
        with (
            patch.object(
                auth,
                "_refresh_access_token",
                side_effect=RuntimeError("simulated failure"),
            ),
            patch.object(
                auth, "authenticate", side_effect=RuntimeError("auth cancelled")
            ),
            pytest.raises(RuntimeError),
        ):
            auth._acquire_token_data()

        # Files must be gone — clear_cache() must have fired before authenticate().
        assert not access_json.exists(), (
            "access_token.json still exists after _acquire_token_data failure — "
            "clear_cache() was not called on the lazy-refresh path"
        )
        assert not refresh_txt.exists(), (
            "refresh_only.txt still exists after _acquire_token_data failure — "
            "clear_cache() was not called on the lazy-refresh path"
        )


# ---------------------------------------------------------------------------
# MCP tool tests
# ---------------------------------------------------------------------------


class TestRefreshAllAccountsMcpTool:
    def test_mcp_tool_raises_under_azure_auth_method(self):
        """Tool must refuse when auth_method is not 'msal'."""
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import refresh_all_accounts as tool_fn

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "azure"
            with pytest.raises(ValueError, match="MSAL authentication method"):
                tool_fn.fn()
        finally:
            tools_mod.auth_method = original

    def test_mcp_tool_delegates_to_library_function(self, tmp_path, monkeypatch):
        """Tool should delegate to auth_msal.refresh_all_accounts with env-resolved args."""
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import refresh_all_accounts as tool_fn

        monkeypatch.setenv("MICROSOFT_MCP_TOKENS_DIR", str(tmp_path))
        monkeypatch.setenv("MICROSOFT_MCP_CLIENT_ID", "test-client-id")
        monkeypatch.setenv("MICROSOFT_MCP_TENANT_ID", "test-tenant-id")

        fake_result = [
            {
                "identifier": TEST_EMAIL,
                "status": "valid",
                "expires_at": "2099-01-01T00:00:00Z",
                "error": None,
            }
        ]

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "msal"
            with patch(
                "microsoft_mcp.auth_msal.refresh_all_accounts",
                return_value=fake_result,
            ) as mock_lib:
                returned = tool_fn.fn()

            mock_lib.assert_called_once()
            call_kwargs = mock_lib.call_args
            # tokens_dir should be the env-resolved Path
            import pathlib as pl

            assert call_kwargs.kwargs.get("tokens_dir") == pl.Path(str(tmp_path)) or (
                call_kwargs.args and str(call_kwargs.args[0]) == str(tmp_path)
            )
            assert returned == fake_result
        finally:
            tools_mod.auth_method = original
