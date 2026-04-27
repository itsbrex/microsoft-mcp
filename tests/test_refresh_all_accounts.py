"""Tests for refresh_all_accounts library function and MCP tool."""

import json
from datetime import datetime, timedelta, timezone
from unittest.mock import patch

import pytest

from microsoft_mcp.auth_msal import MSALRefreshTokenAuth, refresh_all_accounts


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
        email = "user@example.com"
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
        email = "expired@example.com"
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

    def test_one_account_failure_does_not_block_others(self, tmp_path):
        """Failed refresh on one account — other valid accounts still report 'valid'."""
        valid_email = "valid@example.com"
        broken_email = "broken@example.com"

        _write_token_file(tmp_path, valid_email, timedelta(hours=2))
        _write_token_file(tmp_path, broken_email, timedelta(seconds=-5))
        _write_refresh_token(tmp_path, broken_email)

        def _raise_on_broken(self, refresh_token):
            # refresh_token positional arg must match the real
            # _refresh_access_token signature for patch.object; mock body
            # intentionally ignores it.
            del refresh_token
            if self.account_identifier == broken_email:
                raise RuntimeError("nope")
            # Should not be called for the valid account, but guard anyway.
            raise AssertionError("unexpected call for valid account")

        with patch.object(
            MSALRefreshTokenAuth, "_refresh_access_token", _raise_on_broken
        ):
            result = refresh_all_accounts(tokens_dir=tmp_path)

        assert len(result) == 2

        by_id = {r["identifier"]: r for r in result}

        assert by_id[valid_email]["status"] == "valid"
        assert by_id[valid_email]["error"] is None

        assert by_id[broken_email]["status"] == "failed"
        assert "nope" in (by_id[broken_email]["error"] or "")


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
                "identifier": "someone@example.com",
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
