"""Tests for authenticate_new_account MCP tool."""

import json
import os
from unittest.mock import MagicMock, patch

import pytest

from microsoft_mcp.auth_msal import MSALRefreshTokenAuth


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_mock_app(email: str = "new@example.com") -> MagicMock:
    """Return a mock PublicClientApplication that drives the device-code path."""
    mock_app = MagicMock()
    mock_app.get_accounts.return_value = []  # force device-code path
    mock_app.initiate_device_flow.return_value = {
        "user_code": "ABCD-1234",
        "device_code": "fake-device-code",
        "verification_uri": "https://microsoft.com/devicelogin",
        "message": "Go to https://microsoft.com/devicelogin and enter ABCD-1234",
    }
    mock_app.acquire_token_by_device_flow.return_value = {
        "access_token": "fake-access-token-xyz",
        "refresh_token": "fake-refresh-token",
        "expires_in": 3600,
        "scope": "User.Read",
        "id_token_claims": {"preferred_username": email},
    }
    return mock_app


# ---------------------------------------------------------------------------
# Test cases
# ---------------------------------------------------------------------------


class TestAuthenticateNewAccount:
    def test_raises_under_azure_auth_method(self):
        """Tool must refuse when auth_method is not 'msal'."""
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import authenticate_new_account

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "azure"
            with pytest.raises(ValueError, match="MSAL authentication method"):
                authenticate_new_account.fn(email="any@example.com")
        finally:
            tools_mod.auth_method = original

    @pytest.mark.parametrize("bad_email", ["", "   ", "\t\n"])
    def test_raises_on_empty_email(self, bad_email):
        """Tool must reject blank / whitespace-only email strings."""
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import authenticate_new_account

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "msal"
            with pytest.raises(ValueError, match="non-empty"):
                authenticate_new_account.fn(email=bad_email)
        finally:
            tools_mod.auth_method = original

    def test_authenticates_and_writes_token_file(self, tmp_path):
        """Successful flow persists an access-token file via real _save_tokens."""
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import authenticate_new_account

        email = "new@example.com"
        original_method = tools_mod.auth_method
        mock_app = _make_mock_app(email)

        try:
            tools_mod.auth_method = "msal"
            env_patch = {"MICROSOFT_MCP_TOKENS_DIR": str(tmp_path)}
            with patch.dict(os.environ, env_patch, clear=False):
                with patch.object(
                    MSALRefreshTokenAuth, "_get_msal_app", return_value=mock_app
                ):
                    result = authenticate_new_account.fn(email=email)

            token_file = tmp_path / f"{email}_access_token.json"
            assert token_file.exists(), (
                "Token file should have been written by _save_tokens"
            )

            # Verify the file was written by the real _save_tokens (not a fake)
            token_data = json.loads(token_file.read_text())
            assert token_data["access_token"] == "fake-access-token-xyz"
            assert token_data["email"] == email
            assert "expires_at" in token_data
            assert "refreshed_at" in token_data
            assert token_data["token_type"] == "Bearer"

            assert result["status"] == "authenticated"
            assert result["account"] == email
            assert result["active"] is False
            assert "set_active_account" in result["next"]
        finally:
            tools_mod.auth_method = original_method

    def test_does_not_swap_active_auth(self, tmp_path):
        """The global auth instance must not change after authenticate_new_account."""
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import authenticate_new_account

        original_method = tools_mod.auth_method
        original_auth = tools_mod.auth  # capture before the call
        mock_app = _make_mock_app("other@example.com")

        try:
            tools_mod.auth_method = "msal"
            env_patch = {"MICROSOFT_MCP_TOKENS_DIR": str(tmp_path)}
            with patch.dict(os.environ, env_patch, clear=False):
                with patch.object(
                    MSALRefreshTokenAuth, "_get_msal_app", return_value=mock_app
                ):
                    authenticate_new_account.fn(email="other@example.com")

            assert tools_mod.auth is original_auth, (
                "global auth must not be swapped by authenticate_new_account"
            )
        finally:
            tools_mod.auth_method = original_method

    def test_idempotent_for_existing_account(self, tmp_path):
        """Calling twice for the same email succeeds both times; only one file remains."""
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import authenticate_new_account

        email = "repeat@example.com"
        original_method = tools_mod.auth_method

        try:
            tools_mod.auth_method = "msal"
            env_patch = {"MICROSOFT_MCP_TOKENS_DIR": str(tmp_path)}
            with patch.dict(os.environ, env_patch, clear=False):
                # First call
                mock_app1 = _make_mock_app(email)
                mock_app1.acquire_token_by_device_flow.return_value["access_token"] = (
                    "token-1"
                )
                with patch.object(
                    MSALRefreshTokenAuth, "_get_msal_app", return_value=mock_app1
                ):
                    result1 = authenticate_new_account.fn(email=email)

                # Second call — returns a different token, should overwrite
                mock_app2 = _make_mock_app(email)
                mock_app2.acquire_token_by_device_flow.return_value["access_token"] = (
                    "token-2"
                )
                with patch.object(
                    MSALRefreshTokenAuth, "_get_msal_app", return_value=mock_app2
                ):
                    result2 = authenticate_new_account.fn(email=email)

            assert result1["status"] == "authenticated"
            assert result2["status"] == "authenticated"

            token_files = list(tmp_path.glob(f"{email}_access_token.json"))
            assert len(token_files) == 1, (
                "Only one token file should exist (overwritten)"
            )

            # Confirm file reflects the second call's token
            token_data = json.loads(token_files[0].read_text())
            assert token_data["access_token"] == "token-2"
        finally:
            tools_mod.auth_method = original_method
