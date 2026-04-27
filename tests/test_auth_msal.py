"""Tests for MSAL-based authentication module."""

import json
import os
import tempfile
from datetime import datetime, timedelta, timezone
from pathlib import Path
from unittest.mock import patch, MagicMock

import pytest

from src.microsoft_mcp.auth_msal import (
    MSALRefreshTokenAuth,
    MICROSOFT_OFFICE_CLIENT_ID,
    DEFAULT_TENANT_ID,
)
from src.microsoft_mcp.auth_base import AuthProvider


class TestMSALRefreshTokenAuthInit:
    """Tests for MSALRefreshTokenAuth initialization."""

    def test_init_default_values(self):
        """Test initialization with default values."""
        with tempfile.TemporaryDirectory() as tmpdir:
            with patch.dict(
                os.environ, {"MICROSOFT_MCP_TOKENS_DIR": tmpdir}, clear=False
            ):
                auth = MSALRefreshTokenAuth()

                assert auth.client_id == MICROSOFT_OFFICE_CLIENT_ID
                assert auth.tenant_id == DEFAULT_TENANT_ID
                assert auth.account_identifier == "default"
                assert (
                    auth.authority
                    == f"https://login.microsoftonline.com/{DEFAULT_TENANT_ID}"
                )
                assert auth._msal_app is None

    def test_init_custom_values(self):
        """Test initialization with custom values."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tokens_dir = Path(tmpdir)
            auth = MSALRefreshTokenAuth(
                tokens_dir=tokens_dir,
                client_id="custom-client-id",
                tenant_id="custom-tenant",
                account_identifier="user@example.com",
            )

            assert auth.tokens_dir == tokens_dir
            assert auth.client_id == "custom-client-id"
            assert auth.tenant_id == "custom-tenant"
            assert auth.account_identifier == "user@example.com"
            assert auth.authority == "https://login.microsoftonline.com/custom-tenant"

    def test_init_from_env_vars(self):
        """Test initialization from environment variables."""
        with tempfile.TemporaryDirectory() as tmpdir:
            env_vars = {
                "MICROSOFT_MCP_TOKENS_DIR": tmpdir,
                "MICROSOFT_MCP_CLIENT_ID": "env-client-id",
                "MICROSOFT_MCP_TENANT_ID": "env-tenant",
            }
            with patch.dict(os.environ, env_vars, clear=False):
                auth = MSALRefreshTokenAuth()

                assert str(auth.tokens_dir) == tmpdir
                assert auth.client_id == "env-client-id"
                assert auth.tenant_id == "env-tenant"

    def test_init_uses_outlook_creds_metadata_for_account_authority(self):
        """Test initialization prefers outlook-creds tenant metadata for known accounts."""
        with tempfile.TemporaryDirectory() as tmpdir:
            outlook_config_dir = Path(tmpdir) / "outlook-creds"
            account_dir = outlook_config_dir / "tokens" / "broach_cresa_com"
            account_dir.mkdir(parents=True)
            account_dir.joinpath("account_info.json").write_text(
                json.dumps(
                    {
                        "authority": "https://login.microsoftonline.com/tenant-123",
                        "realm": "tenant-123",
                        "login_name": "broach@cresa.com",
                        "additional_properties": json.dumps(
                            {"aud": "metadata-client-id"}
                        ),
                    }
                )
            )

            with patch.dict(
                os.environ,
                {
                    "OUTLOOK_CREDS_CONFIG_DIR": str(outlook_config_dir),
                },
                clear=True,
            ):
                auth = MSALRefreshTokenAuth(
                    tokens_dir=Path(tmpdir) / "tokens",
                    account_identifier="broach@cresa.com",
                )

            assert auth.tenant_id == "tenant-123"
            assert auth.authority == "https://login.microsoftonline.com/tenant-123"
            assert auth.client_id == "metadata-client-id"

    def test_init_prefers_explicit_tenant_over_outlook_creds_metadata(self):
        """Test explicit tenant configuration overrides outlook-creds metadata."""
        with tempfile.TemporaryDirectory() as tmpdir:
            outlook_config_dir = Path(tmpdir) / "outlook-creds"
            account_dir = outlook_config_dir / "tokens" / "broach_cresa_com"
            account_dir.mkdir(parents=True)
            account_dir.joinpath("account_info.json").write_text(
                json.dumps(
                    {
                        "authority": "https://login.microsoftonline.com/tenant-123",
                        "realm": "tenant-123",
                        "login_name": "broach@cresa.com",
                        "additional_properties": json.dumps(
                            {"aud": "metadata-client-id"}
                        ),
                    }
                )
            )

            with patch.dict(
                os.environ,
                {
                    "OUTLOOK_CREDS_CONFIG_DIR": str(outlook_config_dir),
                },
                clear=True,
            ):
                auth = MSALRefreshTokenAuth(
                    tokens_dir=Path(tmpdir) / "tokens",
                    tenant_id="explicit-tenant",
                    client_id="explicit-client",
                    account_identifier="broach@cresa.com",
                )

            assert auth.tenant_id == "explicit-tenant"
            assert auth.authority == "https://login.microsoftonline.com/explicit-tenant"
            assert auth.client_id == "explicit-client"

    def test_init_creates_tokens_directory(self):
        """Test that initialization creates the tokens directory."""
        with tempfile.TemporaryDirectory() as tmpdir:
            tokens_dir = Path(tmpdir) / "nested" / "tokens"
            MSALRefreshTokenAuth(tokens_dir=tokens_dir)

            assert tokens_dir.exists()
            assert tokens_dir.is_dir()


class TestMSALRefreshTokenAuthTokenFiles:
    """Tests for token file operations."""

    def test_access_token_json_path(self):
        """Test access token JSON path generation."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )
            expected = Path(tmpdir) / "test@example.com_access_token.json"
            assert auth._access_token_json_path() == expected

    def test_refresh_token_path(self):
        """Test refresh token path generation."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )
            expected = Path(tmpdir) / "test@example.com_refresh_only.txt"
            assert auth._refresh_token_path() == expected

    def test_access_token_raw_path(self):
        """Test raw access token path generation."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )
            expected = Path(tmpdir) / "test@example.com_access_only.txt"
            assert auth._access_token_raw_path() == expected

    def test_save_tokens(self):
        """Test saving tokens to files."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            auth._save_tokens(
                access_token="test-access-token",
                refresh_token="test-refresh-token",
                expires_in=3600,
                scopes="https://graph.microsoft.com/.default",
                email="test@example.com",
            )

            # Check access token JSON
            access_json_path = auth._access_token_json_path()
            assert access_json_path.exists()
            with open(access_json_path) as f:
                data = json.load(f)
            assert data["access_token"] == "test-access-token"
            assert data["email"] == "test@example.com"
            assert "expires_at" in data

            # Check refresh token
            refresh_path = auth._refresh_token_path()
            assert refresh_path.exists()
            assert refresh_path.read_text().strip() == "test-refresh-token"

            # Check raw access token
            raw_path = auth._access_token_raw_path()
            assert raw_path.exists()
            assert raw_path.read_text().strip() == "test-access-token"

    def test_load_access_token_data(self):
        """Test loading access token data from file."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Save test data
            test_data = {
                "access_token": "test-token",
                "expires_at": "2099-12-31T23:59:59Z",
            }
            with open(auth._access_token_json_path(), "w") as f:
                json.dump(test_data, f)

            # Load and verify
            loaded = auth._load_access_token_data()
            assert loaded is not None
            assert loaded["access_token"] == "test-token"

    def test_load_access_token_data_missing_file(self):
        """Test loading access token when file doesn't exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            loaded = auth._load_access_token_data()
            assert loaded is None

    def test_load_refresh_token(self):
        """Test loading refresh token from file."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Save test refresh token
            auth._refresh_token_path().write_text("test-refresh-token")

            # Load and verify
            loaded = auth._load_refresh_token()
            assert loaded == "test-refresh-token"

    def test_load_refresh_token_missing_file(self):
        """Test loading refresh token when file doesn't exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            loaded = auth._load_refresh_token()
            assert loaded is None


class TestMSALRefreshTokenAuthTokenValidation:
    """Tests for token validation logic."""

    def test_is_token_valid_with_valid_token(self):
        """Test token validation with a valid (non-expired) token."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Create token that expires in 1 hour
            expires_at = datetime.now(timezone.utc) + timedelta(hours=1)
            test_data = {
                "access_token": "test-token",
                "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
            }
            with open(auth._access_token_json_path(), "w") as f:
                json.dump(test_data, f)

            assert auth._is_token_valid() is True

    def test_is_token_valid_with_expired_token(self):
        """Test token validation with an expired token."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Create token that expired 1 hour ago
            expires_at = datetime.now(timezone.utc) - timedelta(hours=1)
            test_data = {
                "access_token": "test-token",
                "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
            }
            with open(auth._access_token_json_path(), "w") as f:
                json.dump(test_data, f)

            assert auth._is_token_valid() is False

    def test_is_token_valid_within_buffer(self):
        """Test token validation when within expiry buffer."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Create token that expires in 30 seconds (within buffer)
            expires_at = datetime.now(timezone.utc) + timedelta(seconds=30)
            test_data = {
                "access_token": "test-token",
                "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
            }
            with open(auth._access_token_json_path(), "w") as f:
                json.dump(test_data, f)

            # Should be invalid because within 60-second buffer
            assert auth._is_token_valid() is False

    def test_is_token_valid_no_file(self):
        """Test token validation when no token file exists."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            assert auth._is_token_valid() is False


class TestMSALRefreshTokenAuthGetToken:
    """Tests for get_token method."""

    def test_get_token_with_valid_cached_token(self):
        """Test get_token returns cached token when valid."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Create valid cached token
            expires_at = datetime.now(timezone.utc) + timedelta(hours=1)
            test_data = {
                "access_token": "cached-token",
                "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
            }
            with open(auth._access_token_json_path(), "w") as f:
                json.dump(test_data, f)

            token = auth.get_token()
            assert token == "cached-token"

    def test_get_token_refreshes_expired_token(self):
        """Test get_token refreshes when token is expired."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Create expired token
            expires_at = datetime.now(timezone.utc) - timedelta(hours=1)
            test_data = {
                "access_token": "expired-token",
                "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
            }
            with open(auth._access_token_json_path(), "w") as f:
                json.dump(test_data, f)

            # Create refresh token
            auth._refresh_token_path().write_text("test-refresh-token")

            # Mock the refresh response
            mock_response = {
                "access_token": "new-access-token",
                "refresh_token": "new-refresh-token",
                "expires_in": 3600,
                "scope": "https://graph.microsoft.com/.default",
            }

            with patch.object(
                auth, "_refresh_access_token", return_value=mock_response
            ):
                token = auth.get_token()
                assert token == "new-access-token"

    def test_get_token_raises_without_refresh_token(self):
        """Test get_token raises error when no refresh token available."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # No token files exist
            with pytest.raises(Exception) as exc_info:
                auth.get_token()

            assert "No refresh token found" in str(exc_info.value)


class TestMSALRefreshTokenAuthExistsValidToken:
    """Tests for exists_valid_token method."""

    def test_exists_valid_token_with_valid_token(self):
        """Test exists_valid_token returns True with valid cached token."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Create valid cached token
            expires_at = datetime.now(timezone.utc) + timedelta(hours=1)
            test_data = {
                "access_token": "valid-token",
                "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
            }
            with open(auth._access_token_json_path(), "w") as f:
                json.dump(test_data, f)

            assert auth.exists_valid_token() is True

    def test_exists_valid_token_with_refresh_token_only(self):
        """Test exists_valid_token returns True when refresh token available."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Only refresh token exists
            auth._refresh_token_path().write_text("test-refresh-token")

            assert auth.exists_valid_token() is True

    def test_exists_valid_token_with_no_tokens(self):
        """Test exists_valid_token returns False when no tokens available."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            assert auth.exists_valid_token() is False


class TestMSALRefreshTokenAuthClearCache:
    """Tests for clear_cache method."""

    def test_clear_cache_removes_all_files(self):
        """Test clear_cache removes all token files."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            # Create token files
            auth._access_token_json_path().write_text("{}")
            auth._refresh_token_path().write_text("refresh")
            auth._access_token_raw_path().write_text("access")

            # Verify files exist
            assert auth._access_token_json_path().exists()
            assert auth._refresh_token_path().exists()
            assert auth._access_token_raw_path().exists()

            # Clear cache
            auth.clear_cache()

            # Verify files removed
            assert not auth._access_token_json_path().exists()
            assert not auth._refresh_token_path().exists()
            assert not auth._access_token_raw_path().exists()


class TestMSALRefreshTokenAuthProtocolCompliance:
    """Tests for AuthProvider protocol compliance."""

    def test_implements_auth_provider_protocol(self):
        """Test that MSALRefreshTokenAuth implements AuthProvider protocol."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(tokens_dir=Path(tmpdir))

            # Check protocol compliance
            assert isinstance(auth, AuthProvider)

    def test_has_required_methods(self):
        """Test that all required protocol methods exist."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(tokens_dir=Path(tmpdir))

            # Verify all required methods exist
            assert hasattr(auth, "get_token")
            assert hasattr(auth, "get_token_with_details")
            assert hasattr(auth, "exists_valid_token")
            assert hasattr(auth, "authenticate")
            assert hasattr(auth, "clear_cache")

            # Verify methods are callable
            assert callable(auth.get_token)
            assert callable(auth.get_token_with_details)
            assert callable(auth.exists_valid_token)
            assert callable(auth.authenticate)
            assert callable(auth.clear_cache)


# ---------------------------------------------------------------------------
# Scope sanitization tests (Bug B — AADSTS70011 fix)
# ---------------------------------------------------------------------------


def _make_fake_urlopen_response(token_data: dict) -> MagicMock:
    """Return a mock context-manager that urlopen() can return."""
    import json as _json

    response_body = _json.dumps(token_data).encode("utf-8")
    mock_response = MagicMock()
    mock_response.read.return_value = response_body
    mock_response.__enter__ = lambda s: s
    mock_response.__exit__ = MagicMock(return_value=False)
    return mock_response


def _capture_scope_from_urlopen_call(mock_urlopen: MagicMock) -> str:
    """Decode the POST body from the captured urlopen() call and extract scope."""
    import urllib.parse as _up

    req = mock_urlopen.call_args[0][0]  # first positional arg = Request object
    body = req.data.decode("utf-8")
    params = dict(_up.parse_qsl(body))
    return params["scope"]


class TestRefreshAccessTokenScopeSanitization:
    """_refresh_access_token must sanitize mixed .default + specific scopes."""

    _FAKE_TOKEN_RESPONSE = {
        "access_token": "new-access-token",
        "refresh_token": "new-refresh-token",
        "expires_in": 3600,
        "scope": "https://graph.microsoft.com/.default offline_access",
    }

    def _make_auth(self, tmp_path: Path, saved_scopes: str) -> MSALRefreshTokenAuth:
        auth = MSALRefreshTokenAuth(
            tokens_dir=tmp_path,
            account_identifier="test@example.com",
        )
        # Write a saved access-token JSON with the desired scope string.
        data = {
            "access_token": "old-token",
            "expires_at": "2020-01-01T00:00:00Z",
            "scopes": saved_scopes,
        }
        auth._access_token_json_path().write_text(json.dumps(data))
        return auth

    def test_refresh_drops_default_when_specific_scopes_present(self, tmp_path):
        """When saved scopes mix .default + specific Graph scopes, .default
        must be dropped from the POST body (AADSTS70011 prevention)."""
        mixed_scopes = (
            "https://graph.microsoft.com/.default "
            "https://graph.microsoft.com/Mail.ReadWrite "
            "offline_access"
        )
        auth = self._make_auth(tmp_path, mixed_scopes)

        with patch(
            "urllib.request.urlopen",
            return_value=_make_fake_urlopen_response(self._FAKE_TOKEN_RESPONSE),
        ) as mock_urlopen:
            auth._refresh_access_token("fake-refresh-token")

        scope_sent = _capture_scope_from_urlopen_call(mock_urlopen)
        scope_parts = scope_sent.split()

        assert not any(p.endswith("/.default") for p in scope_parts), (
            f".default must be absent when specific scopes are present, got: {scope_sent!r}"
        )
        assert "https://graph.microsoft.com/Mail.ReadWrite" in scope_parts
        assert "offline_access" in scope_parts

    def test_refresh_preserves_default_when_only_default_present(self, tmp_path):
        """When saved scopes contain only .default (no specific scopes),
        .default must be kept in the POST body."""
        only_default = "https://graph.microsoft.com/.default offline_access"
        auth = self._make_auth(tmp_path, only_default)

        with patch(
            "urllib.request.urlopen",
            return_value=_make_fake_urlopen_response(self._FAKE_TOKEN_RESPONSE),
        ) as mock_urlopen:
            auth._refresh_access_token("fake-refresh-token")

        scope_sent = _capture_scope_from_urlopen_call(mock_urlopen)
        assert "https://graph.microsoft.com/.default" in scope_sent.split(), (
            f".default must be preserved when no specific scopes present, got: {scope_sent!r}"
        )

    def test_refresh_preserves_specific_scopes_when_no_default_present(self, tmp_path):
        """When saved scopes contain only specific scopes (no .default),
        the scope string must be sent unchanged (plus offline_access)."""
        only_specific = (
            "https://graph.microsoft.com/Mail.ReadWrite "
            "https://graph.microsoft.com/Calendars.ReadWrite "
            "offline_access"
        )
        auth = self._make_auth(tmp_path, only_specific)

        with patch(
            "urllib.request.urlopen",
            return_value=_make_fake_urlopen_response(self._FAKE_TOKEN_RESPONSE),
        ) as mock_urlopen:
            auth._refresh_access_token("fake-refresh-token")

        scope_sent = _capture_scope_from_urlopen_call(mock_urlopen)
        scope_parts = scope_sent.split()

        assert not any(p.endswith("/.default") for p in scope_parts), (
            f".default must not appear when only specific scopes were saved, got: {scope_sent!r}"
        )
        assert "https://graph.microsoft.com/Mail.ReadWrite" in scope_parts
        assert "https://graph.microsoft.com/Calendars.ReadWrite" in scope_parts
        assert "offline_access" in scope_parts


class TestMSALRefreshTokenAuthDeviceCodeFlow:
    """Tests for device code flow authentication."""

    def test_authenticate_calls_msal(self):
        """Test authenticate method calls MSAL properly."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            mock_app = MagicMock()
            mock_app.get_accounts.return_value = []
            mock_app.initiate_device_flow.return_value = {
                "user_code": "ABC123",
                "verification_uri": "https://microsoft.com/devicelogin",
                "message": "Enter code ABC123",
            }
            mock_app.acquire_token_by_device_flow.return_value = {
                "access_token": "new-token",
                "refresh_token": "new-refresh",
                "expires_in": 3600,
                "scope": "https://graph.microsoft.com/.default",
                "id_token_claims": {"preferred_username": "test@example.com"},
            }

            with patch.object(auth, "_get_msal_app", return_value=mock_app):
                # Patch subprocess to avoid actual clipboard/browser operations
                with patch("subprocess.run"):
                    result = auth.authenticate()

            assert result["access_token"] == "new-token"
            assert result["username"] == "test@example.com"
            mock_app.initiate_device_flow.assert_called_once()
            mock_app.acquire_token_by_device_flow.assert_called_once()

    def test_authenticate_uses_cached_account(self):
        """Test authenticate uses cached account for silent auth."""
        with tempfile.TemporaryDirectory() as tmpdir:
            auth = MSALRefreshTokenAuth(
                tokens_dir=Path(tmpdir), account_identifier="test@example.com"
            )

            mock_app = MagicMock()
            mock_account = {"username": "test@example.com"}
            mock_app.get_accounts.return_value = [mock_account]
            mock_app.acquire_token_silent.return_value = {
                "access_token": "silent-token",
                "refresh_token": "silent-refresh",
                "expires_in": 3600,
                "scope": "https://graph.microsoft.com/.default",
            }

            with patch.object(auth, "_get_msal_app", return_value=mock_app):
                result = auth.authenticate()

            assert result["access_token"] == "silent-token"
            assert result["username"] == "test@example.com"
            mock_app.get_accounts.assert_called_once_with(username="test@example.com")
            mock_app.acquire_token_silent.assert_called_once()
            # Device flow should not be initiated
            mock_app.initiate_device_flow.assert_not_called()


import inspect
from microsoft_mcp.auth_msal import MSALRefreshTokenAuth


def test_init_assigns_account_identifier_exactly_once():
    source = inspect.getsource(MSALRefreshTokenAuth.__init__)
    occurrences = source.count("self.account_identifier =")
    assert occurrences == 1, f"expected 1 assignment, found {occurrences}"


def test_init_uses_default_when_no_identifier_given(tmp_path):
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, client_id="test-cid")
    assert auth.account_identifier == "default"


def test_init_preserves_explicit_identifier(tmp_path):
    auth = MSALRefreshTokenAuth(
        tokens_dir=tmp_path, client_id="test-cid", account_identifier="user@example.com"
    )
    assert auth.account_identifier == "user@example.com"


import io
import json as _json
from urllib.parse import parse_qs


def test_msal_refresh_preserves_saved_scopes(tmp_path, monkeypatch):
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    (tmp_path / "x@y.com_refresh_only.txt").write_text("rt-123")
    (tmp_path / "x@y.com_access_token.json").write_text(
        '{"email": "x@y.com", "access_token": "old", "token_type": "Bearer", '
        '"expires_in": 0, "expires_at": "2020-01-01T00:00:00Z", '
        '"refreshed_at": "2020-01-01T00:00:00Z", '
        '"scopes": "Mail.Read Files.Read offline_access", "api_type": "graph"}'
    )

    captured = {}

    def fake_urlopen(req, timeout=30):
        body = req.data.decode()
        params = parse_qs(body)
        captured["scope"] = params["scope"][0]
        return io.BytesIO(
            _json.dumps(
                {
                    "access_token": "new",
                    "refresh_token": "rt-123",
                    "expires_in": 3600,
                    "scope": params["scope"][0],
                }
            ).encode()
        )

    monkeypatch.setattr(
        "microsoft_mcp.auth_msal.urllib.request.urlopen",
        fake_urlopen,
    )

    auth.get_token()
    assert "Mail.Read" in captured["scope"]
    assert "Files.Read" in captured["scope"]
    assert "offline_access" in captured["scope"]


def test_msal_refresh_falls_back_to_default_when_no_scopes_saved(tmp_path, monkeypatch):
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    (tmp_path / "x@y.com_refresh_only.txt").write_text("rt-none")
    # access_token.json exists but has no scopes field
    (tmp_path / "x@y.com_access_token.json").write_text(
        '{"email": "x@y.com", "access_token": "old", "token_type": "Bearer", '
        '"expires_in": 0, "expires_at": "2020-01-01T00:00:00Z", '
        '"refreshed_at": "2020-01-01T00:00:00Z", '
        '"api_type": "graph"}'
    )

    captured = {}

    def fake_urlopen(req, timeout=30):
        body = req.data.decode()
        params = parse_qs(body)
        captured["scope"] = params["scope"][0]
        return io.BytesIO(
            _json.dumps(
                {
                    "access_token": "new",
                    "refresh_token": "rt-none",
                    "expires_in": 3600,
                    "scope": params["scope"][0],
                }
            ).encode()
        )

    monkeypatch.setattr(
        "microsoft_mcp.auth_msal.urllib.request.urlopen",
        fake_urlopen,
    )

    auth.get_token()
    # Falls back to .default offline_access when no saved scopes.
    assert ".default" in captured["scope"]
    assert "offline_access" in captured["scope"]


def test_msal_refresh_appends_offline_access_if_missing(tmp_path, monkeypatch):
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    (tmp_path / "x@y.com_refresh_only.txt").write_text("rt-xyz")
    # saved scopes without offline_access
    (tmp_path / "x@y.com_access_token.json").write_text(
        '{"email": "x@y.com", "access_token": "old", "token_type": "Bearer", '
        '"expires_in": 0, "expires_at": "2020-01-01T00:00:00Z", '
        '"refreshed_at": "2020-01-01T00:00:00Z", '
        '"scopes": "Mail.Read", "api_type": "graph"}'
    )

    captured = {}

    def fake_urlopen(req, timeout=30):
        body = req.data.decode()
        params = parse_qs(body)
        captured["scope"] = params["scope"][0]
        return io.BytesIO(
            _json.dumps(
                {
                    "access_token": "new",
                    "refresh_token": "rt-xyz",
                    "expires_in": 3600,
                    "scope": params["scope"][0],
                }
            ).encode()
        )

    monkeypatch.setattr(
        "microsoft_mcp.auth_msal.urllib.request.urlopen",
        fake_urlopen,
    )

    auth.get_token()
    # offline_access must be appended so refresh tokens keep working.
    assert "Mail.Read" in captured["scope"]
    assert "offline_access" in captured["scope"]


def test_is_token_valid_accepts_microseconds_format(tmp_path):
    from datetime import datetime, timedelta, timezone
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    future = datetime.now(timezone.utc) + timedelta(hours=1)
    # microseconds variant
    data = {
        "access_token": "tok",
        "expires_at": future.strftime("%Y-%m-%dT%H:%M:%S.%fZ"),
    }
    path = tmp_path / "x@y.com_access_token.json"
    path.write_text(_json.dumps(data))
    assert auth._is_token_valid()


def test_is_token_valid_accepts_offset_format(tmp_path):
    from datetime import datetime, timedelta, timezone
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    future = datetime.now(timezone.utc) + timedelta(hours=1)
    # +00:00 offset variant
    data = {
        "access_token": "tok",
        "expires_at": future.strftime("%Y-%m-%dT%H:%M:%S+00:00"),
    }
    path = tmp_path / "x@y.com_access_token.json"
    path.write_text(_json.dumps(data))
    assert auth._is_token_valid()


def test_is_token_valid_accepts_legacy_z_format(tmp_path):
    """The historical format we write must continue to parse."""
    from datetime import datetime, timedelta, timezone
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    future = datetime.now(timezone.utc) + timedelta(hours=1)
    data = {
        "access_token": "tok",
        "expires_at": future.strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    path = tmp_path / "x@y.com_access_token.json"
    path.write_text(_json.dumps(data))
    assert auth._is_token_valid()


def test_is_token_valid_returns_false_for_unparseable(tmp_path):
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    data = {"access_token": "tok", "expires_at": "not-a-datetime"}
    path = tmp_path / "x@y.com_access_token.json"
    path.write_text(_json.dumps(data))
    assert auth._is_token_valid() is False


import threading


def test_concurrent_get_token_refreshes_exactly_once(tmp_path, monkeypatch):
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    (tmp_path / "x@y.com_refresh_only.txt").write_text("rt-1")
    past = datetime.now(timezone.utc) - timedelta(minutes=5)
    (tmp_path / "x@y.com_access_token.json").write_text(
        _json.dumps(
            {
                "access_token": "stale",
                "expires_at": past.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "scopes": "Mail.Read offline_access",
            }
        )
    )

    calls = {"n": 0}
    inner_lock = threading.Lock()

    def fake_refresh(refresh_token):
        with inner_lock:
            calls["n"] += 1
        return {
            "access_token": "fresh",
            "refresh_token": refresh_token,
            "expires_in": 3600,
            "scope": "Mail.Read offline_access",
        }

    monkeypatch.setattr(auth, "_refresh_access_token", fake_refresh)

    tokens: list[str] = []

    def worker():
        tokens.append(auth.get_token())

    threads = [threading.Thread(target=worker) for _ in range(8)]
    for t in threads:
        t.start()
    for t in threads:
        t.join()

    assert calls["n"] == 1, f"expected 1 refresh, got {calls['n']}"
    assert all(t == "fresh" for t in tokens)
    assert len(tokens) == 8


def test_serial_get_token_calls_still_work(tmp_path, monkeypatch):
    """Single-threaded use path must remain unaffected."""
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    (tmp_path / "x@y.com_refresh_only.txt").write_text("rt-serial")
    past = datetime.now(timezone.utc) - timedelta(minutes=5)
    (tmp_path / "x@y.com_access_token.json").write_text(
        _json.dumps(
            {
                "access_token": "stale",
                "expires_at": past.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "scopes": "Mail.Read offline_access",
            }
        )
    )

    calls = {"n": 0}

    def fake_refresh(refresh_token):
        calls["n"] += 1
        return {
            "access_token": f"fresh-{calls['n']}",
            "refresh_token": refresh_token,
            "expires_in": 3600,
            "scope": "Mail.Read offline_access",
        }

    monkeypatch.setattr(auth, "_refresh_access_token", fake_refresh)

    first = auth.get_token()
    # Second call should find the fresh token valid (its expires_at was just written)
    # and NOT trigger another refresh.
    second = auth.get_token()

    assert calls["n"] == 1
    assert first == "fresh-1"
    assert second == "fresh-1"


def test_account_identifier_immutable_after_save_with_default(tmp_path):
    """Default-initialized instance must NOT rewrite account_identifier during save."""
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path)
    assert auth.account_identifier == "default"

    auth._save_tokens(
        access_token="t",
        refresh_token="r",
        expires_in=3600,
        scopes="Mail.Read offline_access",
        email="new@example.com",
    )

    # Identifier must remain "default" — path stability for any caller that cached a path.
    assert auth.account_identifier == "default"


def test_account_identifier_immutable_after_save_with_explicit(tmp_path):
    """Explicit identifier must survive a save that discovers a different email."""
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(
        tokens_dir=tmp_path, account_identifier="explicit@x.com"
    )
    assert auth.account_identifier == "explicit@x.com"

    auth._save_tokens(
        access_token="t",
        refresh_token="r",
        expires_in=3600,
        scopes="Mail.Read offline_access",
        email="different@y.com",
    )

    assert auth.account_identifier == "explicit@x.com"


def test_save_preserves_email_in_payload(tmp_path):
    """The email is still recorded in access_token.json even though the identifier doesn't change."""
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path)
    auth._save_tokens(
        access_token="t",
        refresh_token="r",
        expires_in=3600,
        scopes="Mail.Read offline_access",
        email="payload@example.com",
    )

    payload = _json.loads((tmp_path / "default_access_token.json").read_text())
    assert payload["email"] == "payload@example.com"


def test_msal_get_token_and_get_token_with_details_share_value(tmp_path, monkeypatch):
    """After refactor, MSAL's get_token and get_token_with_details return consistent values."""
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    # Write a fresh token so the happy path hits cached state.
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    future = datetime.now(timezone.utc) + timedelta(hours=1)
    (tmp_path / "x@y.com_access_token.json").write_text(
        _json.dumps(
            {
                "access_token": "cached-tok",
                "expires_at": future.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "scopes": "Mail.Read offline_access",
            }
        )
    )

    plain = auth.get_token()
    detailed, exp = auth.get_token_with_details()
    assert plain == detailed == "cached-tok"
    # expires_on is a unix timestamp; must be within a second of the stored expires_at.
    assert abs(exp - int(future.timestamp())) <= 1


def test_secure_write_file_creates_with_owner_only_mode(tmp_path):
    """Token file must have 0o600 mode immediately on creation, not after a chmod race."""
    import stat
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    target = tmp_path / "secure-test.txt"
    auth._secure_write_file(target, "secret-content")

    assert target.exists()
    assert target.read_text() == "secret-content"
    actual_mode = stat.S_IMODE(target.stat().st_mode)
    assert actual_mode == 0o600, f"expected 0o600, got 0o{actual_mode:o}"


def test_secure_write_file_overwrites_existing_with_correct_mode(tmp_path):
    """When the file already exists with a wider mode, secure_write must reset it to 0o600."""
    import stat
    from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    target = tmp_path / "preexisting.txt"
    target.write_text("old")
    target.chmod(0o644)  # simulate a world-readable existing file

    auth._secure_write_file(target, "new-secret")
    assert target.read_text() == "new-secret"
    actual_mode = stat.S_IMODE(target.stat().st_mode)
    assert actual_mode == 0o600, f"expected 0o600, got 0o{actual_mode:o}"
