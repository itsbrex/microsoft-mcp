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
