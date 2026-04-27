"""
MSAL-based Authentication Module for Microsoft Graph MCP Server.

This module implements authentication using MSAL's device code flow with
file-based token storage. This is an alternative to the Azure SDK-based
authentication in auth.py, useful for:
- CLI/headless environments without browser access
- Servers and automated systems
- Compatibility with outlook-creds tokens

Authentication Flow:
1. Device code flow prompts user to visit URL and enter code
2. Access token and refresh token are saved to files
3. Subsequent calls use cached tokens, refreshing automatically when expired

Token Storage (compatible with outlook-creds format):
- {account}_access_token.json - Structured access token with metadata
- {account}_refresh_only.txt - Raw refresh token
- {account}_access_only.txt - Raw access token for easy extraction

Environment Variables:
- MICROSOFT_MCP_CLIENT_ID: Override default client ID
- MICROSOFT_MCP_TENANT_ID: Azure AD tenant (default: "common")
- MICROSOFT_MCP_TOKENS_DIR: Custom token storage directory
"""

import json
import logging
import os
import stat
import subprocess
import threading
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Optional

from msal import PublicClientApplication

logger = logging.getLogger(__name__)

# Microsoft Office client ID (same as outlook-creds)
# This client ID works out of the box for device code flow
MICROSOFT_OFFICE_CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
DEFAULT_TENANT_ID = "common"

# Token endpoint template
TOKEN_ENDPOINT_TEMPLATE = "https://login.microsoftonline.com/{tenant}/oauth2/v2.0/token"

# Default scopes for Microsoft Graph
DEFAULT_SCOPES = ["https://graph.microsoft.com/.default"]

# File permissions for sensitive data (owner read/write only)
TOKEN_FILE_MODE = stat.S_IRUSR | stat.S_IWUSR  # 0o600
CONFIG_DIR_MODE = stat.S_IRWXU  # 0o700

# Token expiry buffer (refresh if less than this many seconds remaining)
TOKEN_EXPIRY_BUFFER_SECONDS = 60


def _normalize_account_identifier(identifier: str) -> str:
    """Normalize account identifier to outlook-creds directory format."""
    return identifier.replace("@", "_").replace(".", "_")


def _load_outlook_creds_account_metadata(
    account_identifier: Optional[str],
) -> Optional[dict[str, str]]:
    """Load authority metadata from an outlook-creds account, if available.

    outlook-creds stores per-account metadata under:
    - $OUTLOOK_CREDS_CONFIG_DIR/tokens/<normalized>/account_info.json
    - ~/config/outlook-creds/tokens/<normalized>/account_info.json
    """
    if not account_identifier or account_identifier == "default":
        return None

    config_root = Path(
        os.getenv(
            "OUTLOOK_CREDS_CONFIG_DIR", str(Path.home() / "config" / "outlook-creds")
        )
    )
    account_info_path = (
        config_root
        / "tokens"
        / _normalize_account_identifier(account_identifier)
        / "account_info.json"
    )
    if not account_info_path.exists():
        return None

    try:
        account_info = json.loads(account_info_path.read_text())
        additional_properties = json.loads(
            account_info.get("additional_properties", "{}")
        )
    except Exception as e:
        logger.warning(
            f"Failed to load outlook-creds account metadata from {account_info_path}: {e}"
        )
        return None

    authority = account_info.get("authority")
    tenant_id = account_info.get("realm")
    client_id = additional_properties.get("aud")
    if not authority or not tenant_id:
        return None

    return {
        "authority": authority,
        "tenant_id": tenant_id,
        "client_id": client_id or MICROSOFT_OFFICE_CLIENT_ID,
    }


class MSALRefreshTokenAuth:
    """MSAL-based authentication with device code flow and file-based token storage.

    This class provides an alternative authentication method using MSAL's
    device code flow instead of Azure SDK's browser-based authentication.
    Tokens are stored in files, making it suitable for headless environments
    and compatible with the outlook-creds token format.

    Implements the AuthProvider protocol from auth_base.py.
    """

    def __init__(
        self,
        tokens_dir: Optional[Path] = None,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
        account_identifier: Optional[str] = None,
    ):
        """Initialize MSAL authentication.

        Args:
            tokens_dir: Directory for token storage. Defaults to
                ~/.config/microsoft-mcp/tokens/ or MICROSOFT_MCP_TOKENS_DIR env var.
            client_id: Azure AD client ID. Defaults to Microsoft Office client ID
                or MICROSOFT_MCP_CLIENT_ID env var.
            tenant_id: Azure AD tenant ID. Defaults to "common" or
                MICROSOFT_MCP_TENANT_ID env var.
            account_identifier: Identifier for token files (e.g., email address).
                If not provided, uses "default" as identifier.
        """
        # Token storage directory
        if tokens_dir:
            self.tokens_dir = Path(tokens_dir)
        else:
            default_dir = Path.home() / ".config" / "microsoft-mcp" / "tokens"
            self.tokens_dir = Path(
                os.getenv("MICROSOFT_MCP_TOKENS_DIR", str(default_dir))
            )

        # Client and tenant configuration
        explicit_client_id = client_id or os.getenv("MICROSOFT_MCP_CLIENT_ID")
        self.client_id = explicit_client_id or MICROSOFT_OFFICE_CLIENT_ID
        explicit_tenant_id = tenant_id or os.getenv("MICROSOFT_MCP_TENANT_ID")
        self.tenant_id = explicit_tenant_id or DEFAULT_TENANT_ID
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"

        # Prefer tenant-specific authority metadata from outlook-creds when the
        # caller selected an account but did not explicitly set the tenant.
        self.account_identifier = account_identifier or "default"
        if explicit_tenant_id is None:
            account_metadata = _load_outlook_creds_account_metadata(
                self.account_identifier
            )
            if account_metadata:
                self.tenant_id = account_metadata["tenant_id"]
                self.authority = account_metadata["authority"]
                if explicit_client_id is None:
                    self.client_id = account_metadata["client_id"]

        # MSAL app instance (lazy initialized)
        self._msal_app: Optional[PublicClientApplication] = None
        self._refresh_lock = threading.Lock()

        # Ensure token directory exists with secure permissions
        self._ensure_token_dir()

        logger.info(
            f"MSALRefreshTokenAuth initialized with client_id={self.client_id[:8]}..."
        )
        logger.info(f"MSAL authority: {self.authority}")
        logger.info(f"Token storage: {self.tokens_dir}")

    def _ensure_token_dir(self) -> None:
        """Create token directory with secure permissions if it doesn't exist."""
        self.tokens_dir.mkdir(parents=True, exist_ok=True)
        try:
            self.tokens_dir.chmod(CONFIG_DIR_MODE)
        except Exception as e:
            logger.warning(f"Could not set directory permissions: {e}")

    def _get_msal_app(self) -> PublicClientApplication:
        """Get or create MSAL PublicClientApplication instance."""
        if self._msal_app is None:
            self._msal_app = PublicClientApplication(
                client_id=self.client_id,
                authority=self.authority,
            )
            logger.info(f"Created MSAL app with authority: {self.authority}")
        return self._msal_app

    # Token file paths
    def _access_token_json_path(self) -> Path:
        """Path to structured access token JSON file."""
        return self.tokens_dir / f"{self.account_identifier}_access_token.json"

    def _refresh_token_path(self) -> Path:
        """Path to refresh token file."""
        return self.tokens_dir / f"{self.account_identifier}_refresh_only.txt"

    def _access_token_raw_path(self) -> Path:
        """Path to raw access token file."""
        return self.tokens_dir / f"{self.account_identifier}_access_only.txt"

    def _secure_write_file(self, path: Path, content: str) -> None:
        """Write file with secure permissions (0o600) set at create time."""
        fd = os.open(
            str(path),
            os.O_WRONLY | os.O_CREAT | os.O_TRUNC,
            TOKEN_FILE_MODE,
        )
        try:
            with os.fdopen(fd, "w") as f:
                f.write(content)
        except Exception:
            # If anything failed, ensure the fd is closed (fdopen takes ownership).
            # We don't unlink here — the file may already exist with valid content
            # from a previous successful call.
            raise
        # If the file PRE-EXISTED with a wider mode, O_CREAT mode is ignored by the OS.
        # Reset mode after-the-fact so overwrites tighten permissions.
        try:
            path.chmod(TOKEN_FILE_MODE)
        except Exception as e:
            logger.warning(f"Could not set file permissions for {path}: {e}")

    def _secure_write_json(self, path: Path, payload: dict[str, Any]) -> None:
        """Write JSON file with secure permissions (0o600) set at create time."""
        fd = os.open(
            str(path),
            os.O_WRONLY | os.O_CREAT | os.O_TRUNC,
            TOKEN_FILE_MODE,
        )
        with os.fdopen(fd, "w") as f:
            json.dump(payload, f, indent=2)
        # Tighten mode in case the file pre-existed with wider permissions.
        try:
            path.chmod(TOKEN_FILE_MODE)
        except Exception as e:
            logger.warning(f"Could not set file permissions for {path}: {e}")

    def _load_access_token_data(self) -> Optional[dict[str, Any]]:
        """Load structured access token data from JSON file."""
        path = self._access_token_json_path()
        if not path.exists():
            return None
        try:
            with open(path) as f:
                return json.load(f)
        except Exception as e:
            logger.warning(f"Failed to load access token data: {e}")
            return None

    def _load_refresh_token(self) -> Optional[str]:
        """Load refresh token from file."""
        path = self._refresh_token_path()
        if not path.exists():
            return None
        try:
            return path.read_text().strip()
        except Exception as e:
            logger.warning(f"Failed to load refresh token: {e}")
            return None

    def _save_tokens(
        self,
        access_token: str,
        refresh_token: Optional[str],
        expires_in: int,
        scopes: str,
        email: Optional[str] = None,
    ) -> None:
        """Save tokens to files in outlook-creds compatible format.

        Args:
            access_token: The access token string.
            refresh_token: The refresh token string (may be None).
            expires_in: Token lifetime in seconds.
            scopes: Space-separated scope string.
            email: Optional email address for metadata.
        """
        now = datetime.now(timezone.utc)
        expires_at = now + timedelta(seconds=expires_in)

        # 1. Save structured access token JSON
        access_token_data = {
            "email": email or self.account_identifier,
            "access_token": access_token,
            "token_type": "Bearer",
            "expires_in": expires_in,
            "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "refreshed_at": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "scopes": scopes,
            "api_type": "graph",
        }

        self._secure_write_json(self._access_token_json_path(), access_token_data)

        logger.info(f"Access token saved, expires at: {expires_at}")

        # 2. Save refresh token if provided
        if refresh_token:
            self._secure_write_file(self._refresh_token_path(), refresh_token)
            logger.info("Refresh token saved")

        # 3. Save raw access token for easy extraction
        self._secure_write_file(self._access_token_raw_path(), access_token)

    def _is_token_valid(self) -> bool:
        """Check if the cached access token is still valid.

        Returns:
            True if token exists and is not expired (with buffer), False otherwise.
        """
        token_data = self._load_access_token_data()
        if not token_data:
            return False

        expires_at_str = token_data.get("expires_at")
        if not expires_at_str:
            return False

        try:
            raw = expires_at_str.replace("Z", "+00:00")
            expires_at = datetime.fromisoformat(raw)
            if expires_at.tzinfo is None:
                expires_at = expires_at.replace(tzinfo=timezone.utc)
            now = datetime.now(timezone.utc)
            remaining = (expires_at - now).total_seconds()
            if remaining > TOKEN_EXPIRY_BUFFER_SECONDS:
                logger.info(f"Token valid for {remaining:.0f} more seconds")
                return True
            logger.info(f"Token expired or expiring soon ({remaining:.0f}s remaining)")
            return False
        except ValueError as e:
            logger.warning(f"Error parsing token expiration: {e}")
            return False

    def _refresh_access_token(self, refresh_token: str) -> dict[str, Any]:
        """Refresh access token using refresh token via HTTP POST.

        Adapted from outlook-creds/helpers/refresh_access_token.sh

        Args:
            refresh_token: Valid refresh token.

        Returns:
            Token response dictionary with access_token, refresh_token, etc.

        Raises:
            Exception: If token refresh fails.
        """
        logger.info("Refreshing access token using refresh token")

        token_endpoint = TOKEN_ENDPOINT_TEMPLATE.format(tenant=self.tenant_id)
        saved = self._load_access_token_data() or {}
        saved_scopes = saved.get("scopes") or ""
        if saved_scopes.strip():
            scopes = saved_scopes
            if "offline_access" not in scopes.split():
                scopes = f"{scopes} offline_access"
        else:
            scopes = "https://graph.microsoft.com/.default offline_access"

        # Prepare POST data
        data = urllib.parse.urlencode(
            {
                "client_id": self.client_id,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token,
                "scope": scopes,
            }
        ).encode("utf-8")

        headers = {
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        }

        try:
            req = urllib.request.Request(token_endpoint, data=data, headers=headers)
            with urllib.request.urlopen(req, timeout=30) as response:
                result = json.loads(response.read().decode("utf-8"))

                if "access_token" not in result:
                    raise RuntimeError("No access_token in refresh response")

                logger.info("Access token refreshed successfully")
                return result

        except urllib.error.HTTPError as e:
            error_body = e.read().decode("utf-8")
            logger.error(f"Token refresh failed (HTTP {e.code}): {error_body}")
            try:
                error_json = json.loads(error_body)
                error_desc = error_json.get(
                    "error_description", error_json.get("error", "Unknown")
                )
                raise RuntimeError(f"Token refresh failed: {error_desc}") from e
            except json.JSONDecodeError:
                raise RuntimeError(
                    f"Token refresh failed (HTTP {e.code}): {error_body[:200]}"
                ) from e

    def authenticate(self) -> dict[str, Any]:
        """Perform device code flow authentication.

        Initiates MSAL device code flow, which displays a URL and code
        for the user to enter in a browser. On macOS, automatically
        copies the code to clipboard and opens the browser.

        Returns:
            Authentication result dictionary with tokens and metadata.

        Raises:
            Exception: If authentication fails.
        """
        logger.info("Starting device code flow authentication")

        app = self._get_msal_app()

        # Check for cached accounts first
        if self.account_identifier != "default":
            accounts = app.get_accounts(username=self.account_identifier)
        else:
            accounts = app.get_accounts()
        if accounts:
            logger.info("Found cached account, attempting silent authentication")
            result = app.acquire_token_silent(
                scopes=DEFAULT_SCOPES, account=accounts[0]
            )
            if result and "access_token" in result:
                logger.info("Silent authentication successful")
                cached_username = accounts[0].get("username")
                self._save_tokens(
                    access_token=result["access_token"],
                    refresh_token=result.get("refresh_token"),
                    expires_in=result.get("expires_in", 3600),
                    scopes=result.get("scope", " ".join(DEFAULT_SCOPES)),
                    email=cached_username,
                )
                # Add username to result for easy access
                result["username"] = cached_username or self.account_identifier
                return result

        # Initiate device code flow
        flow = app.initiate_device_flow(scopes=DEFAULT_SCOPES)

        if "user_code" not in flow:
            error_msg = flow.get(
                "error_description", flow.get("error", "Unknown error")
            )
            raise RuntimeError(f"Failed to create device flow: {error_msg}")

        # Display authentication instructions
        print()
        print("=" * 70)
        print("  MICROSOFT AUTHENTICATION REQUIRED")
        print("=" * 70)
        print()
        print(flow["message"])
        print()
        print("=" * 70)

        # Auto-copy code to clipboard and open browser (macOS)
        user_code = flow.get("user_code", "")
        verification_uri = flow.get("verification_uri", "")

        if user_code:
            try:
                subprocess.run(
                    ["/usr/bin/pbcopy"],
                    input=user_code.encode(),
                    check=True,
                    capture_output=True,
                )
                print(f"\n  Code '{user_code}' copied to clipboard")
            except Exception:
                pass  # Silently ignore if pbcopy not available

        if verification_uri:
            try:
                subprocess.run(
                    ["/usr/bin/open", verification_uri],
                    check=True,
                    capture_output=True,
                )
                print(f"  Opened {verification_uri} in browser")
            except Exception:
                pass  # Silently ignore if open not available

        print()
        print("Waiting for authentication...")
        print()

        # Wait for user to complete authentication
        result = app.acquire_token_by_device_flow(flow)

        if "error" in result:
            error_desc = result.get("error_description", result.get("error", "Unknown"))
            raise RuntimeError(f"Authentication failed: {error_desc}")

        if "access_token" not in result:
            raise RuntimeError("No access_token in authentication response")

        # Extract email from ID token claims if available
        email = None
        id_token_claims = result.get("id_token_claims", {})
        email = id_token_claims.get("preferred_username") or id_token_claims.get(
            "email"
        )

        # Save tokens
        self._save_tokens(
            access_token=result["access_token"],
            refresh_token=result.get("refresh_token"),
            expires_in=result.get("expires_in", 3600),
            scopes=result.get("scope", " ".join(DEFAULT_SCOPES)),
            email=email,
        )

        logger.info("Authentication completed successfully")
        print("\n  Authentication successful!")

        # Add username to result for easy access
        result["username"] = email or self.account_identifier

        return result

    def _acquire_token_data(self) -> dict[str, Any]:
        """Single-point token acquisition for MSAL. Handles lock + refresh path.

        Returns the saved access-token-data dict (not just the token string)
        so the caller can extract either token or expires_at.
        """
        # Fast path: no lock needed if token is already valid.
        if self._is_token_valid():
            token_data = self._load_access_token_data()
            if token_data and token_data.get("access_token"):
                return token_data

        with self._refresh_lock:
            # Re-check inside the lock.
            if self._is_token_valid():
                token_data = self._load_access_token_data()
                if token_data and token_data.get("access_token"):
                    return token_data

            refresh_token = self._load_refresh_token()
            if not refresh_token:
                logger.error("No refresh token found. Authentication required.")
                raise RuntimeError(
                    "No refresh token found. Run authentication first: "
                    "MICROSOFT_MCP_AUTH_METHOD=msal uv run authenticate.py"
                )

            try:
                result = self._refresh_access_token(refresh_token)
                self._save_tokens(
                    access_token=result["access_token"],
                    refresh_token=result.get("refresh_token", refresh_token),
                    expires_in=result.get("expires_in", 3600),
                    scopes=result.get("scope", "https://graph.microsoft.com/.default"),
                )
                # Re-load the just-saved data so the caller sees canonical shape.
                data = self._load_access_token_data()
                if not data:
                    # Defensive — save should always produce a readable file.
                    raise RuntimeError("Token saved but could not be re-read")
                return data
            except Exception as e:
                logger.error(f"Token refresh failed: {e}")
                self.clear_cache()
                raise RuntimeError(
                    f"Token refresh failed: {e}. Please re-authenticate: "
                    "MICROSOFT_MCP_AUTH_METHOD=msal uv run authenticate.py"
                ) from e

    def get_token(self) -> str:
        """Get a valid access token, refreshing if needed."""
        logger.info("Getting access token")
        data = self._acquire_token_data()
        return data["access_token"]

    def get_token_with_details(self) -> tuple[str, int]:
        """Get access token with expiration timestamp."""
        data = self._acquire_token_data()
        token = data["access_token"]

        expires_at_str = data.get("expires_at")
        if expires_at_str:
            try:
                raw = expires_at_str.replace("Z", "+00:00")
                expires_at = datetime.fromisoformat(raw)
                if expires_at.tzinfo is None:
                    expires_at = expires_at.replace(tzinfo=timezone.utc)
                return token, int(expires_at.timestamp())
            except ValueError:
                pass

        # Fallback: assume 1 hour from now.
        expires_on = int((datetime.now(timezone.utc) + timedelta(hours=1)).timestamp())
        return token, expires_on

    def exists_valid_token(self) -> bool:
        """Check if a valid token exists or can be obtained.

        Returns:
            True if valid token exists, False otherwise.
        """
        # First check if access token is still valid
        if self._is_token_valid():
            return True

        # Check if we have a refresh token we could use
        refresh_token = self._load_refresh_token()
        return refresh_token is not None

    def clear_cache(self) -> None:
        """Clear all cached tokens and credentials."""
        logger.info("Clearing authentication cache")

        files_to_remove = [
            self._access_token_json_path(),
            self._refresh_token_path(),
            self._access_token_raw_path(),
        ]

        for path in files_to_remove:
            try:
                if path.exists():
                    path.unlink()
                    logger.info(f"Removed: {path}")
            except Exception as e:
                logger.warning(f"Failed to remove {path}: {e}")

        # Clear MSAL app instance
        self._msal_app = None

        logger.info("Authentication cache cleared")


def refresh_all_accounts(
    tokens_dir: Optional[Path] = None,
    client_id: Optional[str] = None,
    tenant_id: Optional[str] = None,
) -> list[dict[str, Any]]:
    """Refresh access tokens for every saved MSAL account.

    Iterates every {account_identifier}_access_token.json file in the tokens
    directory and runs the existing refresh-or-skip flow for each account in
    isolation. Failures on one account do NOT block the others. The global
    active account is not touched.

    Mirrors the semantics of ``outlook auth refresh`` in the outlook-creds
    repo: idempotent, safe to run repeatedly, skips refresh when token is
    still valid.

    Args:
        tokens_dir: Directory containing token files. Defaults to
            MICROSOFT_MCP_TOKENS_DIR or ~/.config/microsoft-mcp/tokens/.
        client_id: MSAL client ID. Defaults to env var or Microsoft Office
            public client ID.
        tenant_id: MSAL tenant ID. Defaults to env var or "common".

    Returns:
        A list of result dictionaries, one per account, each containing:
        - identifier (str): the account identifier (filename stem)
        - status (str): "valid" if the token was already valid (no refresh
          needed), "refreshed" if a network refresh occurred, "failed" if
          the refresh attempt errored
        - expires_at (str | None): ISO-format expiry timestamp after the
          operation, or None if the call failed
        - error (str | None): error message when status == "failed"
    """
    # Resolve tokens_dir the same way MSALRefreshTokenAuth.__init__ does.
    if tokens_dir is not None:
        resolved_dir = Path(tokens_dir)
    else:
        default_dir = Path.home() / ".config" / "microsoft-mcp" / "tokens"
        resolved_dir = Path(os.getenv("MICROSOFT_MCP_TOKENS_DIR", str(default_dir)))

    if not resolved_dir.exists():
        logger.info(
            f"refresh_all_accounts: tokens_dir {resolved_dir} does not exist, returning []"
        )
        return []

    token_files = sorted(resolved_dir.glob("*_access_token.json"))
    logger.info(
        f"refresh_all_accounts: found {len(token_files)} account(s) in {resolved_dir}"
    )

    results: list[dict[str, Any]] = []

    for token_file in token_files:
        # Strip the "_access_token" suffix from the stem to get the identifier.
        identifier = token_file.stem[: -len("_access_token")]
        logger.info(f"refresh_all_accounts: processing account '{identifier}'")

        probe = MSALRefreshTokenAuth(
            tokens_dir=resolved_dir,
            client_id=client_id,
            tenant_id=tenant_id,
            account_identifier=identifier,
        )

        if probe._is_token_valid():
            token_data = probe._load_access_token_data() or {}
            expires_at = token_data.get("expires_at")
            logger.info(
                f"refresh_all_accounts: '{identifier}' token is valid, expires_at={expires_at}"
            )
            results.append(
                {
                    "identifier": identifier,
                    "status": "valid",
                    "expires_at": expires_at,
                    "error": None,
                }
            )
            continue

        # Token is not valid — attempt refresh.
        try:
            token_data = probe._acquire_token_data()
            expires_at = token_data.get("expires_at")
            logger.info(
                f"refresh_all_accounts: '{identifier}' refreshed, expires_at={expires_at}"
            )
            results.append(
                {
                    "identifier": identifier,
                    "status": "refreshed",
                    "expires_at": expires_at,
                    "error": None,
                }
            )
        except Exception as e:
            # Read the pre-existing file for expires_at (may be stale or absent).
            stale_data = probe._load_access_token_data() or {}
            expires_at = stale_data.get("expires_at")
            logger.warning(f"refresh_all_accounts: '{identifier}' refresh failed: {e}")
            results.append(
                {
                    "identifier": identifier,
                    "status": "failed",
                    "expires_at": expires_at,
                    "error": str(e),
                }
            )

    return results
