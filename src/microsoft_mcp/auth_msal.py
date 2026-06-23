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

import base64
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

GRAPH_SCOPE = "https://graph.microsoft.com/.default offline_access"
OUTLOOK_SCOPE = "https://outlook.office365.com/.default offline_access"

# File permissions for sensitive data (owner read/write only)
TOKEN_FILE_MODE = stat.S_IRUSR | stat.S_IWUSR  # 0o600
CONFIG_DIR_MODE = stat.S_IRWXU  # 0o700

# Token expiry buffer (refresh if less than this many seconds remaining)
TOKEN_EXPIRY_BUFFER_SECONDS = 60


def _interactive_auth_allowed() -> bool:
    """Whether a silent-refresh failure may fall through to interactive auth.

    The device-code flow prints a code to stderr and blocks until the user
    completes it in a browser. That is fine for an interactive server (the
    documented ``authenticate_new_account`` flow), but in a fully headless
    deployment (cron, CI, a detached service nobody is watching) the fallback
    would hang forever on a refresh failure.

    Opt out by setting ``MICROSOFT_MCP_NONINTERACTIVE`` to a truthy value
    (``1``/``true``/``yes``/``on``); callers then get a clear, actionable
    error instead of a hang. This is **off by default** so existing
    interactive behavior is unchanged. Note: stdin is intentionally NOT used
    as a signal — the MCP server runs over a stdio pipe (never a TTY), yet
    its stderr device-code flow works, so a TTY check would break it.
    """
    val = os.getenv("MICROSOFT_MCP_NONINTERACTIVE", "").strip().lower()
    return val not in ("1", "true", "yes", "on")


def _require_interactive_or_raise(identifier: str) -> None:
    """Raise a clear error when interactive auth is disabled (headless guard)."""
    if not _interactive_auth_allowed():
        raise RuntimeError(
            "Token refresh failed and interactive authentication is disabled "
            "(MICROSOFT_MCP_NONINTERACTIVE is set). Re-authenticate out of band, "
            f"e.g.: microsoft-mcp auth refresh {identifier} --force"
        )


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
        api_type: str = "graph",
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

        # API resource this instance targets ("graph" or "outlook"). Only the
        # access-token files and refresh scope vary; the refresh token is shared.
        self.api_type = api_type if api_type in ("graph", "outlook") else "graph"

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

    def _default_scope(self) -> str:
        return OUTLOOK_SCOPE if self.api_type == "outlook" else GRAPH_SCOPE

    # Token file paths
    def _access_token_json_path(self) -> Path:
        """Path to structured access token JSON file."""
        suffix = (
            "_outlook_access_token" if self.api_type == "outlook" else "_access_token"
        )
        return self.tokens_dir / f"{self.account_identifier}{suffix}.json"

    def _refresh_token_path(self) -> Path:
        """Path to refresh token file."""
        return self.tokens_dir / f"{self.account_identifier}_refresh_only.txt"

    def _access_token_raw_path(self) -> Path:
        """Path to raw access token file."""
        suffix = (
            "_outlook_access_only" if self.api_type == "outlook" else "_access_only"
        )
        return self.tokens_dir / f"{self.account_identifier}{suffix}.txt"

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
            "api_type": self.api_type,
        }

        self._secure_write_json(self._access_token_json_path(), access_token_data)

        logger.info(f"Access token saved, expires at: {expires_at}")

        # 2. Save refresh token if provided — but ONLY from the Graph leg.
        # The refresh token is shared across api_types ({id}_refresh_only.txt),
        # and Azure rotates it on every refresh. An Outlook refresh response
        # carries a refresh token scoped to the Outlook grant; persisting it
        # would clobber the Graph-consented token and make the next Graph
        # `.default` refresh fail with AADSTS65002 (first-party preauthorization).
        # outlook-creds enforces the same rule (get_oauth_tokens.py: only the
        # graph response updates the shared refresh token).
        if refresh_token and self.api_type == "graph":
            self._secure_write_file(self._refresh_token_path(), refresh_token)
            logger.info("Refresh token saved")
        elif refresh_token:
            logger.info(
                "Skipping refresh-token persist for api_type=%s (shared refresh "
                "token stays Graph-consented)",
                self.api_type,
            )

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

        # Sanitize: Azure AD rejects refresh requests that mix `<resource>/.default`
        # with resource-specific scopes (AADSTS70011). If the saved scope string
        # contains both, drop the `.default` and keep the specific scopes — the
        # narrower request still satisfies the existing consent grant.
        parts = saved_scopes.split() if saved_scopes.strip() else []
        has_default = any(p.endswith("/.default") for p in parts)
        has_specific = any(
            p.startswith("https://graph.microsoft.com/") and not p.endswith("/.default")
            for p in parts
        )
        if has_default and has_specific:
            parts = [p for p in parts if not p.endswith("/.default")]
            logger.info(
                "Sanitized scope: dropped .default to avoid AADSTS70011 (mix with "
                "resource-specific scopes)"
            )

        if parts:
            if "offline_access" not in parts:
                parts.append("offline_access")
            scopes = " ".join(parts)
        else:
            scopes = self._default_scope()

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
                    expires_in=int(result.get("expires_in", 3600)),
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

    def _do_refresh_locked(self) -> dict[str, Any]:
        """Perform one refresh-token cycle: load refresh token, call Azure
        AD's token endpoint, persist new tokens, return canonical token data.

        Acquires self._refresh_lock for the entire critical section. Does
        NOT clear cache on failure — caller decides whether eviction is
        appropriate. Raises on any failure (no refresh token on disk,
        network error, Azure AD rejection).

        Returns:
            Canonical token data dict (re-loaded from disk after save).

        Raises:
            RuntimeError: if no refresh token is on disk.
            Exception: if the refresh network call or save fails.
        """
        with self._refresh_lock:
            # Re-check validity under the lock: another thread may have already
            # refreshed while we were waiting to acquire it.
            if self._is_token_valid():
                token_data = self._load_access_token_data()
                if token_data and token_data.get("access_token"):
                    return token_data

            refresh_token = self._load_refresh_token()
            if not refresh_token:
                raise RuntimeError(
                    "No refresh token available; re-authentication required"
                )
            result = self._refresh_access_token(refresh_token)
            self._save_tokens(
                access_token=result["access_token"],
                refresh_token=result.get("refresh_token", refresh_token),
                expires_in=result.get("expires_in", 3600),
                scopes=result.get("scope", self._default_scope()),
            )
            # Re-load the just-saved data so the caller sees canonical shape.
            data = self._load_access_token_data()
            if not data:
                # Defensive — save should always produce a readable file.
                raise RuntimeError("Token saved but could not be re-read")
            return data

    def _acquire_token_data(self) -> dict[str, Any]:
        """Single-point token acquisition for MSAL. Handles lock + refresh path.

        Returns the saved access-token-data dict (not just the token string)
        so the caller can extract either token or expires_at.

        Falls back to interactive device code flow when no refresh token is
        available or when the refresh token itself is expired/revoked, making
        re-authentication invisible to the user as long as a browser is reachable.
        """
        # Fast path: no lock needed if token is already valid.
        if self._is_token_valid():
            token_data = self._load_access_token_data()
            if token_data and token_data.get("access_token"):
                return token_data

        if self._load_refresh_token():
            try:
                return self._do_refresh_locked()
            except Exception as e:
                logger.warning(
                    f"Silent token refresh failed ({e}); falling back to interactive auth"
                )
                self.clear_cache()
        else:
            logger.warning(
                "No refresh token found; initiating interactive authentication"
            )

        # Silent refresh unavailable — prompt the user via device code flow,
        # unless interactive auth is disabled (headless guard).
        _require_interactive_or_raise(self.account_identifier)
        self.authenticate()
        token_data = self._load_access_token_data()
        if not token_data or not token_data.get("access_token"):
            raise RuntimeError(
                "Interactive authentication completed but token data could not be loaded"
            )
        return token_data

    def force_refresh(self) -> None:
        """Force a token refresh even if the cached access token looks valid.

        Used by graph.request when a 401 comes back from Microsoft Graph
        (e.g., after a clock-skew miss, password change, or admin-revoked
        consent). Delegates to _do_refresh_locked(), which acquires the
        refresh lock, loads the refresh token from disk, calls
        _refresh_access_token, and persists the new tokens.

        Falls back to interactive device code flow when the refresh token is
        itself expired or revoked, so the 401-retry path in graph.request can
        succeed without the user having to run a separate script.
        """
        logger.info(f"Force-refreshing access token for {self.account_identifier}")
        try:
            self._do_refresh_locked()
        except Exception as e:
            logger.warning(
                f"Force-refresh failed ({e}); falling back to interactive auth"
            )
            _require_interactive_or_raise(self.account_identifier)
            self.clear_cache()
            self.authenticate()
            # Tokens are now saved on disk; graph.request will call get_token()
            # immediately after force_refresh() returns to pick up the new token.

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


def classify_refresh_error(
    error_text: Optional[str], identifier: Optional[str] = None
) -> Optional[dict[str, str]]:
    """Classify a token-refresh error string into an actionable hint.

    Pure helper (no I/O). Recognizes a handful of AADSTS codes that have a
    clear remedy and returns ``{"code", "summary", "remedy"}``; returns
    ``None`` for unrecognized errors so callers can fall back to the raw
    string.

    The headline case is **AADSTS65002**: a Graph ``.default`` refresh whose
    shared refresh token is scoped to the Outlook grant (no Graph delegated
    consent). This is exactly the failure mode the graph-only refresh-token
    persist guard prevents going forward; this classifier explains how to
    recover an already-tainted token.

    Args:
        error_text: The error message captured from a failed refresh (may be
            ``None`` or empty).
        identifier: Optional account email to interpolate into the remedy
            command; falls back to ``<email>`` when not provided.

    Returns:
        A dict with ``code``/``summary``/``remedy`` keys, or ``None``.
    """
    if not error_text:
        return None

    who = identifier or "<email>"

    if "AADSTS65002" in error_text:
        return {
            "code": "AADSTS65002",
            "summary": (
                "shared refresh token is scoped to the Outlook grant and has "
                "no Graph delegated consent (first-party preauthorization)"
            ),
            "remedy": (
                f"re-consent Graph via device-code flow: "
                f"microsoft-mcp auth refresh {who} --force --api both"
            ),
        }
    if "AADSTS70008" in error_text or "AADSTS700082" in error_text:
        return {
            "code": "AADSTS70008",
            "summary": "the refresh token has expired or been revoked",
            "remedy": (f"re-authenticate: microsoft-mcp auth refresh {who} --force"),
        }
    if "AADSTS50173" in error_text:
        return {
            "code": "AADSTS50173",
            "summary": (
                "the refresh token was invalidated by a password/credential change"
            ),
            "remedy": (f"re-authenticate: microsoft-mcp auth refresh {who} --force"),
        }
    return None


def _refresh_one(
    identifier: str,
    tokens_dir: Path,
    client_id: Optional[str],
    tenant_id: Optional[str],
    api_type: str = "graph",
) -> dict[str, Any]:
    """Run the valid/refresh/failed flow for a single account + api_type.

    Builds a probe ``MSALRefreshTokenAuth`` for ``api_type`` and returns a
    result dict with ``identifier``, ``status``, ``expires_at``, ``error``.
    Failures do NOT call ``clear_cache()`` — the refresh token is preserved
    for later retry. The caller is responsible for tagging the returned dict
    with the ``api_type`` key.
    """
    probe = MSALRefreshTokenAuth(
        tokens_dir=tokens_dir,
        client_id=client_id,
        tenant_id=tenant_id,
        account_identifier=identifier,
        api_type=api_type,
    )
    if probe._is_token_valid():
        token_data = probe._load_access_token_data() or {}
        return {
            "identifier": identifier,
            "status": "valid",
            "expires_at": token_data.get("expires_at"),
            "error": None,
        }
    try:
        token_data = probe._do_refresh_locked()
        return {
            "identifier": identifier,
            "status": "refreshed",
            "expires_at": token_data.get("expires_at"),
            "error": None,
        }
    except Exception as e:
        stale = probe._load_access_token_data() or {}
        logger.warning(f"_refresh_one: '{identifier}' ({api_type}) refresh failed: {e}")
        result: dict[str, Any] = {
            "identifier": identifier,
            "status": "failed",
            "expires_at": stale.get("expires_at"),
            "error": str(e),
        }
        hint = classify_refresh_error(str(e), identifier=identifier)
        if hint:
            result["hint"] = hint
        return result


def refresh_all_accounts(
    tokens_dir: Optional[Path] = None,
    client_id: Optional[str] = None,
    tenant_id: Optional[str] = None,
    api_type: str = "graph",
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
        api_type: API resource to refresh: "graph" (default), "outlook", or
            "both". "both" emits one result entry per (account, api_type).
            Outlook sibling token files (``*_outlook_access_token.json``) are
            never enumerated as their own accounts.

    Returns:
        A list of result dictionaries, one per (account, api_type), each
        containing:
        - identifier (str): the account identifier (filename stem)
        - status (str): "valid" if the token was already valid (no refresh
          needed), "refreshed" if a network refresh occurred, "failed" if
          the refresh attempt errored
        - expires_at (str | None): ISO-format expiry timestamp after the
          operation, or None if the call failed
        - error (str | None): error message when status == "failed"
        - api_type (str): the API resource this entry refreshed
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

    api_types = ["graph", "outlook"] if api_type == "both" else [api_type]

    results: list[dict[str, Any]] = []

    for token_file in token_files:
        # Strip the "_access_token" suffix from the stem to get the identifier.
        identifier = token_file.stem[: -len("_access_token")]
        if identifier.endswith("_outlook"):
            # Skip outlook sibling files ({id}_outlook_access_token.json) during
            # enumeration. (A real account identifier literally ending in
            # "_outlook" would be falsely skipped, but email identifiers don't.)
            continue
        logger.info(f"refresh_all_accounts: processing account '{identifier}'")

        for current_api in api_types:
            entry = _refresh_one(
                identifier=identifier,
                tokens_dir=resolved_dir,
                client_id=client_id,
                tenant_id=tenant_id,
                api_type=current_api,
            )
            entry["api_type"] = current_api
            results.append(entry)

    return results


def _resolve_tokens_dir(tokens_dir: Optional[Path]) -> Path:
    """Resolve tokens_dir using the same precedence as MSALRefreshTokenAuth."""
    if tokens_dir is not None:
        return Path(tokens_dir)
    default_dir = Path.home() / ".config" / "microsoft-mcp" / "tokens"
    return Path(os.getenv("MICROSOFT_MCP_TOKENS_DIR", str(default_dir)))


def _decode_jwt_claims(token: str) -> dict[str, Any]:
    """Decode the payload segment of a JWT without verifying the signature.

    Microsoft access tokens are signed JWTs. The payload claims (`upn`,
    `preferred_username`, `oid`, `tid`, `aud`, `exp`) are sufficient to
    confirm which Azure AD identity the token represents. We do NOT verify
    the signature here — AAD already validated it at issue time, and we
    trust that any token that successfully calls Graph is well-formed.

    Args:
        token: Encoded JWT string (three base64url-encoded segments).

    Returns:
        Dict of claims, or {"_decode_error": str} if parsing fails.
    """
    try:
        parts = token.split(".")
        if len(parts) < 2:
            return {"_decode_error": "token has fewer than 2 segments"}
        payload = parts[1]
        # Pad to multiple of 4 for urlsafe_b64decode.
        payload += "=" * (-len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload))
    except Exception as e:
        return {"_decode_error": str(e)}


def verify_account_tokens(
    tokens_dir: Optional[Path] = None,
    live: bool = False,
    timeout_seconds: float = 15.0,
) -> list[dict[str, Any]]:
    """Verify that each saved token belongs to the account named in its filename.

    For every ``*_access_token.json`` file in the tokens directory, decode
    the JWT payload and compare its ``upn`` / ``preferred_username`` claim
    to the filename identifier. JWT claims are signed by Azure AD at issue
    time, so a mismatch is a hard signal that the wrong tokens were saved
    under that filename (e.g., the user authenticated while
    ``MICROSOFT_MCP_ACCOUNT_ID`` pointed at a different account).

    Optionally call ``GET https://graph.microsoft.com/v1.0/me`` for a live
    confirmation. ``/me`` is heavily throttled, so this is off by default.

    Args:
        tokens_dir: Directory containing token files. Defaults to
            ``MICROSOFT_MCP_TOKENS_DIR`` or
            ``~/.config/microsoft-mcp/tokens/``.
        live: If True, also call Graph ``/me`` for each token.
        timeout_seconds: HTTP timeout for the optional Graph call.

    Returns:
        List of dicts, one per account, each with:
        - identifier (str): the filename stem
        - jwt_upn (str | None): JWT ``upn`` / ``preferred_username`` claim
        - jwt_oid (str | None): JWT ``oid`` claim (immutable object id)
        - jwt_tid (str | None): JWT ``tid`` claim (tenant id)
        - jwt_aud (str | None): JWT ``aud`` claim
        - expires_at (str | None): from the saved JSON metadata
        - graph_userPrincipalName (str | None): present only when live=True
        - graph_mail (str | None): present only when live=True
        - graph_id (str | None): present only when live=True
        - graph_error (str | None): present only when live=True
        - match (bool): True if the filename identifier matches the JWT
          ``upn`` (case-insensitive) or any live Graph identifier
        - jwt_decode_error (str | None): present only on decode failure
    """
    resolved_dir = _resolve_tokens_dir(tokens_dir)

    if not resolved_dir.exists():
        logger.info(
            f"verify_account_tokens: tokens_dir {resolved_dir} does not exist, returning []"
        )
        return []

    token_files = sorted(resolved_dir.glob("*_access_token.json"))
    logger.info(
        f"verify_account_tokens: inspecting {len(token_files)} account(s) in {resolved_dir}"
    )

    results: list[dict[str, Any]] = []

    for token_file in token_files:
        identifier = token_file.stem[: -len("_access_token")]
        entry: dict[str, Any] = {
            "identifier": identifier,
            "jwt_upn": None,
            "jwt_oid": None,
            "jwt_tid": None,
            "jwt_aud": None,
            "expires_at": None,
            "match": False,
        }

        try:
            data = json.loads(token_file.read_text())
        except Exception as e:
            entry["jwt_decode_error"] = f"could not read token file: {e}"
            results.append(entry)
            continue

        entry["expires_at"] = data.get("expires_at")
        access_token = data.get("access_token") or ""
        claims = _decode_jwt_claims(access_token)

        if "_decode_error" in claims:
            entry["jwt_decode_error"] = claims["_decode_error"]
            results.append(entry)
            continue

        jwt_upn = (
            claims.get("upn")
            or claims.get("preferred_username")
            or claims.get("unique_name")
        )
        entry["jwt_upn"] = jwt_upn
        entry["jwt_oid"] = claims.get("oid")
        entry["jwt_tid"] = claims.get("tid")
        entry["jwt_aud"] = claims.get("aud")

        candidates = {c.lower() for c in (jwt_upn,) if c}

        if live:
            try:
                import httpx

                resp = httpx.get(
                    "https://graph.microsoft.com/v1.0/me",
                    headers={"Authorization": f"Bearer {access_token}"},
                    timeout=timeout_seconds,
                )
                if resp.status_code == 200:
                    body = resp.json()
                    entry["graph_userPrincipalName"] = body.get("userPrincipalName")
                    entry["graph_mail"] = body.get("mail")
                    entry["graph_id"] = body.get("id")
                    entry["graph_error"] = None
                    for v in (
                        entry["graph_userPrincipalName"],
                        entry["graph_mail"],
                    ):
                        if v:
                            candidates.add(v.lower())
                else:
                    entry["graph_userPrincipalName"] = None
                    entry["graph_mail"] = None
                    entry["graph_id"] = None
                    entry["graph_error"] = f"HTTP {resp.status_code}: {resp.text[:200]}"
            except Exception as e:
                entry["graph_userPrincipalName"] = None
                entry["graph_mail"] = None
                entry["graph_id"] = None
                entry["graph_error"] = repr(e)

        entry["match"] = identifier.lower() in candidates
        results.append(entry)

    return results


def refresh_account(
    identifier: str,
    tokens_dir: Optional[Path] = None,
    client_id: Optional[str] = None,
    tenant_id: Optional[str] = None,
    api_type: str = "graph",
) -> dict[str, Any]:
    """Refresh access token for a single MSAL account.

    Single-account analogue of :func:`refresh_all_accounts`. Looks up the
    token file at ``{tokens_dir}/{identifier}_access_token.json`` and runs
    the same valid/refresh logic via :func:`_refresh_one`. Failures do NOT
    call ``clear_cache()`` — consistent with refresh_all_accounts, this
    preserves the refresh token for later retry.

    Args:
        identifier: Account identifier (filename stem, typically email).
        tokens_dir: Directory containing token files. Defaults to
            ``MICROSOFT_MCP_TOKENS_DIR`` or
            ``~/.config/microsoft-mcp/tokens/``.
        client_id: MSAL client ID. Defaults to env var or Microsoft Office
            public client ID.
        tenant_id: MSAL tenant ID. Defaults to env var or "common".
        api_type: API resource to refresh: "graph" (default) or "outlook".

    Returns:
        Result dict with the same shape as one entry of refresh_all_accounts:
        - identifier (str)
        - status (str): "valid", "refreshed", "failed", or "missing"
        - expires_at (str | None)
        - error (str | None)
        - api_type (str): the API resource this entry refreshed
    """
    if not identifier or not identifier.strip():
        raise ValueError("identifier must be a non-empty string")

    resolved_dir = _resolve_tokens_dir(tokens_dir)
    token_file = resolved_dir / f"{identifier}_access_token.json"

    if not token_file.exists():
        logger.info(
            f"refresh_account: no token file for '{identifier}' at {token_file}"
        )
        return {
            "identifier": identifier,
            "status": "missing",
            "expires_at": None,
            "error": f"no token file at {token_file}",
        }

    result = _refresh_one(
        identifier, resolved_dir, client_id, tenant_id, api_type=api_type
    )
    result["api_type"] = api_type
    return result


def force_reauthenticate(
    identifier: str,
    tokens_dir: Optional[Path] = None,
    client_id: Optional[str] = None,
    tenant_id: Optional[str] = None,
    also_outlook: bool = False,
) -> dict[str, Any]:
    """Clear an account's saved tokens and re-run the MSAL device-code flow.

    Use when an account's refresh token is revoked or the saved credentials
    are otherwise unrecoverable. This is interactive: it prints a device
    code and verification URL to stdout/stderr and blocks until the user
    completes the browser flow.

    Args:
        identifier: Account identifier (filename stem, typically email).
        tokens_dir: Directory containing token files. Defaults to
            ``MICROSOFT_MCP_TOKENS_DIR`` or
            ``~/.config/microsoft-mcp/tokens/``.
        client_id: MSAL client ID. Defaults to env var or Microsoft Office
            public client ID.
        tenant_id: MSAL tenant ID. Defaults to env var or "common".
        also_outlook: When True, after the Graph re-auth, mint an Outlook
            access token off the freshly-minted shared refresh token (one
            extra silent refresh, no second prompt). The Outlook leg never
            overwrites the shared refresh token (see _save_tokens), so this
            is safe. Mirrors ``auth refresh --api both`` but for the re-auth
            path.

    Returns:
        Dict with keys:
        - identifier (str)
        - status (str): "reauthenticated" on success
        - expires_at (str | None)
        - signed_in_as (str | None): identifier reported by the new token
          (from JWT claims), useful for detecting drift between
          ``identifier`` and the account the user actually signed into
        - outlook (dict | None): present only when ``also_outlook`` is True —
          the result dict from the Outlook refresh (identifier/status/
          expires_at/error/api_type), or None if it could not be attempted.

    Raises:
        ValueError: if identifier is empty.
        RuntimeError: if the device-code flow fails.
    """
    if not identifier or not identifier.strip():
        raise ValueError("identifier must be a non-empty string")

    resolved_dir = _resolve_tokens_dir(tokens_dir)

    auth = MSALRefreshTokenAuth(
        tokens_dir=resolved_dir,
        client_id=client_id,
        tenant_id=tenant_id,
        account_identifier=identifier,
    )

    logger.info(f"force_reauthenticate: clearing cache for '{identifier}'")
    auth.clear_cache()

    logger.info(f"force_reauthenticate: starting device-code flow for '{identifier}'")
    result = auth.authenticate()

    token_data = auth._load_access_token_data() or {}
    expires_at = token_data.get("expires_at")
    claims = _decode_jwt_claims(result.get("access_token", ""))
    signed_in_as = (
        claims.get("upn")
        or claims.get("preferred_username")
        or claims.get("unique_name")
    )

    out: dict[str, Any] = {
        "identifier": identifier,
        "status": "reauthenticated",
        "expires_at": expires_at,
        "signed_in_as": signed_in_as,
    }

    if also_outlook:
        # Mint the Outlook token off the now-fresh shared refresh token. This
        # is a silent refresh (no second device-code prompt).
        logger.info(f"force_reauthenticate: minting Outlook token for '{identifier}'")
        out["outlook"] = _refresh_one(
            identifier, resolved_dir, client_id, tenant_id, api_type="outlook"
        )
        out["outlook"]["api_type"] = "outlook"

    return out
