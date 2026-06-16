import logging
import os
import sys
from dotenv import load_dotenv

logger = logging.getLogger(__name__)


def _maybe_refresh_on_startup() -> None:
    """Refresh tokens for all saved MSAL accounts before the server starts.

    Gated by:
      - MICROSOFT_MCP_AUTH_METHOD=msal (skipped for Azure, which manages
        its own token refresh internally)
      - MICROSOFT_MCP_REFRESH_ON_STARTUP != "0" (default on; set to "0"
        to opt out)

    Failures on individual accounts are logged but do not block server
    startup — a stale token will refresh on first use anyway. Catastrophic
    failures (e.g., the refresh module fails to import) are caught and
    logged as warnings.
    """
    if os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower() != "msal":
        return

    if os.getenv("MICROSOFT_MCP_REFRESH_ON_STARTUP", "1") == "0":
        print(
            "Startup token refresh disabled by MICROSOFT_MCP_REFRESH_ON_STARTUP=0",
            file=sys.stderr,
        )
        return

    try:
        from microsoft_mcp.auth_msal import refresh_all_accounts

        results = refresh_all_accounts()

        if not results:
            # Fresh install or no MSAL tokens saved yet — stay silent to
            # avoid noise on first-run startup.
            return

        print("Refreshing saved MSAL tokens at startup...", file=sys.stderr)

        n_valid = 0
        n_refreshed = 0
        n_failed = 0
        for r in results:
            identifier = r.get("identifier", "?")
            status = r.get("status", "unknown")
            if status == "valid":
                n_valid += 1
                print(f"  {identifier}: {status}", file=sys.stderr)
            elif status == "refreshed":
                n_refreshed += 1
                print(f"  {identifier}: {status}", file=sys.stderr)
            else:
                n_failed += 1
                error = r.get("error") or "unknown error"
                print(f"  {identifier}: failed - {error}", file=sys.stderr)

        print(
            f"Startup refresh complete: {n_valid} valid, {n_refreshed} refreshed, {n_failed} failed",
            file=sys.stderr,
        )
    except Exception as exc:
        logger.warning(
            "Startup token refresh failed; tokens will refresh lazily on first use: %s",
            exc,
        )


def main() -> None:
    # `microsoft-mcp signatures ...` dispatches to the signatures CLI without
    # paying the cost of importing the full tools/Graph stack.
    argv = sys.argv[1:]
    if argv and argv[0] == "signatures":
        from microsoft_mcp import signatures_cli

        sys.exit(signatures_cli.main(argv[1:]))

    # Load local development configuration before importing modules that read env.
    load_dotenv()

    from microsoft_mcp.tools import auth, auth_method, mcp

    if not os.getenv("MICROSOFT_MCP_CLIENT_ID"):
        print(
            "Error: MICROSOFT_MCP_CLIENT_ID environment variable is required",
            file=sys.stderr,
        )
        sys.exit(1)

    configured_auth_method = os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower()
    actual_auth_method = (
        "msal" if auth.__class__.__name__ == "MSALRefreshTokenAuth" else "azure"
    )
    account_identifier = os.getenv("MICROSOFT_MCP_ACCOUNT_ID")

    print(
        (
            f"Microsoft MCP auth startup: configured={configured_auth_method}, "
            f"actual={actual_auth_method}, "
            f"account={account_identifier or 'default'}"
        ),
        file=sys.stderr,
    )
    if (
        actual_auth_method != configured_auth_method
        or auth_method != configured_auth_method
    ):
        print(
            (
                "Error: authentication mode mismatch during startup. "
                f"Configured={configured_auth_method}, tools={auth_method}, actual={actual_auth_method}"
            ),
            file=sys.stderr,
        )
        sys.exit(1)

    # Option 1: Using the new class-based approach directly
    # auth_instance = AzureAuthentication()

    # Option 2: Using backward-compatibility functions (current approach)
    # # Initiate authentication flow at startup
    # try:
    #     print("Initializing Microsoft Graph authentication...", file=sys.stderr)

    #     # Try to get a token to trigger authentication if needed
    #     # This will use cached token if available, or prompt for authentication
    #     token = auth_instance.get_token()  # or auth.get_token() for backward compatibility

    #     print("✓ Authentication successful - MCP server starting...", file=sys.stderr)

    # except Exception as e:
    #     print(f"Authentication failed: {e}", file=sys.stderr)
    #     print(
    #         "Please run the authenticate.py script first to set up authentication.",
    #         file=sys.stderr,
    #     )
    #     sys.exit(1)

    _maybe_refresh_on_startup()

    mcp.run()


if __name__ == "__main__":
    main()
