#!/usr/bin/env python3
"""
Authenticate Microsoft account for use with Microsoft MCP.

Supports two authentication methods:
1. Azure SDK (default) - Interactive browser flow
2. MSAL - Device code flow (for CLI/headless environments)

Set MICROSOFT_MCP_AUTH_METHOD=msal to use device code flow.
"""

import os
import sys
from pathlib import Path

# Add src to path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent / "src"))

from dotenv import load_dotenv
from microsoft_mcp import graph
from microsoft_mcp.auth_base import AuthProvider

# Load environment variables before anything else
load_dotenv()


def get_auth_provider() -> AuthProvider:
    """Create the appropriate auth provider based on environment configuration."""
    auth_method = os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower()

    if auth_method == "msal":
        from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

        return MSALRefreshTokenAuth(
            tokens_dir=os.getenv("MICROSOFT_MCP_TOKENS_DIR"),
            client_id=os.getenv("MICROSOFT_MCP_CLIENT_ID"),
            tenant_id=os.getenv("MICROSOFT_MCP_TENANT_ID"),
            account_identifier=os.getenv("MICROSOFT_MCP_ACCOUNT_ID"),
        )
    else:
        from microsoft_mcp.auth import AzureAuthentication

        return AzureAuthentication(
            auth_record_file=os.getenv("AZURE_CRED_CACHE_FILE"),
            token_cache_file=os.getenv("AZURE_TOKEN_CACHE_FILE"),
        )


def main():
    auth_method = os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower()

    # Validate configuration based on auth method
    if auth_method == "azure":
        if not os.getenv("MICROSOFT_MCP_CLIENT_ID"):
            print("Error: MICROSOFT_MCP_CLIENT_ID environment variable is required")
            print("\nPlease set it in your .env file or environment:")
            print("export MICROSOFT_MCP_CLIENT_ID='your-app-id'")
            print("\nNote: This should be the Application (client) ID from your")
            print("Azure AD app registration configured for delegated access.")
            print("\nOptional environment variables:")
            print("- MICROSOFT_MCP_TENANT_ID: Tenant ID (defaults to 'common')")
            print(
                "- MICROSOFT_MCP_REDIRECT_URI: Custom redirect URI for non-localhost deployments"
            )
            print("\nAlternatively, use MSAL auth (device code flow):")
            print("export MICROSOFT_MCP_AUTH_METHOD=msal")
            sys.exit(1)

    print("Microsoft MCP Authentication")
    print("=" * 50)

    if auth_method == "msal":
        print("Authentication Method: MSAL (Device Code Flow)")
        print()
        print("This method uses device code flow, which:")
        print("• Works in CLI/headless environments")
        print("• Displays a code to enter in a browser")
        print("• Uses file-based token storage")
        print()
        client_id = os.getenv("MICROSOFT_MCP_CLIENT_ID")
        if client_id:
            print(f"Using custom client ID: {client_id[:8]}...")
        else:
            print("Using default Microsoft Office client ID")
        account_id = os.getenv("MICROSOFT_MCP_ACCOUNT_ID")
        if account_id:
            print(f"Using account identifier: {account_id}")
        tokens_dir = os.getenv("MICROSOFT_MCP_TOKENS_DIR", "~/.config/microsoft-mcp/tokens/")
        print(f"Token storage: {tokens_dir}")
    else:
        print("Authentication Method: Azure SDK (Browser Flow)")
        print()
        print("This method opens a browser window for sign-in.")
        redirect_uri = os.getenv("MICROSOFT_MCP_REDIRECT_URI")
        if redirect_uri:
            print(f"Using custom redirect URI: {redirect_uri}")
        else:
            print("Using default localhost redirect URI")

    print()

    # Get auth instance
    auth = get_auth_provider()

    # Set the auth instance for the graph module
    graph.set_auth_instance(auth)

    # Check if already authenticated
    try:
        print("Checking current authentication status...")

        # Check if we have valid authentication
        if auth.exists_valid_token():
            # Try to get user info to verify authentication works
            user_info = graph.request(
                "GET",
                "/me",
                params={"$select": "id,displayName,mail,userPrincipalName"},
            )

            print(f"✓ Already authenticated as: {user_info['displayName']}")
            print(
                f"  Email: {user_info.get('mail') or user_info.get('userPrincipalName')}"
            )
            print(f"  User ID: {user_info['id']}")

            # Display current token information
            try:
                import datetime

                token, expires_on = auth.get_token_with_details()
                expires_dt = datetime.datetime.fromtimestamp(expires_on)

                print(f"\n📋 Current Token Information:")
                print(f"   Token (first 20 chars): {token[:20]}...")
                print(f"   Expires on: {expires_dt.strftime('%Y-%m-%d %H:%M:%S')}")
                print(f"   Expires in: {expires_dt - datetime.datetime.now()}")
            except Exception as e:
                print(f"   ⚠ Could not retrieve token details: {e}")

            choice = input("\nDo you want to re-authenticate? (y/n): ").lower()
            if choice != "y":
                print("Using existing authentication.")
                return
            else:
                # Clear existing cache to force re-authentication
                auth.clear_cache()
                print("Authentication cache cleared. Proceeding with authentication...")
        else:
            print("No valid authentication found. Proceeding with authentication...")

    except Exception as e:
        print(f"Authentication check failed: {e}")
        print("Proceeding with authentication...")

    print()

    try:
        print("Starting authentication process...")

        if auth_method == "msal":
            print("Device code flow will display authentication instructions.")
            from microsoft_mcp.auth_msal import DEFAULT_SCOPES

            scopes = DEFAULT_SCOPES
        else:
            print("This will open a browser window for Microsoft sign-in.")
            from microsoft_mcp.auth import SCOPES

            scopes = SCOPES

        print("\nRequested permissions:")
        for scope in scopes:
            print(f"   - {scope}")
        print("\nStarting authentication...")

        # Perform authentication
        auth_result = auth.authenticate()
        print(f"\n✓ Authentication successful!")

        if auth_method == "msal":
            from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

            if isinstance(auth, MSALRefreshTokenAuth):
                print(f"Tokens saved to: {auth.tokens_dir}")
        else:
            from microsoft_mcp.auth import AzureAuthentication

            if isinstance(auth, AzureAuthentication):
                print(f"AuthenticationRecord saved to: {auth.auth_record_file}")

        # Verify authentication by getting user info
        user_info = graph.request(
            "GET", "/me", params={"$select": "id,displayName,mail,userPrincipalName"}
        )

        print(f"Signed in as: {user_info['displayName']}")
        print(f"Email: {user_info.get('mail') or user_info.get('userPrincipalName')}")
        print(f"User ID: {user_info['id']}")
        print("✓ Authentication verified")

        # Get and display token information
        try:
            import datetime

            token, expires_on = auth.get_token_with_details()
            expires_dt = datetime.datetime.fromtimestamp(expires_on)

            print(f"\n📋 Token Information:")
            print(f"   Token (first 20 chars): {token[:20]}...")
            print(f"   Expires on: {expires_dt.strftime('%Y-%m-%d %H:%M:%S')}")
            print(f"   Expires in: {expires_dt - datetime.datetime.now()}")
        except Exception as e:
            print(f"⚠ Could not retrieve token details: {e}")

    except Exception as e:
        print(f"\n✗ Authentication failed: {e}")
        sys.exit(1)

    print("\n✓ Authentication complete!")
    print("You can now use the Microsoft MCP tools.")

    if auth_method == "msal":
        print("Future runs will use cached tokens, refreshing automatically when needed.")
    else:
        print(
            "Future runs will authenticate silently using the saved AuthenticationRecord."
        )


if __name__ == "__main__":
    main()
