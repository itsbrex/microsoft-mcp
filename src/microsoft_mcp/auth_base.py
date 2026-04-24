"""
Authentication Protocol for Microsoft Graph MCP Server.

This module defines the AuthProvider protocol that all authentication
implementations must follow. This enables pluggable authentication methods
while maintaining a consistent interface.

Implementations:
- AzureAuthentication (auth.py) - Azure SDK InteractiveBrowserCredential
- MSALRefreshTokenAuth (auth_msal.py) - MSAL device code flow with file-based tokens
"""

from typing import Any, Protocol, runtime_checkable


@runtime_checkable
class AuthProvider(Protocol):
    """Protocol for pluggable authentication providers.

    All authentication implementations must provide these methods to be
    compatible with the graph module and MCP tools.
    """

    def get_token(self) -> str:
        """Get a valid access token for Microsoft Graph API calls.

        Should handle token refresh automatically if the current token
        is expired or about to expire.

        Returns:
            Valid access token string (JWT format).

        Raises:
            Exception: If token acquisition fails after all retry attempts.
        """
        ...

    def get_token_with_details(self) -> tuple[str, int]:
        """Get an access token along with its expiration timestamp.

        Returns:
            Tuple of (token_string, expires_on_unix_timestamp).

        Raises:
            Exception: If token acquisition fails.
        """
        ...

    def exists_valid_token(self) -> bool:
        """Check if a valid access token can be obtained.

        This method should attempt to verify that authentication is
        possible without triggering interactive prompts.

        Returns:
            True if a valid token exists or can be obtained silently,
            False otherwise.
        """
        ...

    def authenticate(self) -> Any:
        """Perform initial authentication.

        This method triggers the interactive authentication flow
        (browser-based or device code) and stores credentials for
        future use.

        Returns:
            Dictionary containing authentication metadata (varies by provider).

        Raises:
            Exception: If authentication fails.
        """
        ...

    def clear_cache(self) -> None:
        """Clear the authentication cache and force re-authentication.

        Removes all stored tokens and credentials, requiring the user
        to re-authenticate on the next token request.
        """
        ...
