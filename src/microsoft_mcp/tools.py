import asyncio
import base64
import concurrent.futures
import datetime as dt
import json
import logging
import os
import pathlib as pl
from typing import Any
from urllib.parse import quote
import httpx
from fastmcp import FastMCP
from . import graph
from .auth_base import AuthProvider
from .response_shaping import (
    cleanup_graph_payload,
    compact_location,
    shape_contact_detail,
    shape_contact_summary,
    shape_email_detail,
    shape_email_summary,
    shape_event_detail,
    shape_event_summary,
    flatten_email_address,
)
from .search_cache import get_global_cache
from .inbox_models import InboxItem
from .inbox_ranking import rank_items
from .code_mode import CodeModeRuntime, build_code_mode_runtime
from markitdown import MarkItDown, StreamInfo
from io import BytesIO
from sys import stderr

# Configure logging to stderr (MCP servers should log to stderr, not files)
logger = logging.getLogger(__name__)

# Only configure logging if it hasn't been configured yet
if not logging.getLogger().hasHandlers():
    # Create formatter
    formatter = logging.Formatter(
        "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
    )

    # Console handler (stderr) - MCP protocol requires stdout for JSON-RPC
    console_handler = logging.StreamHandler(stderr)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)

    # Configure root logger - only stderr, no file logging
    # File logging causes issues when cwd is / (read-only on macOS)
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.addHandler(console_handler)

mcp = FastMCP("microsoft-graph-mcp")


def get_response_profile(override: str = "auto") -> str:
    """Return the active response profile.

    Resolution order:
    1. *override* parameter (if not ``"auto"``)
    2. ``MICROSOFT_MCP_RESPONSE_PROFILE`` env var
    3. ``"legacy"`` (safe default for first release)

    Valid values: ``"legacy"`` | ``"assistant"``
    """
    if override and override != "auto":
        return override.lower()
    return os.getenv("MICROSOFT_MCP_RESPONSE_PROFILE", "legacy").lower()


# Create authentication instance based on MICROSOFT_MCP_AUTH_METHOD environment variable
# Supported methods: "azure" (default), "msal"
auth_method = os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower()

if auth_method == "msal":
    # MSAL-based authentication with device code flow and file-based tokens
    from .auth_msal import MSALRefreshTokenAuth

    logger.info("Using MSAL authentication method (device code flow)")

    # MSALRefreshTokenAuth has sensible defaults for all params:
    # - tokens_dir: ~/.config/microsoft-mcp/tokens/ (or MICROSOFT_MCP_TOKENS_DIR)
    # - client_id: Microsoft Office client ID (or MICROSOFT_MCP_CLIENT_ID)
    # - tenant_id: "common" (or MICROSOFT_MCP_TENANT_ID)
    # - account_identifier: "default" (or MICROSOFT_MCP_ACCOUNT_ID)
    auth: AuthProvider = MSALRefreshTokenAuth(
        tokens_dir=os.getenv("MICROSOFT_MCP_TOKENS_DIR"),
        client_id=os.getenv("MICROSOFT_MCP_CLIENT_ID"),
        tenant_id=os.getenv("MICROSOFT_MCP_TENANT_ID"),
        account_identifier=os.getenv("MICROSOFT_MCP_ACCOUNT_ID"),
    )
else:
    # Default: Azure SDK-based authentication with browser flow
    from .auth import AzureAuthentication

    logger.info("Using Azure SDK authentication method (browser flow)")
    auth: AuthProvider = AzureAuthentication(
        auth_record_file=os.getenv("AZURE_CRED_CACHE_FILE"),
        token_cache_file=os.getenv("AZURE_TOKEN_CACHE_FILE"),
    )

# Set the auth instance for the graph module
graph.set_auth_instance(auth)

markitdown = MarkItDown(enable_builtins=True)

CODE_MODE_TOOL_NAMES = (
    "search_tools",
    "list_tools",
    "tools_info",
    "get_required_keys_for_tool",
    "call_tool_chain",
)
VALID_TOOL_MODES = {"codemode_only", "hybrid"}


def _resolve_tool_mode() -> str:
    configured_tool_mode = os.getenv(
        "MICROSOFT_MCP_TOOL_MODE",
        "codemode_only",
    ).lower()
    if configured_tool_mode not in VALID_TOOL_MODES:
        logger.warning(
            "Unsupported MICROSOFT_MCP_TOOL_MODE=%s; falling back to codemode_only",
            configured_tool_mode,
        )
        return "codemode_only"
    return configured_tool_mode


tool_mode = _resolve_tool_mode()

_code_mode_runtime: CodeModeRuntime | None = None


def _run_async(coro: Any) -> Any:
    """Run an async coroutine from sync FastMCP tool functions."""

    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return asyncio.run(coro)

    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        future = executor.submit(lambda: asyncio.run(coro))
        return future.result()


def _get_code_mode_runtime() -> CodeModeRuntime:
    global _code_mode_runtime

    if _code_mode_runtime is None:
        _code_mode_runtime = _run_async(
            build_code_mode_runtime(
                mcp,
                excluded_tools=CODE_MODE_TOOL_NAMES,
                tool_provider=_list_internal_business_tools,
            )
        )
    else:
        _run_async(_code_mode_runtime.refresh())

    return _code_mode_runtime


FOLDERS = {
    k.casefold(): v
    for k, v in {
        "inbox": "inbox",
        "sent": "sentitems",
        "drafts": "drafts",
        "deleted": "deleteditems",
        "junk": "junkemail",
        "archive": "archive",
    }.items()
}

MESSAGE_SUMMARY_SELECT_FIELDS = (
    "id,subject,from,toRecipients,receivedDateTime,hasAttachments,"
    "bodyPreview,conversationId,isRead,webLink"
)
MAIL_FOLDER_SELECT_FIELDS = (
    "id,displayName,parentFolderId,childFolderCount,totalItemCount,"
    "unreadItemCount,isHidden"
)
MASTER_CATEGORY_SELECT_FIELDS = "id,displayName,color"


def _shape_mail_folder(raw: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": raw["id"],
        "display_name": raw.get("displayName", ""),
        "parent_folder_id": raw.get("parentFolderId"),
        "child_folder_count": raw.get("childFolderCount", 0),
        "total_item_count": raw.get("totalItemCount", 0),
        "unread_item_count": raw.get("unreadItemCount", 0),
        "is_hidden": bool(raw.get("isHidden", False)),
    }


def _shape_master_category(raw: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": raw.get("id"),
        "display_name": raw.get("displayName", ""),
        "color": raw.get("color", ""),
    }


def _list_master_category_rows(limit: int = 100) -> list[dict[str, Any]]:
    return list(
        graph.request_paginated(
            "/me/outlook/masterCategories",
            params={
                "$top": min(limit, 100),
                "$select": MASTER_CATEGORY_SELECT_FIELDS,
            },
            limit=limit,
        )
    )


def _resolve_master_category(category: str) -> dict[str, Any]:
    target = category.strip()
    if not target:
        raise ValueError("Category cannot be empty")

    categories = _list_master_category_rows(limit=500)

    for item in categories:
        if item.get("id") == target:
            return item

    matches = [
        item
        for item in categories
        if item.get("displayName", "").casefold() == target.casefold()
    ]
    if not matches:
        raise ValueError(f"Master category '{category}' not found")
    if len(matches) > 1:
        raise ValueError(
            f"Master category '{category}' is ambiguous; use the category ID instead"
        )
    return matches[0]


def _list_mail_folder_children(
    *,
    parent_folder_id: str | None = None,
    include_hidden: bool = False,
    limit: int = 100,
) -> list[dict[str, Any]]:
    endpoint = (
        "/me/mailFolders"
        if parent_folder_id is None
        else f"/me/mailFolders/{parent_folder_id}/childFolders"
    )
    params: dict[str, Any] = {
        "$top": min(limit, 100),
        "$select": MAIL_FOLDER_SELECT_FIELDS,
    }
    if include_hidden:
        params["includeHiddenFolders"] = "true"
    return list(graph.request_paginated(endpoint, params=params, limit=limit))


def _walk_mail_folders(
    *,
    parent_folder_id: str | None = None,
    include_hidden: bool = False,
    limit: int = 100,
    recursive: bool = False,
) -> list[dict[str, Any]]:
    if not recursive:
        return _list_mail_folder_children(
            parent_folder_id=parent_folder_id,
            include_hidden=include_hidden,
            limit=limit,
        )

    results: list[dict[str, Any]] = []
    queue: list[str | None] = [parent_folder_id]
    seen_ids: set[str] = set()

    while queue and len(results) < limit:
        current_parent = queue.pop(0)
        remaining = limit - len(results)
        children = _list_mail_folder_children(
            parent_folder_id=current_parent,
            include_hidden=include_hidden,
            limit=remaining,
        )
        for child in children:
            folder_id = child.get("id")
            if not folder_id or folder_id in seen_ids:
                continue
            seen_ids.add(folder_id)
            results.append(child)
            if child.get("childFolderCount", 0):
                queue.append(folder_id)
            if len(results) >= limit:
                break

    return results


def _find_mail_folder(folder: str, include_hidden: bool = False) -> dict[str, Any]:
    destination = folder.strip().strip("/")
    if not destination:
        raise ValueError("Mail folder cannot be empty")

    matches = _walk_mail_folders(
        include_hidden=include_hidden,
        recursive=True,
        limit=500,
    )
    normalized_destination = destination.casefold()

    path_matches: list[dict[str, Any]] = []
    name_matches: list[dict[str, Any]] = []
    folders_by_id = {item["id"]: item for item in matches if item.get("id")}

    def build_path(item: dict[str, Any]) -> str:
        segments = [item.get("displayName", "")]
        current_parent = item.get("parentFolderId")
        while current_parent and current_parent in folders_by_id:
            parent = folders_by_id[current_parent]
            segments.append(parent.get("displayName", ""))
            current_parent = parent.get("parentFolderId")
        return "/".join(reversed([segment for segment in segments if segment]))

    for item in matches:
        display_name = item.get("displayName", "")
        if display_name.casefold() == normalized_destination:
            name_matches.append(item)
        if "/" in destination and build_path(item).casefold() == normalized_destination:
            path_matches.append(item)

    resolved_matches = path_matches or name_matches

    if not resolved_matches:
        raise ValueError(f"Mail folder '{folder}' not found")
    if len(resolved_matches) > 1:
        options = sorted(build_path(item) for item in resolved_matches)
        raise ValueError(
            "Mail folder name is ambiguous. Use the full folder path instead: "
            + ", ".join(options)
        )
    return resolved_matches[0]


# ============================================================================
# Account Management Tools (Multi-account support)
# ============================================================================


@mcp.tool
def list_accounts() -> list[dict[str, Any]]:
    """List all authenticated Microsoft accounts.

    Returns a list of all available accounts that have been authenticated,
    including their email addresses, identifiers, and token expiration status.
    Use this to see which accounts are available before switching.

    Returns:
        List of account dictionaries, each containing:
        - identifier: The account identifier (usually email address)
        - email: The email address associated with the account
        - expires_at: When the access token expires
        - is_active: Whether this is the currently active account

    Examples:
        - list_accounts() - See all available authenticated accounts
    """
    logger.info("list_accounts called")

    tokens_dir = pl.Path(
        os.getenv(
            "MICROSOFT_MCP_TOKENS_DIR",
            pl.Path.home() / ".config" / "microsoft-mcp" / "tokens",
        )
    )

    accounts = []
    if not tokens_dir.exists():
        logger.info("Token directory does not exist, returning empty list")
        return accounts

    # Get current active account identifier
    active_identifier = None
    if hasattr(auth, "account_identifier"):
        active_identifier = auth.account_identifier

    for token_file in tokens_dir.glob("*_access_token.json"):
        try:
            data = json.loads(token_file.read_text())
            identifier = token_file.stem.replace("_access_token", "")
            accounts.append(
                {
                    "identifier": identifier,
                    "email": data.get("email", identifier),
                    "expires_at": data.get("expires_at"),
                    "is_active": identifier == active_identifier,
                }
            )
        except Exception as e:
            logger.warning(f"Failed to read token file {token_file}: {e}")
            continue

    logger.info(f"list_accounts found {len(accounts)} accounts")
    return accounts


@mcp.tool
def set_active_account(account: str) -> dict[str, str]:
    """Switch the active Microsoft account.

    Changes which Microsoft account is used for all subsequent API calls.
    The account must have been previously authenticated.

    Args:
        account: Email address or identifier of the account to activate.
                 Use list_accounts() to see available accounts.

    Returns:
        Confirmation dictionary with:
        - status: "switched" on success
        - active_account: The new active account identifier

    Raises:
        ValueError: If no tokens are found for the specified account

    Examples:
        - set_active_account("work@company.com") - Switch to work account
        - set_active_account("personal@outlook.com") - Switch to personal account
    """
    global auth

    logger.info(f"set_active_account called: account={account}")

    # Verify this is MSAL auth method
    if auth_method != "msal":
        raise ValueError(
            "Account switching is only supported with MSAL authentication method"
        )

    # Verify account exists
    tokens_dir = pl.Path(
        os.getenv(
            "MICROSOFT_MCP_TOKENS_DIR",
            pl.Path.home() / ".config" / "microsoft-mcp" / "tokens",
        )
    )
    token_file = tokens_dir / f"{account}_access_token.json"

    if not token_file.exists():
        available = [
            f.stem.replace("_access_token", "")
            for f in tokens_dir.glob("*_access_token.json")
        ]
        raise ValueError(
            f"No tokens found for account: {account}. Available accounts: {available}"
        )

    # Import here to avoid circular import issues
    from .auth_msal import MSALRefreshTokenAuth

    # Create new auth instance for this account
    auth = MSALRefreshTokenAuth(
        tokens_dir=os.getenv("MICROSOFT_MCP_TOKENS_DIR"),
        client_id=os.getenv("MICROSOFT_MCP_CLIENT_ID"),
        tenant_id=os.getenv("MICROSOFT_MCP_TENANT_ID"),
        account_identifier=account,
    )
    graph.set_auth_instance(auth)

    logger.info(f"set_active_account successful: switched to {account}")
    return {"status": "switched", "active_account": account}


@mcp.tool
def get_active_account() -> dict[str, Any]:
    """Get the currently active Microsoft account.

    Returns information about which account is currently being used for API calls.

    Returns:
        Dictionary containing:
        - identifier: The account identifier
        - email: The email address (if available)
        - expires_at: When the access token expires (if available)
        - auth_method: The authentication method being used ("msal" or "azure")

    Examples:
        - get_active_account() - See which account is currently active
    """
    logger.info("get_active_account called")

    result: dict[str, Any] = {"auth_method": auth_method}

    if hasattr(auth, "account_identifier"):
        result["identifier"] = auth.account_identifier

        # Try to load token data for more details
        try:
            if hasattr(auth, "_load_access_token_data"):
                token_data = auth._load_access_token_data()
                if token_data:
                    result["email"] = token_data.get("email", auth.account_identifier)
                    result["expires_at"] = token_data.get("expires_at")
        except Exception as e:
            logger.warning(f"Could not load token data: {e}")
            result["email"] = auth.account_identifier
    else:
        result["identifier"] = "default"
        result["email"] = "unknown"

    logger.info(f"get_active_account returning: {result}")
    return result


# ============================================================================
# Utility Functions
# ============================================================================


def convert_to_markdown(html: str, mimetype: str = "text/html") -> str:
    """Convert HTML content to Markdown format."""
    # Use MarkItDown to convert HTML to Markdown
    stream = BytesIO()
    stream.write(html.encode("utf-8"))
    stream.seek(0)
    return markitdown.convert(
        stream, stream_info=StreamInfo(mimetype=mimetype)
    ).text_content


@mcp.tool
def get_user_details(email: str | None = None) -> dict[str, Any]:
    """Get details about a user - either the logged-in user or another user by email address.

    Retrieves user profile information including display name, email, job title, department,
    office location, and other directory information. When no email is provided, returns
    details for the currently signed-in user. When an email is provided, looks up that
    specific user's public profile information.

    Args:
        email: Optional email address of the user to look up. If None, returns current user's details.
               Must be a valid email address format (e.g., "user@company.com").

    Returns:
        User object containing profile information:
        - Basic info: id, displayName, mail, userPrincipalName, givenName, surname
        - Professional: jobTitle, department, companyName, officeLocation, businessPhones
        - Directory info: accountEnabled, userType, createdDateTime
        - When looking up other users, some fields may be limited based on directory permissions

    Examples:
        - get_user_details() - Get current user's profile information
        - get_user_details("colleague@company.com") - Look up specific user's profile
        - get_user_details("manager@company.com") - Get manager's contact information

    Note: Looking up other users requires User.ReadBasic.All permission and the target
    user must be visible in your organization's directory.
    """
    logger.info(f"get_user_details called: email={email}")

    try:
        if email is None:
            # Get current user's details
            result = graph.request("GET", "/me")
            logger.info("get_user_details successful: retrieved current user details")
        else:
            # Look up user by email address
            # Use the /users/{email} endpoint to get user by their email/UPN
            result = graph.request("GET", f"/users/{email}")
            if not result:
                logger.error(
                    f"get_user_details failed: User with email {email} not found"
                )
                raise ValueError(f"User with email {email} not found")
            logger.info(
                f"get_user_details successful: retrieved details for user {email}"
            )

        return cleanup_graph_payload(result)
    except Exception as e:
        logger.error(
            f"get_user_details failed for email={email}: {str(e)}", exc_info=True
        )
        raise


@mcp.prompt
def prepare_work_day():
    return """
    You are a helpful assistant that helps the user to prepare for their work day.
    You can use tools to get information about their calendar, emails, contacts and availability accessing the MS Graph API.
    
    """


@mcp.prompt
def utcp_codemode_usage():
    """Guide assistants toward discovery-first code-mode workflows."""

    return CodeModeRuntime.AGENT_PROMPT_TEMPLATE


@mcp.tool
def list_tools() -> dict[str, Any]:
    """List the active Microsoft MCP tool registry for integrated code-mode workflows.

    Returns the auth-aware business tools that are available inside the
    `microsoft.<tool>()` namespace used by `call_tool_chain`.
    """

    runtime = _get_code_mode_runtime()
    tools = _run_async(runtime.list_tools())
    return {"namespace": runtime.namespace, "count": len(tools), "tools": tools}


@mcp.tool
def search_tools(task_description: str, limit: int = 10) -> dict[str, Any]:
    """Search the active Microsoft tool registry using a natural-language query."""

    runtime = _get_code_mode_runtime()
    tools = _run_async(runtime.search_tools(task_description, limit=limit))
    return {
        "query": task_description,
        "namespace": runtime.namespace,
        "count": len(tools),
        "tools": tools,
    }


@mcp.tool
def tools_info(tool_names: list[str]) -> dict[str, Any]:
    """Return detailed metadata and generated Python interfaces for selected tools."""

    runtime = _get_code_mode_runtime()
    tools = _run_async(runtime.tools_info(tool_names))
    return {"namespace": runtime.namespace, "count": len(tools), "tools": tools}


@mcp.tool
def get_required_keys_for_tool(tool_name: str) -> dict[str, Any]:
    """Return the required configuration keys for a code-mode-visible tool."""

    runtime = _get_code_mode_runtime()
    return _run_async(runtime.get_required_keys_for_tool(tool_name))


@mcp.tool
def call_tool_chain(
    code: str,
    timeout: float = 30.0,
    include_interfaces: bool = False,
) -> dict[str, Any]:
    """Execute a multi-step Python workflow against the active Microsoft tool namespace.

    The sandbox exposes the active business tools as `microsoft.<tool>()`,
    generated interfaces through `interfaces`, and per-tool interface lookup
    via `get_tool_interface(name)`.

    Set ``include_interfaces=True`` only when you need the generated TypedDict
    catalog in the response (default False to keep token cost low).
    """

    runtime = _get_code_mode_runtime()
    return _run_async(
        runtime.call_tool_chain(
            code, timeout=timeout, include_interfaces=include_interfaces
        )
    )


@mcp.tool
def is_logged_in() -> bool:
    return auth.exists_valid_token()


@mcp.tool
def login() -> str:
    """Ensure the user is authenticated and return user info.
    Raises an error if authentication fails.

    `login` is required only if tools report errors.
    """

    if not auth.exists_valid_token():
        try:
            auth.get_token()
            return "logged in"
        except Exception as e:
            logger.error(f"login failed: {str(e)}", exc_info=True)
            raise RuntimeError("Login failed, please check authentication settings.")

    else:
        return "already logged in"


@mcp.tool
def list_mail_folders(
    parent_folder: str | None = None,
    recursive: bool = False,
    include_hidden: bool = False,
    limit: int = 100,
) -> list[dict[str, Any]]:
    """List Outlook mail folders.

    Args:
        parent_folder: Optional folder alias, ID, name, or path whose children should be listed.
            When omitted, lists top-level mail folders.
        recursive: Whether to walk the folder tree beneath the selected parent.
        include_hidden: Whether to include hidden folders when Graph permits it.
        limit: Maximum number of folders to return.

    Returns:
        List of normalized folder objects containing IDs, display names, parent IDs,
        child-folder counts, unread counts, and total-item counts.
    """

    logger.info(
        "list_mail_folders called: parent_folder=%s, recursive=%s, include_hidden=%s, limit=%s",
        parent_folder,
        recursive,
        include_hidden,
        limit,
    )

    try:
        parent_folder_id = None
        if parent_folder:
            parent_folder_id = _resolve_mail_folder(
                parent_folder,
                include_hidden=include_hidden,
            )

        raw_folders = _walk_mail_folders(
            parent_folder_id=parent_folder_id,
            include_hidden=include_hidden,
            limit=min(limit, 500),
            recursive=recursive,
        )
        return [_shape_mail_folder(folder) for folder in raw_folders]
    except Exception as e:
        logger.error("list_mail_folders failed: %s", str(e), exc_info=True)
        raise


@mcp.tool
def get_mail_folder(folder: str, include_hidden: bool = False) -> dict[str, Any]:
    """Get a mail folder by alias, ID, display name, or slash-delimited path."""

    logger.info(
        "get_mail_folder called: folder=%s, include_hidden=%s",
        folder,
        include_hidden,
    )

    try:
        if folder.strip().casefold() in FOLDERS:
            folder_id = FOLDERS[folder.strip().casefold()]
            raw = graph.request(
                "GET",
                f"/me/mailFolders/{folder_id}",
                params={"$select": MAIL_FOLDER_SELECT_FIELDS},
            )
            if not raw:
                raise ValueError(f"Mail folder '{folder}' not found")
            return _shape_mail_folder(raw)

        return _shape_mail_folder(
            _find_mail_folder(folder, include_hidden=include_hidden)
        )
    except Exception as e:
        logger.error(
            "get_mail_folder failed for folder=%s: %s", folder, str(e), exc_info=True
        )
        raise


@mcp.tool
def create_mail_folder(
    display_name: str,
    parent_folder: str | None = None,
) -> dict[str, Any]:
    """Create a new Outlook mail folder under the mailbox root or another folder."""

    logger.info(
        "create_mail_folder called: display_name=%s, parent_folder=%s",
        display_name,
        parent_folder,
    )

    folder_name = display_name.strip()
    if not folder_name:
        raise ValueError("display_name cannot be empty")

    try:
        endpoint = "/me/mailFolders"
        if parent_folder:
            parent_folder_id = _resolve_mail_folder(parent_folder)
            endpoint = f"/me/mailFolders/{parent_folder_id}/childFolders"

        raw = graph.request(
            "POST",
            endpoint,
            json={"displayName": folder_name},
        )
        if not raw:
            raise ValueError("Mail folder could not be created")
        return _shape_mail_folder(raw)
    except Exception as e:
        logger.error("create_mail_folder failed: %s", str(e), exc_info=True)
        raise


@mcp.tool
def rename_mail_folder(folder: str, new_display_name: str) -> dict[str, Any]:
    """Rename an Outlook mail folder."""

    logger.info(
        "rename_mail_folder called: folder=%s, new_display_name=%s",
        folder,
        new_display_name,
    )

    folder_name = new_display_name.strip()
    if not folder_name:
        raise ValueError("new_display_name cannot be empty")

    try:
        folder_id = _resolve_mail_folder(folder)
        raw = graph.request(
            "PATCH",
            f"/me/mailFolders/{folder_id}",
            json={"displayName": folder_name},
        )
        if not raw:
            raw = graph.request(
                "GET",
                f"/me/mailFolders/{folder_id}",
                params={"$select": MAIL_FOLDER_SELECT_FIELDS},
            )
        if not raw:
            raise ValueError(f"Mail folder '{folder}' not found")
        return _shape_mail_folder(raw)
    except Exception as e:
        logger.error(
            "rename_mail_folder failed for folder=%s: %s", folder, str(e), exc_info=True
        )
        raise


@mcp.tool
def delete_mail_folder(folder: str) -> dict[str, Any]:
    """Delete an Outlook mail folder."""

    logger.info("delete_mail_folder called: folder=%s", folder)

    try:
        folder_id = _resolve_mail_folder(folder)
        graph.request("DELETE", f"/me/mailFolders/{folder_id}")
        return {"status": "deleted", "folder_id": folder_id}
    except Exception as e:
        logger.error(
            "delete_mail_folder failed for folder=%s: %s", folder, str(e), exc_info=True
        )
        raise


@mcp.tool
def list_master_categories(limit: int = 100) -> list[dict[str, Any]]:
    """List Outlook master categories available in the signed-in mailbox.

    These are the categories that appear in Outlook with colors and can be reused
    across messages and events.

    Args:
        limit: Maximum number of categories to return.

    Returns:
        List of normalized master-category objects containing id, display_name, and color.
    """

    logger.info("list_master_categories called: limit=%s", limit)

    try:
        return [
            _shape_master_category(item)
            for item in _list_master_category_rows(limit=min(limit, 500))
        ]
    except Exception as e:
        logger.error("list_master_categories failed: %s", str(e), exc_info=True)
        raise


@mcp.tool
def get_master_category(category: str) -> dict[str, Any]:
    """Get a master category by ID or display name."""

    logger.info("get_master_category called: category=%s", category)

    try:
        return _shape_master_category(_resolve_master_category(category))
    except Exception as e:
        logger.error(
            "get_master_category failed for category=%s: %s",
            category,
            str(e),
            exc_info=True,
        )
        raise


@mcp.tool
def create_master_category(display_name: str, color: str) -> dict[str, Any]:
    """Create an Outlook master category with a display name and color."""

    logger.info(
        "create_master_category called: display_name=%s, color=%s",
        display_name,
        color,
    )

    normalized_name = display_name.strip()
    normalized_color = color.strip()
    if not normalized_name:
        raise ValueError("display_name cannot be empty")
    if not normalized_color:
        raise ValueError("color cannot be empty")

    try:
        raw = graph.request(
            "POST",
            "/me/outlook/masterCategories",
            json={"displayName": normalized_name, "color": normalized_color},
        )
        if not raw:
            raise ValueError("Master category could not be created")
        return _shape_master_category(raw)
    except Exception as e:
        logger.error(
            "create_master_category failed for display_name=%s: %s",
            display_name,
            str(e),
            exc_info=True,
        )
        raise


@mcp.tool
def update_master_category(category: str, color: str) -> dict[str, Any]:
    """Update the color of an existing Outlook master category."""

    logger.info("update_master_category called: category=%s, color=%s", category, color)

    normalized_color = color.strip()
    if not normalized_color:
        raise ValueError("color cannot be empty")

    try:
        resolved = _resolve_master_category(category)
        category_id = resolved.get("id")
        if not category_id:
            raise ValueError(f"Master category '{category}' has no ID")

        raw = graph.request(
            "PATCH",
            f"/me/outlook/masterCategories/{quote(category_id, safe='')}",
            json={"color": normalized_color},
        )
        if not raw:
            raw = {
                **resolved,
                "color": normalized_color,
            }
        return _shape_master_category(raw)
    except Exception as e:
        logger.error(
            "update_master_category failed for category=%s: %s",
            category,
            str(e),
            exc_info=True,
        )
        raise


@mcp.tool
def delete_master_category(category: str) -> dict[str, Any]:
    """Delete an Outlook master category by ID or display name."""

    logger.info("delete_master_category called: category=%s", category)

    try:
        resolved = _resolve_master_category(category)
        category_id = resolved.get("id")
        if not category_id:
            raise ValueError(f"Master category '{category}' has no ID")
        graph.request(
            "DELETE",
            f"/me/outlook/masterCategories/{quote(category_id, safe='')}",
        )
        return {
            "status": "deleted",
            "id": category_id,
            "display_name": resolved.get("displayName", ""),
        }
    except Exception as e:
        logger.error(
            "delete_master_category failed for category=%s: %s",
            category,
            str(e),
            exc_info=True,
        )
        raise


@mcp.tool
def ensure_master_categories(
    categories: list[dict[str, str]],
    update_colors: bool = False,
) -> dict[str, Any]:
    """Ensure a set of Outlook master categories exists.

    Args:
        categories: List of category specs with `display_name` and `color`.
        update_colors: Whether to update the color of existing categories when the
            requested color differs.

    Returns:
        Summary of created, updated, and already-existing categories.
    """

    logger.info(
        "ensure_master_categories called: count=%s, update_colors=%s",
        len(categories),
        update_colors,
    )

    existing_rows = _list_master_category_rows(limit=500)
    by_name = {
        item.get("displayName", "").casefold(): item
        for item in existing_rows
        if item.get("displayName")
    }

    created: list[dict[str, Any]] = []
    updated: list[dict[str, Any]] = []
    existing: list[dict[str, Any]] = []

    for spec in categories:
        display_name = spec.get("display_name", "").strip()
        color = spec.get("color", "").strip()
        if not display_name:
            raise ValueError("Each category spec requires a non-empty display_name")
        if not color:
            raise ValueError(
                f"Category '{display_name}' requires a non-empty color value"
            )

        current = by_name.get(display_name.casefold())
        if current is None:
            created_category = create_master_category.fn(
                display_name=display_name,
                color=color,
            )
            created.append(created_category)
            by_name[display_name.casefold()] = {
                "id": created_category.get("id"),
                "displayName": created_category["display_name"],
                "color": created_category["color"],
            }
            continue

        if update_colors and current.get("color") != color:
            updated_category = update_master_category.fn(
                category=current.get("id") or display_name,
                color=color,
            )
            updated.append(updated_category)
            by_name[display_name.casefold()] = {
                "id": updated_category.get("id"),
                "displayName": updated_category["display_name"],
                "color": updated_category["color"],
            }
            continue

        existing.append(_shape_master_category(current))

    return {
        "requested": len(categories),
        "created": created,
        "updated": updated,
        "existing": existing,
    }


@mcp.tool
def list_emails(
    folder: str = "inbox",
    limit: int = 10,
    body_max_length: int = 2000,
    include_body: bool = False,
    start_date: str | None = None,
    end_date: str | None = None,
    response_profile: str = "auto",
) -> list[dict[str, Any]]:
    """List emails from a specified folder in the user's mailbox.

    Retrieves emails from common folders like inbox, sent, drafts, etc. Results are ordered by
    received date (most recent first). Use this to get an overview of emails for a specific date range,
    or find recent messages.

    Args:
        folder: Folder alias, ID, display name, or slash-delimited path to search in.
            Supported aliases include "inbox", "sent", "drafts", "deleted", "junk", and "archive".
        limit: Maximum number of emails to retrieve (1-100, defaults to 10)
        body_max_length: Maximum characters for email body content (default 2000, will truncate if longer)
        include_body: Whether to include email body content (affects response size)
        start_date: Optional start date in ISO format (UTC timezone, e.g., "2024-09-01T00:00:00Z") to filter emails from this date onwards
        end_date: Optional end date in ISO format (UTC timezone, e.g., "2024-09-30T23:59:59Z") to filter emails up to this date
        response_profile: Response shaping profile ("auto", "legacy", or "assistant"). "auto" defers to MICROSOFT_MCP_RESPONSE_PROFILE env var.

    Returns:
        List of email objects containing id, subject, sender, recipients, date, attachments info,
        and optionally body content. Each email has fields like 'id', 'subject', 'from', 'receivedDateTime'.
        The most recent email (within the specified date range) will be the first included in the results.
        Contains also a deep link to the conversation as `conversation_url` that can be shown to the user to open the email
    Examples:
        - list_emails() - Get 10 most recent inbox emails
        - list_emails(folder="sent", limit=20) - Get 20 recent sent emails
        - list_emails(folder="Cresa Deals of the Week", limit=20) - Get emails from a custom folder
        - list_emails(include_body=False) - Get emails without body content for faster response
        - list_emails(start_date="2024-09-01T00:00:00Z", end_date="2024-09-01T23:59:59Z") - Get emails received on September 1st, 2024
        - list_emails(start_date="2024-08-01T00:00:00Z") - Get emails from August 1st, 2024 onwards
        - list_emails(end_date="2024-08-31T23:59:59Z") - Get emails up to August 31st, 2024
    """
    profile = get_response_profile(response_profile)
    if profile == "assistant":
        include_body = False

    logger.info(
        f"list_emails called: folder={folder}, limit={limit}, include_body={include_body}, start_date={start_date}, end_date={end_date}, profile={profile}"
    )

    try:
        folder_path = _resolve_mail_folder(folder)

        if include_body:
            select_fields = "id,subject,from,toRecipients,ccRecipients,receivedDateTime,hasAttachments,body,conversationId,isRead"
        else:
            select_fields = "id,subject,from,toRecipients,receivedDateTime,hasAttachments,bodyPreview,conversationId,isRead"

        params = {
            "$top": min(limit, 100),
            "$select": select_fields,
            "$orderby": "receivedDateTime desc",
        }

        # Add date filtering if provided
        filter_conditions = []
        if start_date:
            filter_conditions.append(f"receivedDateTime ge {start_date}")
        if end_date:
            filter_conditions.append(f"receivedDateTime le {end_date}")

        if filter_conditions:
            params["$filter"] = " and ".join(filter_conditions)

        raw_emails = list(
            graph.request_paginated(
                f"/me/mailFolders/{folder_path}/messages",
                params=params,
                limit=limit,
            )
        )

        if include_body:
            # Detail mode: include body, truncate if needed
            results = []
            for email in raw_emails:
                if "body" in email and "content" in email["body"]:
                    content = email["body"]["content"]
                    if len(content) > body_max_length:
                        email["body"]["content"] = (
                            content[:body_max_length]
                            + f"\n\n[Content truncated - {len(content)} total characters]"
                        )
                        email["body"]["truncated"] = True
                        email["body"]["total_length"] = len(content)
                results.append(shape_email_detail(email))
        else:
            # Summary mode: compact, no body
            results = [shape_email_summary(e) for e in raw_emails]

        logger.info(
            f"list_emails successful: retrieved {len(results)} emails from folder {folder}"
            + (
                f" with date filter start_date={start_date}, end_date={end_date}"
                if start_date or end_date
                else ""
            )
        )
        return results
    except Exception as e:
        logger.error(f"list_emails failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_email(
    email_id: str,
    include_body: bool = True,
    body_max_length: int = 5000,
    include_attachments: bool = True,
) -> dict[str, Any]:
    """Get detailed information about a specific email by its ID.

    Retrieves complete email details including headers, body content, and attachment metadata.
    Body content can be truncated to manage response size. Use this when you need full email details
    after finding emails with list_emails or search_emails.

    Args:
        email_id: Unique identifier of the email (get from list_emails or search results)
        include_body: Whether to include the email body content in the response
        body_max_length: Maximum characters for body content (default 50000, will truncate if longer)
        include_attachments: Whether to include attachment metadata (names, sizes, types)

    Returns:
        Email object with complete details including:
        - Basic info: id, subject, from, to, cc, receivedDateTime, isRead
        - Body: content as text or markdown, contentType, truncation info if applicable
        - Attachments: list with id, name, size, contentType for each attachment
        - Conversation: conversationId for threading
        - a deep link to the conversation as `conversation_url` that can be shown to the user to open the email

    Examples:
        - get_email("AAMkAD...") - Get full email details
        - get_email(email_id, include_body=False) - Get headers only without body
        - get_email(email_id, body_max_length=1000) - Limit body to 1000 characters
    """
    logger.info(
        f"get_email called: email_id={email_id}, include_body={include_body}, body_max_length={body_max_length}, include_attachments={include_attachments}"
    )

    try:
        params = {}
        if include_attachments:
            params["$expand"] = "attachments($select=id,name,size,contentType)"

        raw = graph.request("GET", f"/me/messages/{email_id}", params=params)
        if not raw:
            logger.error(f"get_email failed: Email with ID {email_id} not found")
            raise ValueError(f"Email with ID {email_id} not found")

        # Convert HTML to markdown and truncate body if needed
        if include_body and "body" in raw and "content" in raw["body"]:
            if raw["body"]["contentType"].lower() == "html":
                raw["body"]["content"] = convert_to_markdown(raw["body"]["content"])
                raw["body"]["contentType"] = "text/markdown"

            content = raw["body"]["content"]
            if len(content) > body_max_length:
                raw["body"]["content"] = (
                    content[:body_max_length]
                    + f"\n\n[Content truncated - {len(content)} total characters]"
                )
                raw["body"]["truncated"] = True
                raw["body"]["total_length"] = len(content)
        elif not include_body and "body" in raw:
            del raw["body"]

        result = shape_email_detail(raw)

        # Remove attachment content bytes to reduce size
        if "attachments" in raw and raw["attachments"]:
            result["attachments"] = [
                {k: v for k, v in a.items() if k != "contentBytes"}
                for a in raw["attachments"]
            ]

        logger.info(f"get_email successful: retrieved email {email_id}")
        return result
    except Exception as e:
        logger.error(
            f"get_email failed for email_id={email_id}: {str(e)}", exc_info=True
        )
        raise


def _normalize_draft_type(draft_type: str) -> str:
    normalized = draft_type.strip().casefold()
    normalized = {
        "replyall": "reply_all",
        "reply-all": "reply_all",
        "reply all": "reply_all",
    }.get(normalized, normalized)
    normalized = normalized.replace("-", "_").replace(" ", "_")
    if normalized not in {"new", "reply", "reply_all"}:
        raise ValueError("draft_type must be one of: new, reply, reply_all")
    return normalized


def _normalize_body_content_type(body_content_type: str) -> str:
    normalized = body_content_type.strip().casefold()
    if normalized not in {"text", "html"}:
        raise ValueError("body_content_type must be either 'text' or 'html'")
    return normalized


def _build_recipient_objects(
    recipients: list[str] | None,
) -> list[dict[str, Any]] | None:
    if recipients is None:
        return None

    normalized_recipients = []
    for recipient in recipients:
        address = recipient.strip()
        if address:
            normalized_recipients.append({"emailAddress": {"address": address}})

    return normalized_recipients


def _build_message_update_payload(
    *,
    subject: str | None,
    body: str | None,
    body_content_type: str,
    to_recipients: list[dict[str, Any]] | None,
    cc_recipients: list[dict[str, Any]] | None,
    bcc_recipients: list[dict[str, Any]] | None,
) -> dict[str, Any]:
    payload: dict[str, Any] = {}

    if subject is not None:
        payload["subject"] = subject
    if body is not None:
        payload["body"] = {
            "contentType": body_content_type,
            "content": body,
        }
    if to_recipients is not None:
        payload["toRecipients"] = to_recipients
    if cc_recipients is not None:
        payload["ccRecipients"] = cc_recipients
    if bcc_recipients is not None:
        payload["bccRecipients"] = bcc_recipients

    return payload


def _shape_email_draft(raw: dict[str, Any]) -> dict[str, Any]:
    draft = shape_email_detail(raw)

    if "bccRecipients" in raw:
        draft["bcc"] = [flatten_email_address(r) for r in raw["bccRecipients"]]

    if raw.get("createdDateTime"):
        draft["created"] = raw["createdDateTime"]

    if raw.get("lastModifiedDateTime"):
        draft["last_modified"] = raw["lastModifiedDateTime"]

    draft["is_draft"] = raw.get("isDraft", True)
    return draft


@mcp.tool
def create_email_draft(
    draft_type: str = "new",
    email_id: str | None = None,
    to_recipients: list[str] | None = None,
    cc_recipients: list[str] | None = None,
    bcc_recipients: list[str] | None = None,
    subject: str | None = None,
    body: str | None = None,
    body_content_type: str = "text",
) -> dict[str, Any]:
    """Create an Outlook email draft without sending it.

    Supports three draft modes:
    - `new`: create a brand-new draft addressed to the supplied recipients
    - `reply`: create a reply draft to the sender of an existing message
    - `reply_all`: create a reply-all draft to an existing message thread

    This tool never sends mail. It only creates or updates draft messages in the user's mailbox.

    Args:
        draft_type: Draft mode (`new`, `reply`, or `reply_all`)
        email_id: Required for `reply` and `reply_all`; the source message ID being answered
        to_recipients: Optional list of recipient email addresses
        cc_recipients: Optional list of CC recipient email addresses
        bcc_recipients: Optional list of BCC recipient email addresses
        subject: Optional draft subject
        body: Optional draft body content
        body_content_type: Body format for the draft (`text` or `html`)

    Returns:
        Draft metadata containing the created draft ID plus a shaped draft message object.
    """

    logger.info(
        "create_email_draft called: draft_type=%s, email_id=%s, subject=%s",
        draft_type,
        email_id,
        subject,
    )

    try:
        normalized_draft_type = _normalize_draft_type(draft_type)
        normalized_body_type = _normalize_body_content_type(body_content_type)

        to_objects = _build_recipient_objects(to_recipients)
        cc_objects = _build_recipient_objects(cc_recipients)
        bcc_objects = _build_recipient_objects(bcc_recipients)

        payload = _build_message_update_payload(
            subject=subject,
            body=body,
            body_content_type=normalized_body_type,
            to_recipients=to_objects,
            cc_recipients=cc_objects,
            bcc_recipients=bcc_objects,
        )

        reply_to_message_id: str | None = None
        raw: dict[str, Any] | None = None

        if normalized_draft_type == "new":
            if not any(
                payload.get(key)
                for key in ("toRecipients", "ccRecipients", "bccRecipients")
            ):
                raise ValueError(
                    "New message drafts require at least one recipient in to_recipients, cc_recipients, or bcc_recipients"
                )
            raw = graph.request("POST", "/me/messages", json=payload)
        else:
            if not email_id:
                raise ValueError("email_id is required for reply and reply_all drafts")

            reply_to_message_id = email_id
            action = (
                "createReplyAll"
                if normalized_draft_type == "reply_all"
                else "createReply"
            )
            raw = graph.request("POST", f"/me/messages/{email_id}/{action}")

            if not raw or not raw.get("id"):
                raise ValueError("Draft could not be created")

            if payload:
                raw = _patch_email_message(raw["id"], payload)

        if not raw or not raw.get("id"):
            raise ValueError("Draft could not be created")

        result = {
            "status": "draft_created",
            "draft_type": normalized_draft_type,
            "draft_id": raw["id"],
            "draft": _shape_email_draft(raw),
        }
        if reply_to_message_id is not None:
            result["reply_to_message_id"] = reply_to_message_id

        return result
    except Exception as e:
        logger.error(
            "create_email_draft failed: draft_type=%s, email_id=%s, error=%s",
            draft_type,
            email_id,
            str(e),
            exc_info=True,
        )
        raise


def _resolve_mail_folder(folder: str, include_hidden: bool = False) -> str:
    destination = folder.strip()
    if not destination:
        raise ValueError("Destination folder cannot be empty")
    if destination.casefold() in FOLDERS:
        return FOLDERS[destination.casefold()]
    return _find_mail_folder(destination, include_hidden=include_hidden)["id"]


def _shape_email_management_result(raw: dict[str, Any]) -> dict[str, Any]:
    email = shape_email_summary(raw)

    if raw.get("categories") is not None:
        email["categories"] = raw.get("categories", [])

    flag_status = raw.get("flag", {}).get("flagStatus")
    if flag_status:
        email["flag_status"] = flag_status

    return email


def _looks_like_invite_message(raw: dict[str, Any]) -> bool:
    meeting_message_type = raw.get("meetingMessageType")
    if isinstance(meeting_message_type, str) and meeting_message_type != "none":
        return True

    odata_type = raw.get("@odata.type", "")
    return isinstance(odata_type, str) and "eventMessage" in odata_type


def _shape_invite_message(
    raw: dict[str, Any], include_body: bool = False
) -> dict[str, Any]:
    shaped = shape_email_detail(raw) if include_body else shape_email_summary(raw)
    shaped["kind"] = "invite_message"

    meeting_message_type = raw.get("meetingMessageType")
    if meeting_message_type:
        shaped["meeting_message_type"] = meeting_message_type

    if raw.get("responseRequested") is not None:
        shaped["response_requested"] = raw["responseRequested"]

    if raw.get("allowNewTimeProposals") is not None:
        shaped["allow_new_time_proposals"] = raw["allowNewTimeProposals"]

    if raw.get("isOutOfDate") is not None:
        shaped["is_out_of_date"] = raw["isOutOfDate"]

    if raw.get("startDateTime"):
        shaped["start"] = raw["startDateTime"]

    if raw.get("endDateTime"):
        shaped["end"] = raw["endDateTime"]

    location = compact_location(raw.get("location"))
    if location:
        shaped["location"] = location

    if raw.get("webLink"):
        shaped["web_url"] = raw["webLink"]

    event = raw.get("event")
    if isinstance(event, dict) and event.get("id"):
        shaped["event"] = (
            shape_event_detail(event) if include_body else shape_event_summary(event)
        )

    return shaped


def _get_invite_message(
    invite_message_id: str, *, include_body: bool = False, expand_event: bool = False
) -> dict[str, Any]:
    params: dict[str, Any] = {}
    if expand_event:
        params["$expand"] = "microsoft.graph.eventMessage/event"

    raw = graph.request("GET", f"/me/messages/{invite_message_id}", params=params)
    if not raw:
        raise ValueError(f"Invite message with ID {invite_message_id} not found")
    if not _looks_like_invite_message(raw):
        raise ValueError(
            f"Message with ID {invite_message_id} is not an invite message"
        )
    return raw


def _resolve_event_response_endpoint(response: str) -> tuple[str, str]:
    normalized_response = response.strip().casefold()
    action_map = {
        "accept": "accept",
        "accepted": "accept",
        "decline": "decline",
        "declined": "decline",
        "tentative": "tentativelyAccept",
        "tentatively_accept": "tentativelyAccept",
        "tentativelyaccept": "tentativelyAccept",
    }
    endpoint = action_map.get(normalized_response)
    if endpoint is None:
        raise ValueError("response must be one of: accept, decline, tentative")
    normalized = "tentative" if endpoint == "tentativelyAccept" else endpoint
    return endpoint, normalized


def _patch_email_message(email_id: str, payload: dict[str, Any]) -> dict[str, Any]:
    path = f"/me/messages/{email_id}"
    raw = graph.request("PATCH", path, json=payload)
    if not raw:
        raw = graph.request("GET", path)
    if not raw:
        raise ValueError(f"Email with ID {email_id} not found")
    return raw


def _move_email_message(
    email_id: str, destination_folder: str
) -> tuple[dict[str, Any], str]:
    resolved_destination = _resolve_mail_folder(destination_folder)
    raw = graph.request(
        "POST",
        f"/me/messages/{email_id}/move",
        json={"destinationId": resolved_destination},
    )
    if not raw:
        raise ValueError(f"Email with ID {email_id} could not be moved")
    return raw, resolved_destination


def _delete_email_message(email_id: str) -> dict[str, Any]:
    try:
        graph.request("DELETE", f"/me/messages/{email_id}")
        return {"status": "deleted", "email_id": email_id, "resource": "message"}
    except httpx.HTTPStatusError as e:
        status_code = e.response.status_code if e.response is not None else None
        if status_code not in {400, 404}:
            raise

    graph.request("DELETE", f"/me/events/{email_id}")
    return {"status": "deleted", "email_id": email_id, "resource": "event"}


def _delete_invite_message(invite_message_id: str) -> dict[str, Any]:
    graph.request("DELETE", f"/me/messages/{invite_message_id}")
    return {
        "status": "deleted",
        "invite_message_id": invite_message_id,
        "resource": "eventMessage",
    }


def _list_message_summaries(folder: str, limit: int) -> list[dict[str, Any]]:
    params = {
        "$top": min(limit, 100),
        "$select": MESSAGE_SUMMARY_SELECT_FIELDS,
        "$orderby": "receivedDateTime desc",
    }
    return list(
        graph.request_paginated(
            f"/me/mailFolders/{folder}/messages",
            params=params,
            limit=limit,
        )
    )


def _hydrate_invite_messages_from_summaries(
    raw_messages: list[dict[str, Any]],
    *,
    limit: int,
) -> list[dict[str, Any]]:
    invite_messages: list[dict[str, Any]] = []

    for summary in raw_messages:
        try:
            raw = _get_invite_message(
                summary["id"],
                include_body=False,
                expand_event=True,
            )
        except ValueError:
            continue
        except Exception as e:
            logger.warning(
                "invite message probe failed for message_id=%s: %s",
                summary.get("id"),
                str(e),
            )
            continue

        invite_messages.append(raw)
        if len(invite_messages) >= limit:
            break

    return invite_messages


def _run_email_management_action(
    email_id: str,
    action: str,
    destination_folder: str | None = None,
    categories: list[str] | None = None,
) -> dict[str, Any]:
    if action == "mark_read":
        raw = _patch_email_message(email_id, {"isRead": True})
        return {"status": "updated", "email": _shape_email_management_result(raw)}
    if action == "mark_unread":
        raw = _patch_email_message(email_id, {"isRead": False})
        return {"status": "updated", "email": _shape_email_management_result(raw)}
    if action == "move":
        if destination_folder is None:
            raise ValueError("destination_folder is required for move")
        raw, resolved_destination = _move_email_message(email_id, destination_folder)
        return {
            "status": "moved",
            "destination_folder": resolved_destination,
            "email": _shape_email_management_result(raw),
        }
    if action == "archive":
        raw, resolved_destination = _move_email_message(email_id, "archive")
        return {
            "status": "moved",
            "destination_folder": resolved_destination,
            "email": _shape_email_management_result(raw),
        }
    if action == "delete":
        return _delete_email_message(email_id)
    if action == "set_categories":
        if categories is None:
            raise ValueError("categories are required for set_categories")
        raw = _patch_email_message(email_id, {"categories": categories})
        return {
            "status": "updated",
            "categories": raw.get("categories", []),
            "email": _shape_email_management_result(raw),
        }
    raise ValueError(f"Unsupported action '{action}'")


@mcp.tool
def mark_email_read(email_id: str, is_read: bool = True) -> dict[str, Any]:
    """Mark an email as read or unread.

    Args:
        email_id: Unique identifier of the email to update.
        is_read: True to mark as read, False to mark as unread.

    Returns:
        Update result containing the refreshed compact email summary.
    """

    logger.info(f"mark_email_read called: email_id={email_id}, is_read={is_read}")

    try:
        raw = _patch_email_message(email_id, {"isRead": is_read})
        return {
            "status": "updated",
            "email": _shape_email_management_result(raw),
        }
    except Exception as e:
        logger.error(
            f"mark_email_read failed for email_id={email_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def set_email_categories(email_id: str, categories: list[str]) -> dict[str, Any]:
    """Replace the category labels on an email.

    Args:
        email_id: Unique identifier of the email to update.
        categories: Complete list of categories that should remain on the email.

    Returns:
        Update result containing the refreshed compact email summary.
    """

    logger.info(
        f"set_email_categories called: email_id={email_id}, categories={categories}"
    )

    try:
        raw = _patch_email_message(email_id, {"categories": categories})
        return {
            "status": "updated",
            "categories": raw.get("categories", []),
            "email": _shape_email_management_result(raw),
        }
    except Exception as e:
        logger.error(
            f"set_email_categories failed for email_id={email_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def move_email(email_id: str, destination_folder: str) -> dict[str, Any]:
    """Move an email into another mailbox folder.

    Args:
        email_id: Unique identifier of the email to move.
        destination_folder: Folder alias, ID, display name, or slash-delimited path.
            Supported aliases include inbox, sent, drafts, deleted, junk, and archive.

    Returns:
        Move result containing the refreshed compact email summary.
    """

    logger.info(
        f"move_email called: email_id={email_id}, destination_folder={destination_folder}"
    )

    try:
        raw, resolved_destination = _move_email_message(email_id, destination_folder)
        return {
            "status": "moved",
            "destination_folder": resolved_destination,
            "email": _shape_email_management_result(raw),
        }
    except Exception as e:
        logger.error(
            f"move_email failed for email_id={email_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def archive_email(email_id: str) -> dict[str, Any]:
    """Move an email into the archive folder."""

    logger.info(f"archive_email called: email_id={email_id}")

    try:
        raw, resolved_destination = _move_email_message(email_id, "archive")
        return {
            "status": "moved",
            "destination_folder": resolved_destination,
            "email": _shape_email_management_result(raw),
        }
    except Exception as e:
        logger.error(
            f"archive_email failed for email_id={email_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def delete_email(email_id: str) -> dict[str, Any]:
    """Delete an email from the mailbox."""

    logger.info(f"delete_email called: email_id={email_id}")

    try:
        return _delete_email_message(email_id)
    except Exception as e:
        logger.error(
            f"delete_email failed for email_id={email_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def bulk_manage_emails(
    email_ids: list[str],
    action: str,
    destination_folder: str | None = None,
    categories: list[str] | None = None,
) -> dict[str, Any]:
    """Apply one inbox-management action to many emails.

    Supported actions:
        - mark_read
        - mark_unread
        - move
        - archive
        - delete
        - set_categories

    Args:
        email_ids: List of email IDs to update.
        action: The action to apply to every email ID.
        destination_folder: Required when action is "move".
        categories: Required when action is "set_categories".

    Returns:
        Summary including per-email results and partial failures.
    """

    logger.info(
        "bulk_manage_emails called: action=%s, count=%s, destination_folder=%s",
        action,
        len(email_ids),
        destination_folder,
    )

    supported_actions = {
        "mark_read",
        "mark_unread",
        "move",
        "archive",
        "delete",
        "set_categories",
    }
    if action not in supported_actions:
        raise ValueError(
            f"Unsupported action '{action}'. Supported actions: {sorted(supported_actions)}"
        )

    results: list[dict[str, Any]] = []
    succeeded = 0

    for email_id in email_ids:
        try:
            result = _run_email_management_action(
                email_id=email_id,
                action=action,
                destination_folder=destination_folder,
                categories=categories,
            )
            result["email_id"] = email_id
            results.append(result)
            succeeded += 1
        except Exception as e:
            logger.error(
                "bulk_manage_emails failed for action=%s email_id=%s: %s",
                action,
                email_id,
                str(e),
                exc_info=True,
            )
            results.append(
                {
                    "email_id": email_id,
                    "status": "failed",
                    "error": str(e),
                }
            )

    return {
        "action": action,
        "requested": len(email_ids),
        "succeeded": succeeded,
        "failed": len(email_ids) - succeeded,
        "results": results,
    }


@mcp.tool
def list_invite_messages(
    limit: int = 20, folder: str = "inbox"
) -> list[dict[str, Any]]:
    """List meeting invite-style messages from a mailbox folder.

    This surfaces meeting requests, cancellations, and meeting response messages
    that live in the mailbox as eventMessage objects.

    Args:
        limit: Maximum invite messages to return.
        folder: Mail folder alias or ID to scan. Defaults to inbox.

    Returns:
        Compact invite-message summaries, including the associated event when Graph
        has already materialized it.
    """

    logger.info("list_invite_messages called: limit=%s, folder=%s", limit, folder)

    resolved_folder = _resolve_mail_folder(folder)
    fetch_limit = max(limit * 5, min(limit + 20, 100))

    try:
        raw_messages = _list_message_summaries(
            resolved_folder,
            fetch_limit,
        )
        invite_messages = _hydrate_invite_messages_from_summaries(
            raw_messages,
            limit=limit,
        )
        invite_messages = [_shape_invite_message(raw) for raw in invite_messages]
        return invite_messages[:limit]
    except Exception as e:
        logger.error(
            "list_invite_messages failed for folder=%s: %s",
            folder,
            str(e),
            exc_info=True,
        )
        raise


@mcp.tool
def delete_invite_message(invite_message_id: str) -> dict[str, Any]:
    """Delete a meeting invite-style message from the mailbox."""

    logger.info("delete_invite_message called: invite_message_id=%s", invite_message_id)

    try:
        return _delete_invite_message(invite_message_id)
    except Exception as e:
        logger.error(
            "delete_invite_message failed for invite_message_id=%s: %s",
            invite_message_id,
            str(e),
            exc_info=True,
        )
        raise


@mcp.tool
def list_events(
    days_ahead: int = 7,
    days_back: int = 0,
    max_body_length: int = 500,
    include_details: bool = False,
    response_profile: str = "auto",
) -> list[dict[str, Any]]:
    """List calendar events within a specified date range.

    Retrieves calendar events including recurring event instances. Events are ordered by start time.
    Use this to check upcoming meetings, find events in a date range, or get calendar overview.

    Args:
        days_ahead: Number of days into the future to search (default 7)
        days_back: Number of days into the past to search (default 0 = today onwards)
        include_details: Whether to include full event details like body, attendees, online meeting info
        response_profile: Response shaping profile ("auto", "legacy", or "assistant"). "auto" defers to MICROSOFT_MCP_RESPONSE_PROFILE env var.

    Returns:
        List of calendar event objects containing:
        - Basic info: id, subject, start/end times, location, organizer (note: All times are in UTC time zone and may require conversion)
        - Details (if include_details=True): body, attendees list, recurrence info, online meeting links
        - Recurring events: individual instances with seriesMasterId for the recurring series

    Examples:
        - list_events() - Get next 7 days of events
        - list_events(days_ahead=30) - Get next 30 days of events
        - list_events(days_back=7, days_ahead=7) - Get events from past week to next week
        - list_events(include_details=False) - Get basic event info only for faster and shorter response
    """
    profile = get_response_profile(response_profile)
    if profile == "assistant":
        include_details = False

    logger.info(
        f"list_events called: days_ahead={days_ahead}, days_back={days_back}, include_details={include_details}, profile={profile}"
    )

    try:
        now = dt.datetime.now(dt.timezone.utc)
        start = (now - dt.timedelta(days=days_back)).isoformat()
        end = (now + dt.timedelta(days=days_ahead)).isoformat()

        params = {
            "startDateTime": start,
            "endDateTime": end,
            "$orderby": "start/dateTime",
            "$top": 100,
        }

        if include_details:
            params["$select"] = (
                "id,subject,start,end,location,body,attendees,organizer,isAllDay,recurrence,onlineMeeting,seriesMasterId"
            )
        else:
            params["$select"] = "id,subject,start,end,location,organizer,seriesMasterId"

        # Use calendarView to get recurring event instances
        raw_events = list(graph.request_paginated("/me/calendarView", params=params))

        if include_details:
            # truncate the body content if it exceeds max_body_length
            for event in raw_events:
                if "body" in event:
                    if (
                        "content" in event["body"]
                        and len(event["body"]["content"]) > max_body_length
                    ):
                        event["body"]["content"] = (
                            event["body"]["content"][:max_body_length] + "..."
                        )
            events = [shape_event_detail(e) for e in raw_events]
        else:
            events = [shape_event_summary(e) for e in raw_events]

        logger.info(
            f"list_events successful: retrieved {len(events)} events from {start} to {end}"
        )
        return events
    except Exception as e:
        logger.error(f"list_events failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_event(event_id: str) -> dict[str, Any]:
    """Get complete details for a specific calendar event by its ID.

    Retrieves full event information including attendees, recurrence, online meeting details, etc.
    Use this when you need complete event details after finding events with list_events or search_events.

    Args:
        event_id: Unique identifier of the calendar event (get from list_events or search results)

    Returns:
        Complete event object containing:
        - Basic info: id, subject, start/end times, location, isAllDay, organizer
        - Attendees: list with names, email addresses, response status
        - Body: event description/notes
        - Recurrence: pattern info for recurring events
        - Online meeting: Teams/Zoom links and dial-in info if applicable
        - Categories: event categorization tags

    Examples:
        - get_event("AAMkAD...") - Get full details for a specific event
        - Use after list_events() to get complete info about interesting events
    """
    logger.info(f"get_event called: event_id={event_id}")

    try:
        raw = graph.request("GET", f"/me/events/{event_id}")
        if not raw:
            logger.error(f"get_event failed: Event with ID {event_id} not found")
            raise ValueError(f"Event with ID {event_id} not found")

        logger.info(f"get_event successful: retrieved event {event_id}")
        return shape_event_detail(raw)
    except Exception as e:
        logger.error(
            f"get_event failed for event_id={event_id}: {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def rsvp_to_event(
    event_id: str,
    response: str,
    comment: str | None = None,
    send_response: bool = False,
) -> dict[str, Any]:
    """RSVP to a calendar event without emailing the organizer by default.

    Args:
        event_id: Unique identifier of the calendar event.
        response: One of "accept", "decline", or "tentative".
        comment: Optional organizer-facing message.
        send_response: Whether to email the organizer. Defaults to False.

    Returns:
        Confirmation of the RSVP action that was submitted.
    """

    logger.info(
        "rsvp_to_event called: event_id=%s, response=%s, send_response=%s",
        event_id,
        response,
        send_response,
    )

    endpoint, normalized_response = _resolve_event_response_endpoint(response)

    payload = {
        "comment": comment,
        "sendResponse": send_response,
    }

    try:
        graph.request(
            "POST",
            f"/me/events/{event_id}/{endpoint}",
            json=payload,
        )
        return {
            "status": "responded",
            "event_id": event_id,
            "response": normalized_response,
            "send_response": send_response,
        }
    except Exception as e:
        logger.error(
            f"rsvp_to_event failed for event_id={event_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def rsvp_to_invite_message(
    invite_message_id: str,
    response: str,
    comment: str | None = None,
    send_response: bool = False,
) -> dict[str, Any]:
    """RSVP to a meeting invite message via its associated calendar event.

    Args:
        invite_message_id: Unique identifier of the invite message in the mailbox.
        response: One of "accept", "decline", or "tentative".
        comment: Optional organizer-facing message.
        send_response: Whether to email the organizer. Defaults to False.

    Returns:
        Confirmation of the RSVP action that was submitted.
    """

    logger.info(
        "rsvp_to_invite_message called: invite_message_id=%s, response=%s, send_response=%s",
        invite_message_id,
        response,
        send_response,
    )

    endpoint, normalized_response = _resolve_event_response_endpoint(response)

    try:
        invite_message = _get_invite_message(
            invite_message_id,
            include_body=True,
            expand_event=True,
        )
        event = invite_message.get("event")
        if not isinstance(event, dict) or not event.get("id"):
            raise ValueError(
                f"Invite message with ID {invite_message_id} has no associated event"
            )

        graph.request(
            "POST",
            f"/me/events/{event['id']}/{endpoint}",
            json={"comment": comment, "sendResponse": send_response},
        )
        return {
            "status": "responded",
            "invite_message_id": invite_message_id,
            "event_id": event["id"],
            "meeting_message_type": invite_message.get("meetingMessageType", "none"),
            "response": normalized_response,
            "send_response": send_response,
        }
    except Exception as e:
        logger.error(
            "rsvp_to_invite_message failed for invite_message_id=%s: %s",
            invite_message_id,
            str(e),
            exc_info=True,
        )
        raise


@mcp.tool
def check_availability(
    start: str,
    end: str,
    attendees: str | list[str] | None = None,
) -> dict[str, Any]:
    """Check calendar availability for the user and optionally other attendees within a time range.

    Determines free/busy status to help schedule meetings. Shows when people are available,
    busy, or tentatively booked. Useful for finding meeting times that work for everyone.
    All times are in UTC time zone and may require conversion.

    Args:
        start: Start time in ISO format (e.g., "2024-09-02T09:00:00Z" or "2024-09-02T09:00:00")
        end: End time in ISO format
        attendees: Email address(es) of other people to check (optional). Can be single email or list

    Returns:
        Availability information containing:
        - schedules: Array of availability data for each person checked
        - freeBusyViewType: Type of view (e.g., "freeBusy")
        - For each person: email, availability intervals showing free/busy/tentative status
        - availabilityView: Numeric representation of availability (0=free, 1=tentative, 2=busy),
          each number represents a 30-minute interval within the specified time range, starting from the start time.

    Examples:
        - check_availability("2024-09-02T14:00:00Z", "2024-09-02T15:00:00Z") - Check your availability
        - check_availability(start, end, "colleague@company.com") - Check you + one person
        - check_availability(start, end, ["person1@co.com", "person2@co.com"]) - Check multiple people
    """
    logger.info(
        f"check_availability called: start={start}, end={end}, attendees={attendees}"
    )

    try:
        me_info = graph.request("GET", "/me")
        if not me_info or "mail" not in me_info:
            logger.error("check_availability failed: could not get user email address")
            raise ValueError("Failed to get user email address")

        schedules = [me_info["mail"]]
        if attendees:
            attendees_list = [attendees] if isinstance(attendees, str) else attendees
            schedules.extend(attendees_list)
            logger.info(f"check_availability: checking {len(schedules)} schedules")

        payload = {
            "schedules": schedules,
            "startTime": {"dateTime": start, "timeZone": "UTC"},
            "endTime": {"dateTime": end, "timeZone": "UTC"},
            "availabilityViewInterval": 30,
        }

        raw = graph.request("POST", "/me/calendar/getSchedule", json=payload)
        if not raw:
            logger.error("check_availability failed: no response from server")
            raise ValueError("Failed to check availability")

        logger.info(
            f"check_availability successful: checked availability for {len(schedules)} schedules"
        )

        # Shape into compact assistant-facing format
        raw_schedules = raw.get("value", [])
        participants = []
        for sched in raw_schedules:
            entry: dict[str, Any] = {
                "email": sched.get("scheduleId", ""),
                "availability": sched.get("availabilityView", ""),
            }
            items = sched.get("scheduleItems", [])
            if items:
                entry["slots"] = [
                    {
                        "status": s.get("status"),
                        "subject": s.get("subject"),
                        "start": s.get("start"),
                        "end": s.get("end"),
                    }
                    for s in items
                ]
            wh = sched.get("workingHours")
            if wh:
                entry["working_hours"] = wh
            participants.append(entry)

        return {
            "participants": participants,
            "time_range": {"start": start, "end": end},
        }
    except Exception as e:
        logger.error(f"check_availability failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def list_contacts(
    limit: int = 50, response_profile: str = "auto"
) -> list[dict[str, Any]]:
    """List contacts from the user's address book.

    Retrieves personal contacts with names, email addresses, phone numbers, and other details.
    Use this to find contact information, get email addresses for sending messages, or browse contacts.

    Args:
        limit: Maximum number of contacts to retrieve (1-100, defaults to 50)
        response_profile: Response shaping profile ("auto", "legacy", or "assistant"). "auto" defers to MICROSOFT_MCP_RESPONSE_PROFILE env var.

    Returns:
        List of contact objects containing:
        - Names: givenName, surname, displayName, nickname
        - Email addresses: array of email addresses with labels
        - Phone numbers: businessPhones, homePhones, mobilePhone
        - Addresses: business and home addresses
        - Other: jobTitle, companyName, birthday, notes

    Examples:
        - list_contacts() - Get first 50 contacts
        - list_contacts(limit=100) - Get more contacts
        - Use to find someone's email before sending messages
    """
    profile = get_response_profile(response_profile)

    logger.info(f"list_contacts called: limit={limit}, profile={profile}")

    try:
        params = {"$top": min(limit, 100)}

        raw_contacts = list(
            graph.request_paginated("/me/contacts", params=params, limit=limit)
        )

        contacts = [shape_contact_summary(c) for c in raw_contacts]

        logger.info(f"list_contacts successful: retrieved {len(contacts)} contacts")
        return contacts
    except Exception as e:
        logger.error(f"list_contacts failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_contact(contact_id: str) -> dict[str, Any]:
    """Get detailed information for a specific contact by ID.

    Retrieves complete contact details including all phone numbers, email addresses,
    postal addresses, and personal information. Use after finding contacts with list_contacts
    or search_contacts when you need full contact details.

    Args:
        contact_id: Unique identifier of the contact (get from list_contacts or search results)

    Returns:
        Complete contact object containing:
        - Names: givenName, surname, displayName, nickname, title
        - Communications: emailAddresses array, businessPhones, homePhones, mobilePhone
        - Addresses: businessAddress, homeAddress with street, city, state, country, postalCode
        - Professional: jobTitle, companyName, department, officeLocation
        - Personal: birthday, spouseName, children, personalNotes
        - Categories: assigned category tags

    Examples:
        - get_contact("AAMkAD...") - Get full contact details
        - Use to get complete info after finding contact in search results
    """
    logger.info(f"get_contact called: contact_id={contact_id}")

    try:
        raw = graph.request("GET", f"/me/contacts/{contact_id}")
        if not raw:
            logger.error(f"get_contact failed: Contact with ID {contact_id} not found")
            raise ValueError(f"Contact with ID {contact_id} not found")

        logger.info(f"get_contact successful: retrieved contact {contact_id}")
        return shape_contact_detail(raw)
    except Exception as e:
        logger.error(
            f"get_contact failed for contact_id={contact_id}: {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def list_files(path: str = "/", limit: int = 50) -> list[dict[str, Any]]:
    """List files and folders in OneDrive at a specified path.

    Browse OneDrive contents to see what files and folders are available. Use this to navigate
    the file system, find documents, or get an overview of stored content.

    Args:
        path: OneDrive path to browse (default "/" for root). Use forward slashes like "Documents/Projects"
        limit: Maximum number of items to retrieve (1-100, defaults to 50)

    Returns:
        List of file/folder objects containing:
        - Basic info: id, name, type (file/folder), size (bytes), modified (timestamp)
        - Download info: download_url for direct file access (for files only)
        - Use 'type' field to distinguish between "file" and "folder"
        - Size is 0 for folders

    Examples:
        - list_files() - List root directory contents
        - list_files(path="Documents") - List contents of Documents folder
        - list_files(path="Pictures/Vacation", limit=100) - Browse specific folder with more results
        - Check 'type' field to see if item is file or folder for navigation
    """
    logger.info(f"list_files called: path={path}, limit={limit}")

    try:
        endpoint = (
            "/me/drive/root/children"
            if path == "/"
            else f"/me/drive/root:/{path}:/children"
        )
        params = {
            "$top": min(limit, 100),
            "$select": "id,name,size,lastModifiedDateTime,folder,file,@microsoft.graph.downloadUrl",
        }

        items = list(graph.request_paginated(endpoint, params=params, limit=limit))

        result = [
            {
                "id": item["id"],
                "name": item["name"],
                "type": "folder" if "folder" in item else "file",
                "size": item.get("size", 0),
                "modified": item.get("lastModifiedDateTime"),
                "download_url": item.get("@microsoft.graph.downloadUrl"),
            }
            for item in items
        ]

        logger.info(
            f"list_files successful: retrieved {len(result)} items from path {path}"
        )
        return result
    except Exception as e:
        logger.error(f"list_files failed for path={path}: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_file(file_id: str, download_path: str) -> dict[str, Any]:
    """Download a file from OneDrive to a local file path.

    Downloads any file from OneDrive to your local computer. Use this after finding files
    with list_files or search_files when you need to access the actual file content.

    Args:
        file_id: Unique identifier of the file (get from list_files or search_files results)
        download_path: Local file path where to save the downloaded file (e.g., "/tmp/document.pdf")

    Returns:
        Download result information:
        - path: Local path where file was saved
        - name: Original filename from OneDrive
        - size_mb: File size in megabytes (rounded to 2 decimals)
        - mime_type: File MIME type (e.g., "application/pdf", "image/jpeg")

    Examples:
        - get_file("AAMkAD...", "/tmp/report.pdf") - Download specific file
        - get_file(file_id, "~/Downloads/document.docx") - Download to Downloads folder
        - Use file_id from list_files() or search_files() results
    """
    logger.info(f"get_file called: file_id={file_id}, download_path={download_path}")

    try:
        import subprocess

        metadata = graph.request("GET", f"/me/drive/items/{file_id}")
        if not metadata:
            logger.error(f"get_file failed: File with ID {file_id} not found")
            raise ValueError(f"File with ID {file_id} not found")

        download_url = metadata.get("@microsoft.graph.downloadUrl")
        if not download_url:
            logger.error(
                f"get_file failed: No download URL available for file {file_id}"
            )
            raise ValueError("No download URL available for this file")

        try:
            subprocess.run(
                ["curl", "-L", "-o", download_path, download_url],
                check=True,
                capture_output=True,
            )

            result = {
                "path": download_path,
                "name": metadata.get("name", "unknown"),
                "size_mb": round(metadata.get("size", 0) / (1024 * 1024), 2),
                "mime_type": (
                    metadata.get("file", {}).get("mimeType") if metadata else None
                ),
            }

            logger.info(
                f"get_file successful: downloaded {result['name']} ({result['size_mb']} MB) to {download_path}"
            )
            return result
        except subprocess.CalledProcessError as e:
            logger.error(f"get_file failed: curl command failed - {e.stderr.decode()}")
            raise RuntimeError(f"Failed to download file: {e.stderr.decode()}")
    except Exception as e:
        logger.error(f"get_file failed for file_id={file_id}: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_attachment(email_id: str, attachment_id: str, save_path: str) -> dict[str, Any]:
    """Download an email attachment to a local file path.

    Downloads attachments from emails (documents, images, etc.) to your local computer.
    Use this after finding emails with attachments via get_email() when you need the actual attachment files.

    Args:
        email_id: Unique identifier of the email containing the attachment
        attachment_id: Unique identifier of the specific attachment (from get_email attachments list)
        save_path: Local file path where to save the attachment (e.g., "/tmp/attachment.pdf")

    Returns:
        Attachment download information:
        - name: Original filename of the attachment
        - content_type: MIME type (e.g., "application/pdf", "image/jpeg", "text/plain")
        - size: File size in bytes
        - saved_to: Absolute local path where file was saved

    Examples:
        - get_attachment(email_id, attachment_id, "/tmp/document.pdf") - Download specific attachment
        - get_attachment(email_id, attachment_id, "~/Downloads/image.jpg") - Save to Downloads
        - First use get_email() to see what attachments are available, then download specific ones
    """
    logger.info(
        f"get_attachment called: email_id={email_id}, attachment_id={attachment_id}, save_path={save_path}"
    )

    try:
        result = graph.request(
            "GET", f"/me/messages/{email_id}/attachments/{attachment_id}"
        )

        if not result:
            logger.error(
                f"get_attachment failed: Attachment {attachment_id} not found in email {email_id}"
            )
            raise ValueError("Attachment not found")

        if "contentBytes" not in result:
            logger.error(
                f"get_attachment failed: Attachment content not available for {attachment_id}"
            )
            raise ValueError("Attachment content not available")

        # Save attachment to file
        path = pl.Path(save_path).expanduser().resolve()
        path.parent.mkdir(parents=True, exist_ok=True)
        content_bytes = base64.b64decode(result["contentBytes"])
        path.write_bytes(content_bytes)

        attachment_result = {
            "name": result.get("name", "unknown"),
            "content_type": result.get("contentType", "application/octet-stream"),
            "size": result.get("size", 0),
            "saved_to": str(path),
        }

        logger.info(
            f"get_attachment successful: saved {attachment_result['name']} ({attachment_result['size']} bytes) to {save_path}"
        )
        return attachment_result
    except Exception as e:
        logger.error(
            f"get_attachment failed for email_id={email_id}, attachment_id={attachment_id}: {str(e)}",
            exc_info=True,
        )
        raise


def _analyze_search_error(error: Exception, request_payload: dict) -> str:
    """Analyze Microsoft Graph Search API errors and provide helpful diagnostics."""

    error_msg = str(error)

    if "400 Bad Request" in error_msg:
        entity_types = request_payload.get("requests", [{}])[0].get("entityTypes", [])
        query_string = (
            request_payload.get("requests", [{}])[0]
            .get("query", {})
            .get("queryString", "")
        )

        suggestions = []

        # Check for problematic entity type combinations
        if len(entity_types) > 1:
            if "event" in entity_types:
                suggestions.append("'event' cannot be combined with other entity types")
            if "person" in entity_types:
                suggestions.append(
                    "'person' cannot be combined with other entity types"
                )
            if "chatMessage" in entity_types and any(
                et in ["driveItem", "site", "list", "listItem"] for et in entity_types
            ):
                suggestions.append(
                    "'chatMessage' cannot be combined with file-related entity types"
                )

        # Check for empty or invalid query
        if not query_string or query_string.isspace():
            suggestions.append(
                "Query string cannot be empty (use '*' for wildcard search)"
            )

        # Check for unsupported entity types
        valid_entities = {
            "message",
            "event",
            "driveItem",
            "site",
            "drive",
            "chatMessage",
            "person",
            "list",
            "listItem",
        }
        invalid_entities = [et for et in entity_types if et not in valid_entities]
        if invalid_entities:
            suggestions.append(f"Invalid entity types: {invalid_entities}")

        if suggestions:
            return f"Bad Request - Possible issues: {'; '.join(suggestions)}"
        else:
            return "Bad Request - Check entity type combinations and query format"

    elif "401 Unauthorized" in error_msg:
        return "Authentication failed - token may be expired or invalid"
    elif "403 Forbidden" in error_msg:
        return "Insufficient permissions - check that required scopes are granted"
    elif "404 Not Found" in error_msg:
        return "Search endpoint not found - API may not be available in this tenant"
    elif "429 Too Many Requests" in error_msg:
        return "Rate limited - too many requests, please retry later"
    else:
        return f"Unexpected error: {error_msg}"


@mcp.tool
def unified_search(
    query: str,
    entity_types: list[str] | None = None,
    limit: int = 50,
    kql_filters: str | None = None,
    include_body: bool = False,
    body_max_length: int = 1000,
) -> dict[str, Any]:
    """Universal search across Microsoft 365 content using the Microsoft Search API.

    Searches across emails, calendar events, files, SharePoint content, Teams messages, and people
    using a single unified API. Supports advanced KQL (Keyword Query Language) filters for precise
    results and efficient token usage.

    Args:
        query: Search terms or KQL query string (e.g., "project meeting", "from:john@company.com budget")
        entity_types: List of content types to search. Options:
            - "message": Outlook emails
            - "event": Calendar events
            - "driveItem": OneDrive/SharePoint files and folders
            - "list": SharePoint lists
            - "listItem": SharePoint list items
            - "site": SharePoint sites
            - "drive": OneDrive/SharePoint drives
            - "chatMessage": Teams chat and channel messages
            - "person": People in your organization
            If None, searches all supported types for comprehensive results.
        limit: Maximum number of results to return (1-100, defaults to 50)
        kql_filters: Additional KQL filters for precise search. Examples:
            - "from:john@company.com" - Emails from specific sender
            - "sent>=2024-01-01" - Items after specific date
            - "to:manager@company.com" - Emails to specific recipient
            - "IsMentioned:true" - Teams messages where you're mentioned
            - "filetype:pdf" - Only PDF files
            - "author:\"John Smith\"" - Content authored by John Smith
        include_body: Whether to include full body/content (increases response size)
        body_max_length: Maximum characters for body content when included (default 1000)

    Returns:
        Search results containing:
        - summary: Search statistics (total results, entity type breakdown)
        - results: Array of matching items containing all available fields from Microsoft Graph API:
            - All original fields from the API resource
            - entity_type: Type of content (message, event, driveItem, etc.)
            - search_rank: Search relevance score
            - search_summary: Brief content preview from search
            - conversation_url: Deep link for messages (when available)
            - body: Full content if include_body=True (converted to markdown for HTML content)

    Examples:
        - unified_search("budget meeting") - Find all content about budget meetings
        - unified_search("project alpha", ["message", "chatMessage"]) - Search emails and Teams messages only
        - unified_search("quarterly report", kql_filters="filetype:pdf OR filetype:docx") - Find documents only
        - unified_search("", kql_filters="from:manager@company.com sent>=2024-01-01") - Recent emails from manager
        - unified_search("presentation", kql_filters="author:\"Sarah Wilson\"") - Sarah's presentations
        - unified_search("important", entity_types=["message"], kql_filters="IsMentioned:true") - Important emails mentioning you

    Note: KQL filters allow precise control over search scope and can significantly improve relevance.
    Results now include all available fields from the Microsoft Graph API for maximum information.
    """
    logger.info(
        f"unified_search called: query='{query}', entity_types={entity_types}, "
        f"kql_filters='{kql_filters}', limit={limit}, include_body={include_body}"
    )

    try:
        # Default to inbox-first entity types for assistant workflows
        if entity_types is None:
            entity_types = [
                "message",
                "event",
                "chatMessage",
            ]

        # Validate entity types
        valid_entity_types = {
            "message",
            "event",
            "driveItem",
            "site",
            "drive",
            "chatMessage",
            "person",
        }

        filtered_entity_types = [et for et in entity_types if et in valid_entity_types]

        if not filtered_entity_types:
            logger.warning(
                f"unified_search: No valid entity types provided from {entity_types}"
            )
            return {
                "summary": {
                    "total_results": 0,
                    "query": query,
                    "kql_filters": kql_filters,
                    "entity_types_requested": entity_types,
                    "error": "No valid entity types provided",
                },
                "results": [],
            }

        # Build the search query
        search_query = query.strip()
        if kql_filters:
            if search_query:
                search_query = f"({search_query}) AND ({kql_filters})"
            else:
                search_query = kql_filters

        # Ensure we have a valid search query - Microsoft Graph requires non-empty query
        if not search_query or search_query.isspace():
            search_query = "*"  # Use wildcard for "all content" search

        # Prepare the search request payload
        request_payload = {
            "requests": [
                {
                    "entityTypes": filtered_entity_types,
                    "query": {"queryString": search_query},
                    "size": min(limit, 25),  # Microsoft Graph max per request
                    "from": 0,
                }
            ]
        }

        # Add fields for better results
        # if include_body:
        #     request_payload["requests"][0]["fields"] = [
        #         "id",
        #         "subject",
        #         "title",
        #         "name",
        #         "body",
        #         "content",
        #         "summary",
        #         "from",
        #         "to",
        #         "sender",
        #         "author",
        #         "createdDateTime",
        #         "lastModifiedDateTime",
        #         "receivedDateTime",
        #         "sentDateTime",
        #         "size",
        #         "webUrl",
        #         "webLink",
        #     ]

        all_results = []
        entity_type_counts = {}
        total_results = 0

        # Split incompatible entity types into separate request groups
        # MS Graph Search API: event/person must be alone, message/chat can't mix with file types
        message_chat_types = {"message", "chatMessage"}
        file_types = {"driveItem", "list", "listItem", "site", "drive"}
        solo_types = {"event", "person"}

        request_groups: list[list[str]] = []
        remaining = list(filtered_entity_types)

        for st in solo_types:
            if st in remaining:
                remaining.remove(st)
                request_groups.append([st])

        msg_group = [et for et in remaining if et in message_chat_types]
        file_group = [et for et in remaining if et in file_types]
        if msg_group:
            request_groups.append(msg_group)
        if file_group:
            request_groups.append(file_group)

        if not request_groups:
            request_groups = [filtered_entity_types]

        logger.info(
            f"unified_search: Split into {len(request_groups)} request groups: {request_groups}"
        )

        # Execute search for each request group
        degraded = False
        degraded_reason = None
        try:
            for group in request_groups:
                group_payload = {
                    "requests": [
                        {
                            "entityTypes": group,
                            "query": {"queryString": search_query},
                            "size": min(limit, 25),
                            "from": 0,
                        }
                    ]
                }
                result = graph.request("POST", "/search/query", json=group_payload)

                if result and "value" in result:
                    for resp in result["value"]:
                        if "hitsContainers" in resp:
                            for container in resp["hitsContainers"]:
                                total_results += container.get("total", 0)
                                if "hits" in container:
                                    for hit in container["hits"]:
                                        processed_item = _process_search_hit(
                                            hit,
                                            include_body,
                                            body_max_length,
                                        )
                                        if processed_item:
                                            all_results.append(processed_item)
                                            kind = processed_item.get("kind", "unknown")
                                            entity_type_counts[kind] = (
                                                entity_type_counts.get(kind, 0) + 1
                                            )

        except Exception as search_error:
            status = getattr(
                getattr(search_error, "response", None), "status_code", None
            )
            if status in (403, 404):
                # Fall back to cache search
                logger.warning(
                    f"unified_search: Graph Search returned {status}, falling back to cache"
                )
                cache = get_global_cache()
                all_results = cache.search(query or "*", kinds=filtered_entity_types)
                total_results = len(all_results)
                degraded = True
                degraded_reason = f"Graph Search returned HTTP {status}"
                entity_type_counts = {}
                for item in all_results:
                    k = item.get("kind", "unknown")
                    entity_type_counts[k] = entity_type_counts.get(k, 0) + 1
            else:
                error_details = _analyze_search_error(search_error, request_payload)
                logger.error(
                    f"unified_search API error: {str(search_error)}\nError analysis: {error_details}",
                    exc_info=True,
                )
                raise RuntimeError(
                    f"Microsoft Graph Search API failed: {error_details}"
                ) from search_error

        # Populate cache with successful results for future fallback
        if not degraded and all_results:
            cache = get_global_cache()
            by_kind: dict[str, list] = {}
            for item in all_results:
                by_kind.setdefault(item.get("kind", "unknown"), []).append(item)
            for kind, items in by_kind.items():
                cache.store(kind, items)

        # Build response
        summary: dict[str, Any] = {
            "total_results": len(all_results),
            "total_available": total_results,
            "query": query,
            "kql_filters": kql_filters,
            "entity_types_searched": filtered_entity_types,
            "entity_type_counts": entity_type_counts,
            "limit_applied": limit,
            "include_body": include_body,
            "mode": "degraded_cache_search" if degraded else "graph_search",
        }

        if degraded:
            summary["degraded_reason"] = degraded_reason
            summary["data_freshness"] = get_global_cache().freshness_info()

        response = {
            "summary": summary,
            "results": all_results[:limit],
        }

        logger.info(
            f"unified_search successful: found {len(all_results)} results "
            f"across {len(entity_type_counts)} entity types with query '{search_query}'"
            + (f" (degraded: {degraded_reason})" if degraded else "")
        )
        return response

    except Exception as e:
        logger.error(f"unified_search failed: {str(e)}", exc_info=True)
        error_response = {
            "summary": {
                "total_results": 0,
                "query": query,
                "kql_filters": kql_filters,
                "entity_types_requested": entity_types,
                "error": str(e),
            },
            "results": [],
        }
        return error_response


def _detect_entity_kind(odata_type: str) -> str:
    odata_type = odata_type.lower()
    if "message" in odata_type:
        return "message"
    if "event" in odata_type:
        return "event"
    if "driveitem" in odata_type:
        return "driveItem"
    if "chatmessage" in odata_type:
        return "chatMessage"
    if "person" in odata_type:
        return "person"
    if "site" in odata_type:
        return "site"
    if "listitem" in odata_type:
        return "listItem"
    if "list" in odata_type:
        return "list"
    if "drive" in odata_type:
        return "drive"
    return "unknown"


def _process_search_hit(
    hit: dict[str, Any],
    include_body: bool,
    body_max_length: int,
) -> dict[str, Any] | None:
    """Process a single search hit into a normalized contract.

    Returns a dict with: id, kind, title, snippet, score, and kind-specific extras.
    """
    try:
        resource = hit.get("resource", {})
        if not resource:
            return None

        kind = _detect_entity_kind(resource.get("@odata.type", ""))
        title = (
            resource.get("subject")
            or resource.get("name")
            or resource.get("displayName")
            or ""
        )

        result: dict[str, Any] = {
            "id": resource.get("id") or hit.get("hitId", ""),
            "kind": kind,
            "title": title,
            "snippet": hit.get("summary", ""),
            "score": hit.get("rank", 0),
        }

        # Kind-specific extras
        if kind == "message":
            if resource.get("from"):
                result["from"] = flatten_email_address(resource["from"])
            if resource.get("receivedDateTime"):
                result["received"] = resource["receivedDateTime"]
            conv_id = resource.get("conversationId")
            if conv_id:
                result["conversation_url"] = (
                    f"https://outlook.office.com/mail/deeplink/readconv/{quote(conv_id, safe='')}"
                )
        elif kind == "event":
            for key in ("start", "end"):
                if key in resource:
                    result[key] = resource[key]
            if resource.get("location"):
                from .response_shaping import compact_location

                result["location"] = compact_location(resource["location"])
            if resource.get("organizer"):
                result["organizer"] = flatten_email_address(resource["organizer"])
        elif kind == "driveItem":
            result["size"] = resource.get("size", 0)
            result["modified"] = resource.get("lastModifiedDateTime")
            result["web_url"] = resource.get("webUrl")

        # Body handling
        if include_body and "body" in resource:
            body = resource["body"]
            if isinstance(body, dict):
                body_content = body.get("content", "")
                content_type = body.get("contentType", "")
                if content_type.lower() == "html" and body_content:
                    body_content = convert_to_markdown(body_content)
                if body_content and len(body_content) > body_max_length:
                    body_content = body_content[:body_max_length] + "...[truncated]"
                result["body"] = body_content

        return result

    except Exception as e:
        logger.warning(f"Failed to process search hit: {str(e)}")
        return None


@mcp.tool
def search_files(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for files and folders in OneDrive using text queries.

    Find files by name, content, or metadata across your entire OneDrive. More powerful than
    browsing folders - searches filenames, document content, and file properties.

    Args:
        query: Search terms to find files (e.g., "budget report", "vacation photos", "presentation")
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching file/folder objects containing:
        - Basic info: id, name, type (file/folder), size (bytes), modified (timestamp)
        - Download info: download_url for direct file access (for files only)
        - Results ranked by relevance to search query

    Examples:
        - search_files("presentation") - Find files with "presentation" in name or content
        - search_files("budget 2024") - Find budget-related files from 2024
        - search_files("photos vacation") - Find vacation photos
        - search_files(".pdf report") - Find PDF files containing "report"
    """
    logger.info(f"search_files called: query='{query}', limit={limit}")

    try:
        items = list(graph.search_query(query, ["driveItem"], limit))

        result = [
            {
                "id": item["id"],
                "name": item["name"],
                "type": "folder" if "folder" in item else "file",
                "size": item.get("size", 0),
                "modified": item.get("lastModifiedDateTime"),
                "download_url": item.get("@microsoft.graph.downloadUrl"),
            }
            for item in items
        ]

        logger.info(
            f"search_files successful: found {len(result)} files matching '{query}'"
        )
        return result
    except Exception as e:
        logger.error(
            f"search_files failed for query='{query}': {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def search_emails(
    query: str,
    limit: int = 50,
    folder: str | None = None,
) -> list[dict[str, Any]]:
    """Search for emails using text queries across subjects, content, and metadata.

    Find emails by searching subject lines, body content, sender/recipient names, and other metadata.
    Can search across all emails or within a specific folder. More powerful than browsing folders.

    Args:
        query: Search terms (e.g., "meeting notes", "project update", sender name, subject keywords)
        limit: Maximum number of results to return (1-100, defaults to 50)
        folder: Optional folder alias, ID, display name, or slash-delimited path to search within.
            If None, searches across all emails.

    Returns:
        List of matching email objects containing:
        - Basic info: id, subject, from, toRecipients, receivedDateTime, isRead
        - Indicators: hasAttachments, conversationId
        - Body content (if available)
        - Results ranked by relevance to search query
        - a deep link to the conversation as `conversation_url` that can be shown to the user to open the email

    Examples:
        - search_emails("project alpha") - Find emails about "project alpha" anywhere
        - search_emails("meeting", folder="inbox") - Find meeting emails only in inbox
        - search_emails("deals", folder="Cresa Deals of the Week") - Search inside a custom folder
        - search_emails("emails received today") - Finds emails from today
        - search_emails("john.doe@company.com") - Find emails from/to specific person
        - search_emails("budget approval") - Find emails about budget approvals
    """
    logger.info(
        f"search_emails called: query='{query}', limit={limit}, folder={folder}"
    )

    try:
        if folder:
            folder_path = _resolve_mail_folder(folder)
            endpoint = f"/me/mailFolders/{folder_path}/messages"

            params = {
                "$search": f'"{query}"',
                "$top": min(limit, 100),
                "$select": "id,subject,from,toRecipients,receivedDateTime,hasAttachments,bodyPreview,conversationId,isRead",
            }

            raw = list(graph.request_paginated(endpoint, params=params, limit=limit))
        else:
            raw = list(graph.search_query(query, ["message"], limit))

        result = [shape_email_summary(e) for e in raw]

        logger.info(
            f"search_emails successful: found {len(result)} emails matching '{query}'"
        )
        return result
    except Exception as e:
        logger.error(
            f"search_emails failed for query='{query}', folder={folder}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def search_events(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for calendar events using text queries across titles, content, and metadata.

    Find calendar events by searching event titles, descriptions, locations, and attendee information.
    Useful for finding specific meetings, events with certain keywords, or events in particular locations.

    Args:
        query: Search terms (e.g., "team meeting", "conference room", "project review", attendee names)
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching calendar event objects containing:
        - Basic info: id, subject, start/end times, location, organizer
        - Details: body/description, attendees, isAllDay status
        - Meeting info: onlineMeeting links if applicable
        - Recurrence: seriesMasterId for recurring events
        - Results ranked by relevance to search query

    Examples:
        - search_events("standup") - Find daily standup meetings
        - search_events("conference room A") - Find events in specific room
        - search_events("john smith") - Find events with John Smith as organizer/attendee
        - search_events("quarterly review") - Find quarterly review meetings
    """
    logger.info(f"search_events called: query='{query}', limit={limit}")

    try:
        raw_events = list(graph.search_query(query, ["event"], limit))
        events = [shape_event_summary(e) for e in raw_events]

        logger.info(
            f"search_events successful: found {len(events)} events matching '{query}'"
        )
        return events
    except Exception as e:
        logger.error(
            f"search_events failed for query='{query}': {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def search_contacts(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for contacts using text queries across names, email addresses, and other fields.

    Find contacts by searching names, email addresses, phone numbers, company names, and other
    contact information. More efficient than browsing all contacts when looking for specific people.

    Args:
        query: Search terms (e.g., person name, email address, company name, phone number)
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching contact objects containing:
        - Names: givenName, surname, displayName, nickname
        - Communications: emailAddresses, businessPhones, homePhones, mobilePhone
        - Professional: jobTitle, companyName, department
        - Addresses: business and home address information
        - Results ranked by relevance to search query

    Examples:
        - search_contacts("john") - Find contacts with "john" in their name
        - search_contacts("microsoft") - Find contacts who work at Microsoft
        - search_contacts("john.doe@company.com") - Find contact with specific email
        - search_contacts("555-0123") - Find contact with specific phone number
    """
    logger.info(f"search_contacts called: query='{query}', limit={limit}")

    try:
        params = {
            "$search": f'"{query}"',
            "$top": min(limit, 100),
        }

        raw_contacts = list(
            graph.request_paginated("/me/contacts", params=params, limit=limit)
        )
        contacts = [shape_contact_summary(c) for c in raw_contacts]

        logger.info(
            f"search_contacts successful: found {len(contacts)} contacts matching '{query}'"
        )
        return contacts
    except Exception as e:
        logger.error(
            f"search_contacts failed for query='{query}': {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def list_chat_messages(
    chat_id: str | None = None,
    limit: int = 10,
    recent_container_limit: int = 10,
    body_max_length: int = 2000,
    include_body: bool = True,
    start_date: str | None = None,
    end_date: str | None = None,
    response_profile: str = "auto",
) -> list[dict[str, Any]]:
    """List recent chat messages from Teams chats.

    Retrieves messages from Microsoft Teams chats, ordered by most recent first. This includes
    both one-on-one chats and group chat conversations. Use this to get an overview of recent
    chat activity across all conversations.

    Args:
        chat_id: Optional ID of a specific chat to retrieve messages from. If provided, only
                 messages from this chat are returned (targeted mode).
        limit: Maximum number of messages to retrieve (1-100, defaults to 10)
        recent_container_limit: Maximum number of recent chats to scan when chat_id is not
                                provided (default 10). Prevents scanning every chat.
        body_max_length: Maximum characters for message body content (default 2000, will truncate if longer)
        include_body: Whether to include message body content (affects response size)
        start_date: Optional start date in ISO format (UTC timezone, e.g., "2024-09-01T00:00:00Z") to filter messages from this date onwards
        end_date: Optional end date in ISO format (UTC timezone, e.g., "2024-09-30T23:59:59Z") to filter messages up to this date
        response_profile: Response shaping profile ("auto", "legacy", or "assistant"). "auto" defers to MICROSOFT_MCP_RESPONSE_PROFILE env var.

    Returns:
        List of message objects containing:
        - Basic info: id, chatId, messageType, createdDateTime, from (sender info)
        - Content: body with text/html content and optionally attachments info
        - Chat context: chat title, chat type (oneOnOne/group), participant count
        - Web URL for opening the message in Teams client
        - The most recent message will be first in the results

    Examples:
        - list_chat_messages() - Get 10 most recent chat messages
        - list_chat_messages(chat_id="19:abc...") - Get messages from a specific chat
        - list_chat_messages(limit=50) - Get more recent messages
        - list_chat_messages(include_body=False) - Get messages without body content for faster response
        - list_chat_messages(start_date="2024-09-01T00:00:00Z") - Get messages from September 1st onwards
    """
    profile = get_response_profile(response_profile)
    if profile == "assistant":
        include_body = False

    logger.info(
        f"list_chat_messages called: chat_id={chat_id}, limit={limit}, include_body={include_body}, start_date={start_date}, end_date={end_date}, profile={profile}"
    )

    try:
        if chat_id:
            # Targeted mode: single chat
            chats = [{"id": chat_id, "topic": "", "chatType": "", "webUrl": ""}]
        else:
            # Bounded scan: fetch only top-N recent chats
            chats = list(
                graph.request_paginated(
                    "/me/chats",
                    params={"$top": min(recent_container_limit, 50)},
                    limit=recent_container_limit,
                )
            )

        all_messages = []

        for chat in chats:
            chat_id = chat["id"]

            # Build message query parameters
            if include_body:
                select_fields = "id,messageType,createdDateTime,from,body,attachments"
            else:
                select_fields = "id,messageType,createdDateTime,from"

            params = {
                "$top": min(limit, 100),
                "$select": select_fields,
                "$orderby": "createdDateTime desc",
            }

            # Add date filtering if provided
            filter_conditions = []
            if start_date:
                filter_conditions.append(f"createdDateTime ge {start_date}")
            if end_date:
                filter_conditions.append(f"createdDateTime le {end_date}")

            if filter_conditions:
                params["$filter"] = " and ".join(filter_conditions)

            try:
                messages = list(
                    graph.request_paginated(
                        f"/me/chats/{chat_id}/messages",
                        params=params,
                        limit=min(
                            limit, 20
                        ),  # Limit per chat to avoid too many results
                    )
                )

                for message in messages:
                    # Add chat context to each message
                    message["chatId"] = chat_id
                    message["chatTopic"] = chat.get("topic", "")
                    message["chatType"] = chat.get("chatType", "")
                    message["webUrl"] = chat.get("webUrl", "")

                    # Process message body
                    if (
                        include_body
                        and "body" in message
                        and "content" in message["body"]
                    ):
                        content = message["body"]["content"]
                        if message["body"].get("contentType") == "html":
                            content = convert_to_markdown(content)
                            message["body"]["contentType"] = "text/markdown"
                            message["body"]["content"] = content

                        if len(content) > body_max_length:
                            message["body"]["content"] = (
                                content[:body_max_length]
                                + f"\n\n[Content truncated - {len(content)} total characters]"
                            )
                            message["body"]["truncated"] = True
                            message["body"]["total_length"] = len(content)

                all_messages.extend(messages)

            except Exception as chat_error:
                logger.warning(
                    f"Failed to get messages from chat {chat_id}: {str(chat_error)}"
                )
                continue

        # Sort all messages by creation time (most recent first) and limit results
        all_messages.sort(key=lambda x: x.get("createdDateTime", ""), reverse=True)
        result = all_messages[:limit]

        logger.info(
            f"list_chat_messages successful: retrieved {len(result)} messages from {len(chats)} chats"
        )
        return result
    except Exception as e:
        logger.error(f"list_chat_messages failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def list_channel_messages(
    team_id: str | None = None,
    channel_id: str | None = None,
    limit: int = 10,
    recent_team_limit: int = 5,
    body_max_length: int = 2000,
    include_body: bool = True,
    start_date: str | None = None,
    end_date: str | None = None,
) -> list[dict[str, Any]]:
    """List recent messages from Microsoft Teams channels.

    Retrieves messages from team channels, ordered by most recent first. If no team/channel is specified,
    gets messages from only the most recent teams to avoid excessive fan-out.

    Args:
        team_id: Optional ID of specific team to get messages from
        channel_id: Optional ID of specific channel (requires team_id)
        limit: Maximum number of messages to retrieve (1-100, defaults to 10)
        recent_team_limit: Maximum number of teams to scan when team_id is not provided (default 5)
        body_max_length: Maximum characters for message body content (default 2000, will truncate if longer)
        include_body: Whether to include message body content (affects response size)
        start_date: Optional start date in ISO format (UTC timezone, e.g., "2024-09-01T00:00:00Z") to filter messages from this date onwards
        end_date: Optional end date in ISO format (UTC timezone, e.g., "2024-09-30T23:59:59Z") to filter messages up to this date

    Returns:
        List of message objects containing:
        - Basic info: id, messageType, createdDateTime, from (sender info)
        - Content: body with text/html content and optionally attachments info
        - Channel context: team name, channel name, channel description
        - Web URL for opening the message in Teams client
        - The most recent message will be first in the results

    Examples:
        - list_channel_messages() - Get 10 most recent messages from all channels
        - list_channel_messages(team_id="abc123") - Get messages from specific team's channels
        - list_channel_messages(team_id="abc123", channel_id="def456") - Get messages from specific channel
        - list_channel_messages(limit=50, include_body=False) - Get more messages without body content
    """
    logger.info(
        f"list_channel_messages called: team_id={team_id}, channel_id={channel_id}, limit={limit}, include_body={include_body}"
    )

    try:
        all_messages = []

        if team_id and channel_id:
            # Get messages from specific channel
            teams_channels = [(team_id, channel_id)]
        elif team_id:
            # Get all channels from specific team
            channels = list(graph.request_paginated(f"/teams/{team_id}/channels"))
            teams_channels = [(team_id, ch["id"]) for ch in channels]
        else:
            # Bounded scan: only top-N recent teams
            teams = list(
                graph.request_paginated(
                    "/me/joinedTeams",
                    params={"$top": min(recent_team_limit, 25)},
                    limit=recent_team_limit,
                )
            )
            teams_channels = []
            for team in teams:
                try:
                    channels = list(
                        graph.request_paginated(f"/teams/{team['id']}/channels")
                    )
                    teams_channels.extend([(team["id"], ch["id"]) for ch in channels])
                except Exception as team_error:
                    logger.warning(
                        f"Failed to get channels from team {team['id']}: {str(team_error)}"
                    )
                    continue

        for t_id, c_id in teams_channels:
            # Build message query parameters
            if include_body:
                select_fields = "id,messageType,createdDateTime,from,body,attachments"
            else:
                select_fields = "id,messageType,createdDateTime,from"

            params = {
                "$top": min(limit, 50),
                "$select": select_fields,
                "$orderby": "createdDateTime desc",
            }

            # Add date filtering if provided
            filter_conditions = []
            if start_date:
                filter_conditions.append(f"createdDateTime ge {start_date}")
            if end_date:
                filter_conditions.append(f"createdDateTime le {end_date}")

            if filter_conditions:
                params["$filter"] = " and ".join(filter_conditions)

            try:
                # Get team and channel info for context
                team_info = graph.request("GET", f"/teams/{t_id}")
                channel_info = graph.request("GET", f"/teams/{t_id}/channels/{c_id}")

                messages = list(
                    graph.request_paginated(
                        f"/teams/{t_id}/channels/{c_id}/messages",
                        params=params,
                        limit=min(limit, 20),  # Limit per channel
                    )
                )

                for message in messages:
                    # Add channel context to each message
                    message["teamId"] = t_id
                    message["channelId"] = c_id
                    message["teamName"] = (
                        team_info.get("displayName", "") if team_info else ""
                    )
                    message["channelName"] = (
                        channel_info.get("displayName", "") if channel_info else ""
                    )
                    message["webUrl"] = (
                        channel_info.get("webUrl", "") if channel_info else ""
                    )

                    # Process message body
                    if (
                        include_body
                        and "body" in message
                        and "content" in message["body"]
                    ):
                        content = message["body"]["content"]
                        if message["body"].get("contentType") == "html":
                            content = convert_to_markdown(content)
                            message["body"]["contentType"] = "text/markdown"
                            message["body"]["content"] = content

                        if len(content) > body_max_length:
                            message["body"]["content"] = (
                                content[:body_max_length]
                                + f"\n\n[Content truncated - {len(content)} total characters]"
                            )
                            message["body"]["truncated"] = True
                            message["body"]["total_length"] = len(content)

                all_messages.extend(messages)

            except Exception as channel_error:
                logger.warning(
                    f"Failed to get messages from team {t_id}, channel {c_id}: {str(channel_error)}"
                )
                continue

        # Sort all messages by creation time (most recent first) and limit results
        all_messages.sort(key=lambda x: x.get("createdDateTime", ""), reverse=True)
        result = all_messages[:limit]

        logger.info(
            f"list_channel_messages successful: retrieved {len(result)} messages from {len(teams_channels)} channels"
        )
        return result
    except Exception as e:
        logger.error(f"list_channel_messages failed: {str(e)}", exc_info=True)
        raise


@mcp.tool
def get_chat_message(chat_id: str, message_id: str) -> dict[str, Any]:
    """Get detailed information about a specific chat message by its ID.

    Retrieves complete message details including content, attachments, reactions, and replies.
    Use this when you need full message details after finding messages with list_chat_messages
    or search_chat_messages.

    Args:
        chat_id: Unique identifier of the chat containing the message
        message_id: Unique identifier of the specific message

    Returns:
        Complete message object containing:
        - Basic info: id, messageType, createdDateTime, lastModifiedDateTime, from (sender)
        - Content: body with full text/html content, contentType
        - Attachments: list with attachment details if any
        - Reactions: emoji reactions and who reacted
        - Chat context: chat info, participants
        - Web URL for opening in Teams client

    Examples:
        - get_chat_message(chat_id, message_id) - Get full message details
        - Use after finding interesting messages with list_chat_messages()
    """
    logger.info(f"get_chat_message called: chat_id={chat_id}, message_id={message_id}")

    try:
        # Get the message details
        message = graph.request("GET", f"/me/chats/{chat_id}/messages/{message_id}")
        if not message:
            logger.error(
                f"get_chat_message failed: Message {message_id} not found in chat {chat_id}"
            )
            raise ValueError(f"Message {message_id} not found in chat {chat_id}")

        # Get chat context
        try:
            chat_info = graph.request("GET", f"/me/chats/{chat_id}")
            if chat_info:
                message["chatContext"] = {
                    "topic": chat_info.get("topic", ""),
                    "chatType": chat_info.get("chatType", ""),
                    "webUrl": chat_info.get("webUrl", ""),
                }
        except Exception as context_error:
            logger.warning(f"Could not get chat context: {str(context_error)}")

        # Convert HTML body to markdown if needed
        if "body" in message and "content" in message["body"]:
            if message["body"].get("contentType") == "html":
                message["body"]["content"] = convert_to_markdown(
                    message["body"]["content"]
                )
                message["body"]["contentType"] = "text/markdown"

        # Get message replies if any
        try:
            replies = list(
                graph.request_paginated(
                    f"/me/chats/{chat_id}/messages/{message_id}/replies"
                )
            )
            if replies:
                message["replies"] = replies
                message["replyCount"] = len(replies)
        except Exception as replies_error:
            logger.warning(f"Could not get message replies: {str(replies_error)}")

        logger.info(f"get_chat_message successful: retrieved message {message_id}")
        return message
    except Exception as e:
        logger.error(
            f"get_chat_message failed for chat_id={chat_id}, message_id={message_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def get_channel_message(
    team_id: str, channel_id: str, message_id: str
) -> dict[str, Any]:
    """Get detailed information about a specific channel message by its ID.

    Retrieves complete message details including content, attachments, reactions, and replies.
    Use this when you need full message details after finding messages with list_channel_messages
    or search_channel_messages.

    Args:
        team_id: Unique identifier of the team containing the channel
        channel_id: Unique identifier of the channel containing the message
        message_id: Unique identifier of the specific message

    Returns:
        Complete message object containing:
        - Basic info: id, messageType, createdDateTime, lastModifiedDateTime, from (sender)
        - Content: body with full text/html content, contentType
        - Attachments: list with attachment details if any
        - Reactions: emoji reactions and who reacted
        - Channel context: team name, channel name, web URL
        - Web URL for opening in Teams client

    Examples:
        - get_channel_message(team_id, channel_id, message_id) - Get full message details
        - Use after finding interesting messages with list_channel_messages()
    """
    logger.info(
        f"get_channel_message called: team_id={team_id}, channel_id={channel_id}, message_id={message_id}"
    )

    try:
        # Get the message details
        message = graph.request(
            "GET", f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}"
        )
        if not message:
            logger.error(f"get_channel_message failed: Message {message_id} not found")
            raise ValueError(f"Message {message_id} not found in channel {channel_id}")

        # Get team and channel context
        try:
            team_info = graph.request("GET", f"/teams/{team_id}")
            channel_info = graph.request(
                "GET", f"/teams/{team_id}/channels/{channel_id}"
            )

            if team_info and channel_info:
                message["channelContext"] = {
                    "teamName": team_info.get("displayName", ""),
                    "channelName": channel_info.get("displayName", ""),
                    "channelDescription": channel_info.get("description", ""),
                    "webUrl": channel_info.get("webUrl", ""),
                }
        except Exception as context_error:
            logger.warning(f"Could not get channel context: {str(context_error)}")

        # Convert HTML body to markdown if needed
        if "body" in message and "content" in message["body"]:
            if message["body"].get("contentType") == "html":
                message["body"]["content"] = convert_to_markdown(
                    message["body"]["content"]
                )
                message["body"]["contentType"] = "text/markdown"

        # Get message replies if any
        try:
            replies = list(
                graph.request_paginated(
                    f"/teams/{team_id}/channels/{channel_id}/messages/{message_id}/replies"
                )
            )
            if replies:
                message["replies"] = replies
                message["replyCount"] = len(replies)
        except Exception as replies_error:
            logger.warning(f"Could not get message replies: {str(replies_error)}")

        logger.info(f"get_channel_message successful: retrieved message {message_id}")
        return message
    except Exception as e:
        logger.error(
            f"get_channel_message failed for team_id={team_id}, channel_id={channel_id}, message_id={message_id}: {str(e)}",
            exc_info=True,
        )
        raise


@mcp.tool
def search_chat_messages(
    query: str,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for chat messages using text queries across message content and metadata.

    Find chat messages by searching message content, sender names, and other metadata across
    all accessible 1:1 and group chats. More powerful than browsing recent messages when looking
    for specific content or conversations.

    Args:
        query: Search terms (e.g., "project update", "meeting notes", sender name, keywords)
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching message objects containing:
        - Basic info: id, chatId, messageType, createdDateTime, from (sender)
        - Content: body with text content matching the search query
        - Chat context: chat topic, chat type, participant info
        - Web URL for opening in Teams client
        - Results ranked by relevance to search query

    Examples:
        - search_chat_messages("project alpha") - Find chat messages about "project alpha"
        - search_chat_messages("john smith") - Find messages from/mentioning John Smith
        - search_chat_messages("meeting tomorrow") - Find messages about upcoming meetings
        - search_chat_messages("budget approval") - Find chat messages about budget approvals
    """
    logger.info(f"search_chat_messages called: query='{query}', limit={limit}")

    try:
        # Use Microsoft Graph search API to search across chat messages
        result = list(graph.search_query(query, ["chatMessage"], limit))

        # Process results to add chat context
        for message in result:
            # Extract chat ID from message if available
            if "chatId" in message:
                chat_id = message["chatId"]
                try:
                    chat_info = graph.request("GET", f"/me/chats/{chat_id}")
                    if chat_info:
                        message["chatTopic"] = chat_info.get("topic", "")
                        message["chatType"] = chat_info.get("chatType", "")
                        message["webUrl"] = chat_info.get("webUrl", "")
                except Exception as context_error:
                    logger.warning(
                        f"Could not get chat context for chat {chat_id}: {str(context_error)}"
                    )

            # Convert HTML to markdown if needed
            if "body" in message and "content" in message["body"]:
                if message["body"].get("contentType") == "html":
                    message["body"]["content"] = convert_to_markdown(
                        message["body"]["content"]
                    )
                    message["body"]["contentType"] = "text/markdown"

        logger.info(
            f"search_chat_messages successful: found {len(result)} messages matching '{query}'"
        )
        return result
    except Exception as e:
        logger.error(
            f"search_chat_messages failed for query='{query}': {str(e)}", exc_info=True
        )
        raise


@mcp.tool
def search_channel_messages(
    query: str,
    team_id: str | None = None,
    limit: int = 50,
) -> list[dict[str, Any]]:
    """Search for channel messages using text queries across message content and metadata.

    Find channel messages by searching message content, sender names, and other metadata across
    all accessible team channels or within a specific team. More powerful than browsing recent
    messages when looking for specific content or discussions.

    Args:
        query: Search terms (e.g., "project update", "announcement", sender name, keywords)
        team_id: Optional ID of specific team to search within (searches all teams if not provided)
        limit: Maximum number of results to return (1-100, defaults to 50)

    Returns:
        List of matching message objects containing:
        - Basic info: id, teamId, channelId, messageType, createdDateTime, from (sender)
        - Content: body with text content matching the search query
        - Channel context: team name, channel name, web URL
        - Results ranked by relevance to search query

    Examples:
        - search_channel_messages("project alpha") - Find channel messages about "project alpha"
        - search_channel_messages("announcement", team_id="abc123") - Find announcements in specific team
        - search_channel_messages("quarterly review") - Find messages about quarterly reviews
        - search_channel_messages("john smith") - Find messages from/mentioning John Smith
    """
    logger.info(
        f"search_channel_messages called: query='{query}', team_id={team_id}, limit={limit}"
    )

    try:
        # Use Microsoft Graph search API to search across channel messages
        # If team_id is provided, we could filter results, but Graph search doesn't directly support this
        # So we'll search all and filter afterward if needed
        result = list(graph.search_query(query, ["chatMessage"], limit))

        # Filter for channel messages only (not chat messages) and optionally by team
        channel_messages = []
        for message in result:
            # Channel messages have teamId and channelId, chat messages have chatId
            if "teamId" in message and "channelId" in message:
                if team_id is None or message.get("teamId") == team_id:
                    try:
                        # Get team and channel context
                        t_id = message["teamId"]
                        c_id = message["channelId"]

                        team_info = graph.request("GET", f"/teams/{t_id}")
                        channel_info = graph.request(
                            "GET", f"/teams/{t_id}/channels/{c_id}"
                        )

                        if team_info and channel_info:
                            message["teamName"] = team_info.get("displayName", "")
                            message["channelName"] = channel_info.get("displayName", "")
                            message["webUrl"] = channel_info.get("webUrl", "")
                    except Exception as context_error:
                        logger.warning(
                            f"Could not get channel context: {str(context_error)}"
                        )

                    # Convert HTML to markdown if needed
                    if "body" in message and "content" in message["body"]:
                        if message["body"].get("contentType") == "html":
                            message["body"]["content"] = convert_to_markdown(
                                message["body"]["content"]
                            )
                            message["body"]["contentType"] = "text/markdown"

                    channel_messages.append(message)

        # Limit results
        result = channel_messages[:limit]

        logger.info(
            f"search_channel_messages successful: found {len(result)} channel messages matching '{query}'"
            + (f" in team {team_id}" if team_id else "")
        )
        return result
    except Exception as e:
        logger.error(
            f"search_channel_messages failed for query='{query}', team_id={team_id}: {str(e)}",
            exc_info=True,
        )
        raise


TEAMS_TOOL_NAMES = (
    "list_chat_messages",
    "list_channel_messages",
    "get_chat_message",
    "get_channel_message",
    "search_chat_messages",
    "search_channel_messages",
)


def _list_internal_business_tools() -> list[Any]:
    tools: list[Any] = []
    for tool_name, tool in mcp._tool_manager._tools.items():
        if tool_name in CODE_MODE_TOOL_NAMES:
            continue
        if auth_method == "msal" and tool_name in TEAMS_TOOL_NAMES:
            continue
        tools.append(tool)
    return tools


def _configure_teams_tools_for_auth_method() -> None:
    if auth_method != "msal":
        return

    for tool_name in TEAMS_TOOL_NAMES:
        tool = mcp._tool_manager._tools.get(tool_name)
        if tool is None:
            logger.warning(
                "Expected Teams tool '%s' was not registered before MSAL gating",
                tool_name,
            )
            continue
        tool.disable()

    logger.info("Disabled Teams tools for MSAL authentication method")


_configure_teams_tools_for_auth_method()


def _configure_public_tool_mode() -> None:
    if tool_mode == "hybrid":
        logger.info("Using hybrid public tool mode")
        return

    for tool_name, tool in mcp._tool_manager._tools.items():
        if tool_name in CODE_MODE_TOOL_NAMES:
            continue
        tool.disable()

    logger.info("Enabled code-mode-only public tool mode")


# ============================================================================
# Assistant-Native Inbox Tools
# ============================================================================


def _emails_to_inbox_items(raw_emails: list[dict[str, Any]]) -> list[InboxItem]:
    items = []
    for e in raw_emails:
        from_addr = ""
        if "from" in e:
            from_addr = flatten_email_address(e["from"])
        items.append(
            InboxItem(
                id=e["id"],
                kind="email",
                source_tool="list_emails",
                title=e.get("subject", ""),
                snippet=e.get("bodyPreview", "")[:200],
                participants=[from_addr] if from_addr else [],
                when=e.get("receivedDateTime"),
                unread=not e.get("isRead", True),
                web_url=f"https://outlook.office.com/mail/deeplink/readconv/{quote(e.get('conversationId', ''), safe='')}"
                if e.get("conversationId")
                else "",
            )
        )
    return items


def _invite_messages_to_inbox_items(
    raw_invite_messages: list[dict[str, Any]],
) -> list[InboxItem]:
    items = []
    for message in raw_invite_messages:
        from_addr = ""
        if "from" in message:
            from_addr = flatten_email_address(message["from"])

        meeting_message_type = message.get("meetingMessageType", "")
        action_hints = ["review"]
        if meeting_message_type == "meetingRequest":
            action_hints = ["rsvp", "delete"]
        elif meeting_message_type == "meetingCancelled":
            action_hints = ["delete"]

        items.append(
            InboxItem(
                id=message["id"],
                kind="invite_message",
                source_tool="list_invite_messages",
                title=message.get("subject", ""),
                snippet=message.get("bodyPreview", "")[:200],
                participants=[from_addr] if from_addr else [],
                when=message.get("startDateTime", {}).get("dateTime")
                or message.get("receivedDateTime"),
                unread=not message.get("isRead", True),
                state=meeting_message_type,
                action_hints=action_hints,
                web_url=message.get("webLink", ""),
            )
        )
    return items


def _events_to_inbox_items(raw_events: list[dict[str, Any]]) -> list[InboxItem]:
    items = []
    for ev in raw_events:
        organizer = ""
        if "organizer" in ev:
            organizer = flatten_email_address(ev["organizer"])
        start_dt = ev.get("start", {}).get("dateTime", "")
        items.append(
            InboxItem(
                id=ev["id"],
                kind="event",
                source_tool="list_events",
                title=ev.get("subject", ""),
                participants=[organizer] if organizer else [],
                when=start_dt,
            )
        )
    return items


@mcp.tool
def list_inbox_items(
    limit: int = 20,
    include_kinds: list[str] | None = None,
) -> dict[str, Any]:
    """Get a ranked, mixed-kind inbox summary across email and calendar.

    Returns a priority-ranked list of inbox items (emails, invite messages, and events) scored by
    urgency signals like unread status, mentions, and meeting proximity.

    Args:
        limit: Maximum items to return (default 20)
        include_kinds: Optional filter, e.g. ["email"], ["event"], ["invite_message"], or a mix.

    Returns:
        Dictionary with 'items' (ranked list) and 'meta' (counts, sources).
    """
    logger.info(
        f"list_inbox_items called: limit={limit}, include_kinds={include_kinds}"
    )
    all_items: list[InboxItem] = []
    kinds = (
        set(include_kinds) if include_kinds else {"email", "event", "invite_message"}
    )

    if kinds & {"email", "invite_message"}:
        try:
            message_fetch_limit = (
                max(limit * 5, min(limit + 20, 100))
                if "invite_message" in kinds
                else limit
            )
            raw_messages = _list_message_summaries("inbox", message_fetch_limit)

            invite_messages: list[dict[str, Any]] = []
            invite_ids: set[str] = set()
            if "invite_message" in kinds:
                invite_messages = _hydrate_invite_messages_from_summaries(
                    raw_messages,
                    limit=limit,
                )
                invite_ids = {message["id"] for message in invite_messages}
                all_items.extend(_invite_messages_to_inbox_items(invite_messages))

            if "email" in kinds:
                all_items.extend(
                    _emails_to_inbox_items(
                        [
                            message
                            for message in raw_messages
                            if message.get("id") not in invite_ids
                        ]
                    )
                )
        except Exception as e:
            logger.error("list_inbox_items message fetch failed: %s", e, exc_info=True)

    if "event" in kinds:
        try:
            now = dt.datetime.now(dt.timezone.utc)
            params = {
                "startDateTime": now.isoformat(),
                "endDateTime": (now + dt.timedelta(days=2)).isoformat(),
                "$orderby": "start/dateTime",
                "$top": min(limit, 25),
                "$select": "id,subject,start,end,location,organizer,seriesMasterId",
            }
            raw = list(
                graph.request_paginated("/me/calendarView", params=params, limit=limit)
            )
            all_items.extend(_events_to_inbox_items(raw))
        except Exception as e:
            logger.error("list_inbox_items event fetch failed: %s", e, exc_info=True)

    ranked = rank_items(all_items)[:limit]

    return {
        "items": [item.to_dict() for item in ranked],
        "meta": {
            "total_fetched": len(all_items),
            "returned": len(ranked),
            "kinds": list(kinds),
        },
    }


@mcp.tool
def get_inbox_item_detail(item_id: str, kind: str) -> dict[str, Any]:
    """Hydrate full details for a single inbox item.

    Args:
        item_id: The item ID from list_inbox_items results
        kind: The item kind ("email", "invite_message", or "event")

    Returns:
        Full item detail with body content included.
    """
    logger.info(f"get_inbox_item_detail called: item_id={item_id}, kind={kind}")

    if kind == "email":
        raw = graph.request("GET", f"/me/messages/{item_id}")
        if not raw:
            raise ValueError(f"Email {item_id} not found")
        detail = shape_email_detail(raw)
        detail["kind"] = "email"
        return detail

    if kind == "event":
        raw = graph.request("GET", f"/me/events/{item_id}")
        if not raw:
            raise ValueError(f"Event {item_id} not found")
        detail = shape_event_detail(raw)
        detail["kind"] = "event"
        return detail

    if kind == "invite_message":
        raw = _get_invite_message(item_id, include_body=True, expand_event=True)
        return _shape_invite_message(raw, include_body=True)

    raise ValueError(f"Unsupported kind: {kind}")


_configure_public_tool_mode()
