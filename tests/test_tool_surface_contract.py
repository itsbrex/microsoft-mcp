"""Contract tests enforcing tool-surface audit invariants.

Locks in the fixes from the 2026-04-23 audit so any future regression
(missing response_profile, raw httpx call, unpopulated ranker signal,
re-introduced envelope bloat, lost SCOPES dedupe, dropped exception
chaining) fails CI.

DO NOT remove an assertion in this file without first checking that the
underlying invariant is intentionally being relaxed — many of these
guards exist because the audit found a real production bug.
"""

import inspect
import re

from microsoft_mcp import auth, auth_msal, code_mode
from microsoft_mcp import tools as tools_mod


# ---------------------------------------------------------------------------
# Tool registry invariants
# ---------------------------------------------------------------------------

LIST_OR_SEARCH_TOOLS = [
    "list_emails",
    "list_events",
    "list_contacts",
    "list_chat_messages",
    "list_mail_folders",
    "list_master_categories",
    "list_invite_messages",
    "list_files",
    "unified_search",
    "search_files",
    "search_emails",
    "search_events",
    "search_contacts",
    "list_channel_messages",
    "search_chat_messages",
    "search_channel_messages",
    "list_inbox_items",
]


def test_all_list_search_tools_accept_response_profile():
    """A1: Every list/search tool exposes response_profile = 'auto'."""
    missing = []
    for name in LIST_OR_SEARCH_TOOLS:
        tool = getattr(tools_mod, name, None)
        assert tool is not None, f"{name} not exported from tools module"
        fn = getattr(tool, "fn", tool)
        sig = inspect.signature(fn)
        param = sig.parameters.get("response_profile")
        if param is None:
            missing.append(name)
            continue
        # The default must be 'auto' so MICROSOFT_MCP_RESPONSE_PROFILE env var resolves.
        assert param.default == "auto", (
            f"{name}.response_profile default is {param.default!r}, expected 'auto'"
        )
    assert not missing, f"tools missing response_profile: {missing}"


def test_no_direct_httpx_calls_in_tools_module():
    """B6/B7 spirit: tools.py must route every Graph call through graph.request.

    Raw httpx.Client / httpx.get etc. bypass the retry, pagination, and auth
    handling baked into microsoft_mcp.graph.
    """
    src = inspect.getsource(tools_mod)
    assert not re.search(r"httpx\.(Async)?Client\(", src), (
        "tools.py instantiates an httpx client directly; route through graph.request"
    )
    assert not re.search(r"httpx\.(get|post|put|delete|patch)\(", src), (
        "tools.py has a raw httpx HTTP call; route through graph.request"
    )


def test_inbox_ranker_signals_are_populated():
    """B3: Each ranker signal must be populated by its adapter helper.

    Constructs raw inputs that exercise each signal and verifies the
    InboxItem returned by the relevant _*_to_inbox_items helper has the
    field set. This catches regressions where a populator is removed
    even if the substring 'signal=' still appears elsewhere in the file.
    """
    import datetime as dt

    # Email: direct_to, on_cc, on_bcc, flagged, is_newsletter, has_attachments.
    # `mentioned` is currently always False because $select=mentionsPreview is
    # rejected by Microsoft Graph v1.0 ("Could not find a property named
    # 'mentionsPreview' on type 'Microsoft.OutlookServices.Message'") — see
    # mcp-tool-responses/v1/audit/inbox-triage/probe3_select_field_isolation.json
    import os as _os

    _os.environ["MICROSOFT_MCP_ACCOUNT_ID"] = "me@example.com"

    raw_emails = [
        {
            "id": "m-direct",
            "subject": "Direct ask",
            "isRead": False,
            "toRecipients": [{"emailAddress": {"address": "me@example.com"}}],
        },
        {
            "id": "m-cc",
            "subject": "FYI cc",
            "isRead": False,
            "toRecipients": [{"emailAddress": {"address": "team@example.com"}}],
            "ccRecipients": [{"emailAddress": {"address": "me@example.com"}}],
        },
        {
            "id": "m-bcc",
            "subject": "Quiet copy",
            "isRead": False,
            "toRecipients": [{"emailAddress": {"address": "team@example.com"}}],
            "bccRecipients": [{"emailAddress": {"address": "me@example.com"}}],
        },
        {
            "id": "m-flag",
            "subject": "Action",
            "isRead": True,
            "flag": {"flagStatus": "flagged"},
        },
        {
            "id": "m-news",
            "subject": "Digest",
            "isRead": False,
            "from": {"emailAddress": {"address": "noreply@substack.com"}},
        },
        {
            "id": "m-attach",
            "subject": "see attached",
            "isRead": True,
            "hasAttachments": True,
        },
    ]
    items = tools_mod._emails_to_inbox_items(raw_emails)
    by_id = {it.id: it for it in items}
    assert by_id["m-direct"].direct_to is True, (
        "_emails_to_inbox_items lost the direct_to populator"
    )
    assert by_id["m-cc"].on_cc is True, (
        "_emails_to_inbox_items lost the on_cc populator"
    )
    assert by_id["m-bcc"].on_bcc is True, (
        "_emails_to_inbox_items lost the on_bcc populator"
    )
    assert by_id["m-flag"].flagged is True, (
        "_emails_to_inbox_items lost the flagged signal populator (B3b)"
    )
    assert by_id["m-news"].is_newsletter is True, (
        "_emails_to_inbox_items lost the is_newsletter signal populator (B3c)"
    )
    assert by_id["m-attach"].has_attachments is True, (
        "_emails_to_inbox_items lost the has_attachments populator"
    )
    # mentioned currently always False on this tenant (Graph v1.0 limitation).
    assert all(not it.mentioned for it in items)

    # Invite: starts_in_minutes
    future_iso = (
        dt.datetime.now(dt.timezone.utc) + dt.timedelta(minutes=10)
    ).isoformat()
    raw_invites = [
        {
            "id": "i-1",
            "subject": "Meeting",
            "meetingMessageType": "meetingRequest",
            "startDateTime": {"dateTime": future_iso},
            "isRead": False,
        }
    ]
    invite_items = tools_mod._invite_messages_to_inbox_items(raw_invites)
    assert invite_items[0].starts_in_minutes is not None, (
        "_invite_messages_to_inbox_items lost the starts_in_minutes populator (B3a)"
    )

    # Event: starts_in_minutes
    raw_events = [
        {
            "id": "e-1",
            "subject": "Standup",
            "start": {"dateTime": future_iso},
        }
    ]
    event_items = tools_mod._events_to_inbox_items(raw_events)
    assert event_items[0].starts_in_minutes is not None, (
        "_events_to_inbox_items lost the starts_in_minutes populator (B3a)"
    )


# ---------------------------------------------------------------------------
# Code-mode runtime invariants
# ---------------------------------------------------------------------------


def test_call_tool_chain_default_response_is_lean():
    """B2: call_tool_chain must NOT include the catalog by default."""
    tool = tools_mod.call_tool_chain
    fn = getattr(tool, "fn", tool)
    sig = inspect.signature(fn)
    assert "include_interfaces" in sig.parameters, (
        "call_tool_chain lost its include_interfaces parameter"
    )
    assert sig.parameters["include_interfaces"].default is False, (
        "include_interfaces default must remain False to keep token cost low"
    )


def test_code_mode_sandbox_supports_iteration():
    """B1: The sandbox's RestrictedPython guards must include _getiter_/_inplacevar_."""
    src = inspect.getsource(code_mode)
    for required_guard in ("_getiter_", "_iter_unpack_sequence_", "_inplacevar_"):
        assert f'"{required_guard}"' in src or f"'{required_guard}'" in src, (
            f"sandbox missing guard {required_guard!r} — list comprehensions / "
            "for loops / += would fail inside call_tool_chain"
        )


# ---------------------------------------------------------------------------
# Auth invariants
# ---------------------------------------------------------------------------


def test_scopes_has_no_duplicates():
    """B9: SCOPES must be unique."""
    assert len(auth.SCOPES) == len(set(auth.SCOPES)), "auth.SCOPES contains duplicates"


def test_no_bare_exception_raises_in_auth_modules():
    """A11: Auth modules must not raise bare Exception (use RuntimeError + from e)."""
    for module in (auth, auth_msal):
        source = inspect.getsource(module)
        offending = [
            (i, line)
            for i, line in enumerate(source.splitlines(), start=1)
            if "raise Exception(" in line
        ]
        assert not offending, f"{module.__name__} raises bare Exception:\n" + "\n".join(
            f"  L{i}: {line.strip()}" for i, line in offending
        )


def test_msal_account_identifier_assigned_exactly_once():
    """B13: __init__ assigns self.account_identifier exactly once (not twice as a rebase artifact)."""
    src = inspect.getsource(auth_msal.MSALRefreshTokenAuth.__init__)
    occurrences = src.count("self.account_identifier =")
    assert occurrences == 1, (
        f"MSAL __init__ has {occurrences} account_identifier assignments; expected 1"
    )


def test_response_shaping_does_not_export_dead_types():
    """A2: ResponseProfile enum and BudgetHints dataclass were unused; must stay removed."""
    import microsoft_mcp.response_shaping as rs

    assert not hasattr(rs, "ResponseProfile"), (
        "ResponseProfile was unused and removed in A2; do not re-export"
    )
    assert not hasattr(rs, "BudgetHints"), (
        "BudgetHints was unused and removed in A2; do not re-export"
    )


# ---------------------------------------------------------------------------
# UTCP bridge invariants
# ---------------------------------------------------------------------------


def test_utcp_bridge_command_not_user_specific():
    """A10: DEFAULT_BRIDGE_COMMAND must not hardcode any user's home directory.

    Inspect the source rather than the resolved runtime value: the resolved
    value may legitimately point at a user-specific install path (e.g. a mise
    or nvm shim under ~/.local), which is fine because it came from
    shutil.which / the env override — not from a hardcoded string literal.

    The guard matches any `/Users/<name>/` or `/home/<name>/` literal, so it
    catches leaks from any developer's machine (not just the one who wrote
    the test).
    """
    from microsoft_mcp import utcp_bridge_config

    src = inspect.getsource(utcp_bridge_config)
    hardcoded_home = re.search(r"/(?:Users|home)/[A-Za-z0-9_.-]+/", src)
    assert hardcoded_home is None, (
        f"utcp_bridge_config.py contains a hardcoded user home path "
        f"({hardcoded_home.group(0)!r}); use shutil.which, Path.home(), "
        f"or the env override"
    )
    assert utcp_bridge_config.DEFAULT_BRIDGE_COMMAND, (
        "DEFAULT_BRIDGE_COMMAND must not be empty"
    )
