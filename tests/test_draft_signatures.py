"""Tests for signature injection in create_email_draft / update_email_draft."""

from __future__ import annotations

from unittest.mock import patch

import pytest

from microsoft_mcp import signatures


def _draft_message(draft_id: str, **overrides):
    draft = {
        "id": draft_id,
        "subject": "Draft",
        "from": {"emailAddress": {"name": "Me", "address": "me@example.com"}},
        "toRecipients": [{"emailAddress": {"address": "alice@example.com"}}],
        "ccRecipients": [],
        "bccRecipients": [],
        "body": {"contentType": "text", "content": "Draft body"},
        "bodyPreview": "Draft body",
        "conversationId": f"conv-{draft_id}",
        "webLink": f"https://outlook.office.com/mail/deeplink/compose/{draft_id}",
        "createdDateTime": "2026-04-02T10:00:00Z",
        "lastModifiedDateTime": "2026-04-02T10:01:00Z",
        "isDraft": True,
        "isRead": True,
    }
    draft.update(overrides)
    return draft


@pytest.fixture(autouse=True)
def _isolate(tmp_path, monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURES_DIR", str(tmp_path))
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    monkeypatch.delenv("MICROSOFT_MCP_DEFAULT_SIGNATURE", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_REPLY_SIGNATURE", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_SIGNATURE_RFC3676", raising=False)
    yield


# --- create_email_draft --------------------------------------------------


@patch("microsoft_mcp.tools.graph")
def test_signature_param_appends_to_new_draft(mock_graph):
    from microsoft_mcp.tools import create_email_draft

    signatures.write_signature("default", "Cheers,\nBrian")
    mock_graph.request.return_value = _draft_message("d-1")

    result = create_email_draft.fn(
        draft_type="new",
        to_recipients=["alice@example.com"],
        body="Hi there",
        signature="default",
    )

    args, kwargs = mock_graph.request.call_args
    assert args == ("POST", "/me/messages")
    assert kwargs["json"]["body"]["content"] == "Hi there\n\nCheers,\nBrian"
    assert result["signature_applied"] == {
        "account": "brian-work",
        "name": "default",
        "html": False,
    }
    assert "signature_warning" not in result


@patch("microsoft_mcp.tools.graph")
def test_env_default_used_when_no_signature_arg(mock_graph, monkeypatch):
    from microsoft_mcp.tools import create_email_draft

    monkeypatch.setenv("MICROSOFT_MCP_DEFAULT_SIGNATURE", "default")
    signatures.write_signature("default", "Brian")
    mock_graph.request.return_value = _draft_message("d-1")

    create_email_draft.fn(
        draft_type="new",
        to_recipients=["alice@example.com"],
        body="Hi",
    )

    _, kwargs = mock_graph.request.call_args
    assert kwargs["json"]["body"]["content"] == "Hi\n\nBrian"


@patch("microsoft_mcp.tools.graph")
def test_reply_uses_reply_env_when_set(mock_graph, monkeypatch):
    from microsoft_mcp.tools import create_email_draft

    monkeypatch.setenv("MICROSOFT_MCP_DEFAULT_SIGNATURE", "default")
    monkeypatch.setenv("MICROSOFT_MCP_REPLY_SIGNATURE", "replies")
    signatures.write_signature("default", "Brian")
    signatures.write_signature("replies", "B.")
    mock_graph.request.side_effect = [
        _draft_message("d-2"),
        _draft_message("d-2"),
    ]

    create_email_draft.fn(
        draft_type="reply",
        email_id="msg-1",
        body="Thanks",
    )

    # Second call is the PATCH; check the patched body.
    patch_call = mock_graph.request.call_args_list[1]
    assert patch_call.kwargs["json"]["body"]["content"] == "Thanks\n\nB."


@patch("microsoft_mcp.tools.graph")
def test_reply_falls_back_to_default_env_when_reply_var_unset(mock_graph, monkeypatch):
    from microsoft_mcp.tools import create_email_draft

    monkeypatch.setenv("MICROSOFT_MCP_DEFAULT_SIGNATURE", "default")
    signatures.write_signature("default", "Brian")
    mock_graph.request.side_effect = [
        _draft_message("d-2"),
        _draft_message("d-2"),
    ]

    create_email_draft.fn(
        draft_type="reply",
        email_id="msg-1",
        body="Thanks",
    )
    patch_call = mock_graph.request.call_args_list[1]
    assert patch_call.kwargs["json"]["body"]["content"] == "Thanks\n\nBrian"


@patch("microsoft_mcp.tools.graph")
def test_signature_none_suppresses_env_default(mock_graph, monkeypatch):
    from microsoft_mcp.tools import create_email_draft

    monkeypatch.setenv("MICROSOFT_MCP_DEFAULT_SIGNATURE", "default")
    signatures.write_signature("default", "Brian")
    mock_graph.request.return_value = _draft_message("d-1")

    result = create_email_draft.fn(
        draft_type="new",
        to_recipients=["alice@example.com"],
        body="Hi",
        signature="none",
    )

    _, kwargs = mock_graph.request.call_args
    assert kwargs["json"]["body"]["content"] == "Hi"
    assert "signature_applied" not in result


@patch("microsoft_mcp.tools.graph")
def test_missing_signature_surfaces_warning_and_creates_draft(mock_graph):
    from microsoft_mcp.tools import create_email_draft

    mock_graph.request.return_value = _draft_message("d-1")

    result = create_email_draft.fn(
        draft_type="new",
        to_recipients=["alice@example.com"],
        body="Hi",
        signature="missing-sig",
    )

    _, kwargs = mock_graph.request.call_args
    assert kwargs["json"]["body"]["content"] == "Hi"
    assert result["status"] == "draft_created"
    assert "signature not found" in result["signature_warning"]
    assert "signature_applied" not in result


@patch("microsoft_mcp.tools.graph")
def test_html_body_uses_html_sibling(mock_graph):
    from microsoft_mcp.tools import create_email_draft

    signatures.write_signature("default", "Brian (text)")
    signatures.write_signature("default", "<b>Brian</b>", html=True)
    mock_graph.request.return_value = _draft_message("d-1")

    result = create_email_draft.fn(
        draft_type="new",
        to_recipients=["alice@example.com"],
        body="<p>Hi</p>",
        body_content_type="html",
        signature="default",
    )
    _, kwargs = mock_graph.request.call_args
    assert kwargs["json"]["body"]["content"] == "<p>Hi</p>\n\n<b>Brian</b>"
    assert result["signature_applied"]["html"] is True


@patch("microsoft_mcp.tools.graph")
def test_html_body_falls_back_to_converted_text(mock_graph):
    from microsoft_mcp.tools import create_email_draft

    signatures.write_signature("default", "Cheers,\nBrian")
    mock_graph.request.return_value = _draft_message("d-1")

    create_email_draft.fn(
        draft_type="new",
        to_recipients=["alice@example.com"],
        body="<p>Hi</p>",
        body_content_type="html",
        signature="default",
    )
    _, kwargs = mock_graph.request.call_args
    content = kwargs["json"]["body"]["content"]
    assert content.startswith("<p>Hi</p>\n\n")
    assert '<div class="signature">' in content
    assert "Cheers,<br>\nBrian" in content


# --- update_email_draft --------------------------------------------------


@patch("microsoft_mcp.tools.graph")
def test_update_email_draft_applies_signature_when_body_supplied(mock_graph):
    from microsoft_mcp.tools import update_email_draft

    signatures.write_signature("default", "Brian")
    mock_graph.request.return_value = _draft_message("d-1")

    result = update_email_draft.fn(
        email_id="d-1",
        body="Revised body",
        signature="default",
    )

    _, kwargs = mock_graph.request.call_args
    assert kwargs["json"]["body"]["content"] == "Revised body\n\nBrian"
    assert result["signature_applied"]["name"] == "default"


@patch("microsoft_mcp.tools.graph")
def test_update_email_draft_skips_signature_when_no_body(mock_graph, monkeypatch):
    from microsoft_mcp.tools import update_email_draft

    monkeypatch.setenv("MICROSOFT_MCP_DEFAULT_SIGNATURE", "default")
    signatures.write_signature("default", "Brian")
    mock_graph.request.return_value = _draft_message("d-1")

    result = update_email_draft.fn(
        email_id="d-1",
        subject="New subject",
    )

    _, kwargs = mock_graph.request.call_args
    # body must NOT appear in the PATCH payload when body wasn't supplied.
    assert "body" not in kwargs["json"]
    assert "signature_applied" not in result


# --- read-only MCP tools -------------------------------------------------


def test_list_signatures_tool_returns_records():
    from microsoft_mcp.tools import list_signatures

    signatures.write_signature("default", "Brian")
    rows = list_signatures.fn()
    assert len(rows) == 1
    assert rows[0]["account"] == "brian-work"
    assert rows[0]["name"] == "default"


def test_get_signature_tool_returns_content():
    from microsoft_mcp.tools import get_signature

    signatures.write_signature("default", "Brian")
    result = get_signature.fn("default")
    assert result["status"] == "ok"
    assert result["content"] == "Brian"
    assert result["account"] == "brian-work"


def test_get_signature_tool_returns_not_found():
    from microsoft_mcp.tools import get_signature

    result = get_signature.fn("missing")
    assert result["status"] == "not_found"
    assert result["name"] == "missing"


def test_get_signature_tool_returns_error_for_bad_name():
    from microsoft_mcp.tools import get_signature

    result = get_signature.fn("bad name!")
    assert result["status"] == "error"
    assert "invalid" in result["error"].lower()
