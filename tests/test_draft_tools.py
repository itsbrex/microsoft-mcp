from unittest.mock import call, patch

import pytest


@pytest.fixture(autouse=True)
def _clear_signature_env(monkeypatch):
    """Force signature env vars to empty so drafts don't get a real local
    signature appended during tests. The repo's .env file sets
    MICROSOFT_MCP_DEFAULT_SIGNATURE; ``load_dotenv()`` calls inside auth
    helpers can re-populate after ``monkeypatch.delenv``, so we set them
    to empty strings (which python-dotenv preserves rather than overwrites)."""
    for var in (
        "MICROSOFT_MCP_DEFAULT_SIGNATURE",
        "MICROSOFT_MCP_REPLY_SIGNATURE",
        "MICROSOFT_MCP_SIGNATURE_ACCOUNT",
        "MICROSOFT_MCP_SIGNATURES_DIR",
    ):
        monkeypatch.setenv(var, "")


def _draft_message(draft_id: str, **overrides):
    draft = {
        "id": draft_id,
        "subject": "Draft subject",
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


@patch("microsoft_mcp.tools.graph")
def test_create_email_draft_creates_new_message_draft(mock_graph):
    from microsoft_mcp.tools import create_email_draft

    mock_graph.request.return_value = _draft_message(
        "d-1",
        subject="Quarterly update",
        toRecipients=[
            {"emailAddress": {"address": "alice@example.com"}},
            {"emailAddress": {"address": "bob@example.com"}},
        ],
        ccRecipients=[{"emailAddress": {"address": "manager@example.com"}}],
        body={"contentType": "html", "content": "<p>Status update</p>"},
    )

    result = create_email_draft.fn(
        draft_type="new",
        to_recipients=["alice@example.com", "bob@example.com"],
        cc_recipients=["manager@example.com"],
        subject="Quarterly update",
        body="<p>Status update</p>",
        body_content_type="html",
    )

    mock_graph.request.assert_called_once_with(
        "POST",
        "/me/messages",
        json={
            "subject": "Quarterly update",
            "toRecipients": [
                {"emailAddress": {"address": "alice@example.com"}},
                {"emailAddress": {"address": "bob@example.com"}},
            ],
            "ccRecipients": [{"emailAddress": {"address": "manager@example.com"}}],
            "body": {"contentType": "html", "content": "<p>Status update</p>"},
        },
    )
    assert result["status"] == "draft_created"
    assert result["draft_type"] == "new"
    assert result["draft_id"] == "d-1"
    assert result["draft"]["to"] == ["alice@example.com", "bob@example.com"]
    assert result["draft"]["cc"] == ["manager@example.com"]
    assert result["draft"]["is_draft"] is True


@patch("microsoft_mcp.tools.graph")
def test_create_email_draft_creates_reply_draft_and_updates_body(mock_graph):
    from microsoft_mcp.tools import create_email_draft

    mock_graph.request.side_effect = [
        _draft_message(
            "d-2", body={"contentType": "text", "content": "Original quote"}
        ),
        _draft_message(
            "d-2", body={"contentType": "text", "content": "Thanks for the update."}
        ),
    ]

    result = create_email_draft.fn(
        draft_type="reply",
        email_id="msg-1",
        body="Thanks for the update.",
    )

    assert mock_graph.request.call_args_list == [
        call("POST", "/me/messages/msg-1/createReply"),
        call(
            "PATCH",
            "/me/messages/d-2",
            json={"body": {"contentType": "text", "content": "Thanks for the update."}},
        ),
    ]
    assert result["draft_type"] == "reply"
    assert result["reply_to_message_id"] == "msg-1"
    assert result["draft"]["body"]["content"] == "Thanks for the update."


@patch("microsoft_mcp.tools.graph")
def test_create_email_draft_creates_reply_all_draft(mock_graph):
    from microsoft_mcp.tools import create_email_draft

    mock_graph.request.side_effect = [
        _draft_message(
            "d-3",
            toRecipients=[{"emailAddress": {"address": "team@example.com"}}],
            ccRecipients=[{"emailAddress": {"address": "lead@example.com"}}],
        ),
        _draft_message(
            "d-3",
            toRecipients=[{"emailAddress": {"address": "team@example.com"}}],
            ccRecipients=[{"emailAddress": {"address": "lead@example.com"}}],
            body={"contentType": "text", "content": "Replying to everyone."},
        ),
    ]

    result = create_email_draft.fn(
        draft_type="reply_all",
        email_id="msg-2",
        body="Replying to everyone.",
    )

    assert mock_graph.request.call_args_list == [
        call("POST", "/me/messages/msg-2/createReplyAll"),
        call(
            "PATCH",
            "/me/messages/d-3",
            json={"body": {"contentType": "text", "content": "Replying to everyone."}},
        ),
    ]
    assert result["draft_type"] == "reply_all"
    assert result["draft"]["to"] == ["team@example.com"]
    assert result["draft"]["cc"] == ["lead@example.com"]


def test_create_email_draft_requires_recipients_for_new_messages():
    from microsoft_mcp.tools import create_email_draft

    with pytest.raises(ValueError, match="at least one recipient"):
        create_email_draft.fn(draft_type="new", subject="Missing recipients")


def test_create_email_draft_requires_email_id_for_reply_modes():
    from microsoft_mcp.tools import create_email_draft

    with pytest.raises(ValueError, match="email_id is required"):
        create_email_draft.fn(draft_type="reply", body="Hello")
