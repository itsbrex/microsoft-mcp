from unittest.mock import call, patch

import pytest
import httpx


def _managed_email(email_id: str, **overrides):
    email = {
        "id": email_id,
        "subject": f"Email {email_id}",
        "from": {"emailAddress": {"name": "Sender", "address": "sender@example.com"}},
        "receivedDateTime": "2026-04-02T10:00:00Z",
        "isRead": False,
        "conversationId": f"conv-{email_id}",
    }
    email.update(overrides)
    return email


@patch("microsoft_mcp.tools.graph")
def test_mark_email_read_updates_read_state(mock_graph):
    from microsoft_mcp.tools import mark_email_read

    mock_graph.request.return_value = _managed_email("e-1", isRead=True)

    result = mark_email_read.fn(email_id="e-1", is_read=True)

    mock_graph.request.assert_called_once_with(
        "PATCH",
        "/me/messages/e-1",
        json={"isRead": True},
    )
    assert result["status"] == "updated"
    assert result["email"]["id"] == "e-1"
    assert result["email"]["is_read"] is True


@patch("microsoft_mcp.tools.graph")
def test_set_email_categories_replaces_categories(mock_graph):
    from microsoft_mcp.tools import set_email_categories

    mock_graph.request.return_value = _managed_email(
        "e-2",
        categories=["Focus", "Follow Up"],
    )

    result = set_email_categories.fn(
        email_id="e-2",
        categories=["Focus", "Follow Up"],
    )

    mock_graph.request.assert_called_once_with(
        "PATCH",
        "/me/messages/e-2",
        json={"categories": ["Focus", "Follow Up"]},
    )
    assert result["status"] == "updated"
    assert result["categories"] == ["Focus", "Follow Up"]


@patch("microsoft_mcp.tools.graph")
def test_move_email_uses_well_known_folder_alias(mock_graph):
    from microsoft_mcp.tools import move_email

    mock_graph.request.return_value = _managed_email("e-3")

    result = move_email.fn(email_id="e-3", destination_folder="deleted")

    mock_graph.request.assert_called_once_with(
        "POST",
        "/me/messages/e-3/move",
        json={"destinationId": "deleteditems"},
    )
    assert result["status"] == "moved"
    assert result["destination_folder"] == "deleteditems"


@patch("microsoft_mcp.tools.graph")
def test_move_email_resolves_custom_folder_name(mock_graph):
    from microsoft_mcp.tools import move_email

    mock_graph.request_paginated.return_value = iter(
        [
            {
                "id": "folder-1",
                "displayName": "Cresa Deals of the Week",
                "childFolderCount": 0,
                "totalItemCount": 12,
                "unreadItemCount": 2,
            }
        ]
    )
    mock_graph.request.return_value = _managed_email("e-3")

    result = move_email.fn(
        email_id="e-3",
        destination_folder="Cresa Deals of the Week",
    )

    assert mock_graph.request.call_args_list == [
        call(
            "POST",
            "/me/messages/e-3/move",
            json={"destinationId": "folder-1"},
        )
    ]
    assert result["status"] == "moved"
    assert result["destination_folder"] == "folder-1"


@patch("microsoft_mcp.tools.graph")
def test_archive_email_moves_to_archive(mock_graph):
    from microsoft_mcp.tools import archive_email

    mock_graph.request.return_value = _managed_email("e-4")

    result = archive_email.fn(email_id="e-4")

    mock_graph.request.assert_called_once_with(
        "POST",
        "/me/messages/e-4/move",
        json={"destinationId": "archive"},
    )
    assert result["status"] == "moved"
    assert result["destination_folder"] == "archive"


@patch("microsoft_mcp.tools.graph")
def test_delete_email_removes_message(mock_graph):
    from microsoft_mcp.tools import delete_email

    mock_graph.request.return_value = None

    result = delete_email.fn(email_id="e-5")

    mock_graph.request.assert_called_once_with("DELETE", "/me/messages/e-5")
    assert result == {"status": "deleted", "email_id": "e-5", "resource": "message"}


@patch("microsoft_mcp.tools.graph")
def test_delete_email_falls_back_to_calendar_event_delete(mock_graph):
    from microsoft_mcp.tools import delete_email

    request = httpx.Request(
        "DELETE", "https://graph.microsoft.com/v1.0/me/messages/evt-1"
    )
    response = httpx.Response(400, request=request)

    mock_graph.request.side_effect = [
        httpx.HTTPStatusError("bad request", request=request, response=response),
        None,
    ]

    result = delete_email.fn(email_id="evt-1")

    assert mock_graph.request.call_args_list == [
        call("DELETE", "/me/messages/evt-1"),
        call("DELETE", "/me/events/evt-1"),
    ]
    assert result == {
        "status": "deleted",
        "email_id": "evt-1",
        "resource": "event",
    }


@patch("microsoft_mcp.tools.graph")
def test_bulk_manage_emails_processes_archive_actions(mock_graph):
    from microsoft_mcp.tools import bulk_manage_emails

    mock_graph.request.side_effect = [
        _managed_email("e-10"),
        _managed_email("e-11"),
    ]

    result = bulk_manage_emails.fn(
        email_ids=["e-10", "e-11"],
        action="archive",
    )

    assert mock_graph.request.call_args_list == [
        call("POST", "/me/messages/e-10/move", json={"destinationId": "archive"}),
        call("POST", "/me/messages/e-11/move", json={"destinationId": "archive"}),
    ]
    assert result["action"] == "archive"
    assert result["requested"] == 2
    assert result["succeeded"] == 2
    assert result["failed"] == 0


@patch("microsoft_mcp.tools.graph")
def test_bulk_manage_emails_reports_partial_failures(mock_graph):
    from microsoft_mcp.tools import bulk_manage_emails

    mock_graph.request.side_effect = [
        _managed_email("e-20"),
        RuntimeError("boom"),
    ]

    result = bulk_manage_emails.fn(
        email_ids=["e-20", "e-21"],
        action="mark_read",
    )

    assert result["requested"] == 2
    assert result["succeeded"] == 1
    assert result["failed"] == 1
    assert result["results"][1]["status"] == "failed"
    assert "boom" in result["results"][1]["error"]


@patch("microsoft_mcp.tools.graph")
def test_bulk_manage_emails_delete_action_handles_calendar_event_ids(mock_graph):
    from microsoft_mcp.tools import bulk_manage_emails

    request = httpx.Request(
        "DELETE", "https://graph.microsoft.com/v1.0/me/messages/evt-9"
    )
    response = httpx.Response(400, request=request)

    mock_graph.request.side_effect = [
        httpx.HTTPStatusError("bad request", request=request, response=response),
        None,
    ]

    result = bulk_manage_emails.fn(
        email_ids=["evt-9"],
        action="delete",
    )

    assert mock_graph.request.call_args_list == [
        call("DELETE", "/me/messages/evt-9"),
        call("DELETE", "/me/events/evt-9"),
    ]
    assert result["succeeded"] == 1
    assert result["results"][0]["resource"] == "event"


def test_bulk_manage_emails_requires_supported_action():
    from microsoft_mcp.tools import bulk_manage_emails

    with pytest.raises(ValueError, match="Unsupported action"):
        bulk_manage_emails.fn(email_ids=["e-1"], action="snooze")
