from unittest.mock import patch

from src.microsoft_mcp import tools


# ---------------------------------------------------------------------------
# list_focused_overrides
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools.graph.request_paginated")
def test_list_focused_overrides_assistant(mock_paged):
    mock_paged.return_value = iter(
        [
            {
                "id": "ov1",
                "classifyAs": "focused",
                "senderEmailAddress": {
                    "address": "alice@example.com",
                    "name": "Alice",
                },
            }
        ]
    )
    out = tools.list_focused_overrides.fn(response_profile="assistant")
    assert len(out) == 1
    assert out[0] == {
        "id": "ov1",
        "classify_as": "focused",
        "email": "alice@example.com",
        "name": "Alice",
    }
    call_args = mock_paged.call_args
    assert call_args[0][0] == "/me/inferenceClassification/overrides"


@patch("src.microsoft_mcp.tools.graph.request_paginated")
def test_list_focused_overrides_legacy(mock_paged):
    mock_paged.return_value = iter(
        [
            {
                "id": "ov2",
                "classifyAs": "other",
                "senderEmailAddress": {"address": "bob@example.com", "name": "Bob"},
                "@odata.etag": "should-strip",
            }
        ]
    )
    out = tools.list_focused_overrides.fn(response_profile="legacy")
    assert out[0]["id"] == "ov2"
    assert "@odata.etag" not in out[0]


# ---------------------------------------------------------------------------
# create_focused_override
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools.graph.request")
def test_create_focused_override(mock_req):
    mock_req.return_value = {
        "id": "new1",
        "classifyAs": "focused",
        "senderEmailAddress": {"address": "alice@example.com", "name": "Alice"},
    }
    out = tools.create_focused_override.fn(
        sender_email="alice@example.com",
        classify_as="focused",
        name="Alice",
    )
    mock_req.assert_called_once()
    args, kwargs = mock_req.call_args
    assert args[0] == "POST"
    assert args[1] == "/me/inferenceClassification/overrides"
    body = kwargs["json"]
    assert body["classifyAs"] == "focused"
    assert body["senderEmailAddress"]["address"] == "alice@example.com"
    assert body["senderEmailAddress"]["name"] == "Alice"
    assert out["id"] == "new1"


@patch("src.microsoft_mcp.tools.graph.request")
def test_create_focused_override_name_defaults_to_email(mock_req):
    mock_req.return_value = {
        "id": "new2",
        "classifyAs": "other",
        "senderEmailAddress": {"address": "bob@example.com", "name": "bob@example.com"},
    }
    tools.create_focused_override.fn(
        sender_email="bob@example.com", classify_as="other"
    )
    _, kwargs = mock_req.call_args
    assert kwargs["json"]["senderEmailAddress"]["name"] == "bob@example.com"


def test_create_focused_override_bad_classify_as_raises():
    import pytest

    with pytest.raises(ValueError, match="classify_as"):
        tools.create_focused_override.fn(
            sender_email="x@example.com", classify_as="invalid"
        )


# ---------------------------------------------------------------------------
# update_focused_override
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools.graph.request")
def test_update_focused_override(mock_req):
    mock_req.return_value = {
        "id": "ov1",
        "classifyAs": "other",
        "senderEmailAddress": {"address": "alice@example.com", "name": "Alice"},
    }
    out = tools.update_focused_override.fn(override_id="ov1", classify_as="other")
    args, kwargs = mock_req.call_args
    assert args[0] == "PATCH"
    assert args[1] == "/me/inferenceClassification/overrides/ov1"
    assert kwargs["json"] == {"classifyAs": "other"}
    assert out["id"] == "ov1"


def test_update_focused_override_bad_classify_as_raises():
    import pytest

    with pytest.raises(ValueError, match="classify_as"):
        tools.update_focused_override.fn(override_id="ov1", classify_as="bogus")


# ---------------------------------------------------------------------------
# delete_focused_override
# ---------------------------------------------------------------------------


@patch("src.microsoft_mcp.tools.graph.request")
def test_delete_focused_override(mock_req):
    mock_req.return_value = None
    out = tools.delete_focused_override.fn(override_id="ov1")
    args, _ = mock_req.call_args
    assert args[0] == "DELETE"
    assert args[1] == "/me/inferenceClassification/overrides/ov1"
    assert out == {"status": "deleted", "override_id": "ov1"}
