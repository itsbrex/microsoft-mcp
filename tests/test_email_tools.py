from unittest.mock import call, patch


@patch("microsoft_mcp.tools.graph")
def test_list_emails_defaults_to_no_body(mock_graph):
    from microsoft_mcp.tools import list_emails

    mock_graph.request_paginated.return_value = iter(
        [
            {
                "id": "e-1",
                "subject": "Hello",
                "from": {"emailAddress": {"name": "JP", "address": "jp@x.com"}},
                "toRecipients": [{"emailAddress": {"address": "me@x.com"}}],
                "receivedDateTime": "2026-03-23T10:00:00Z",
                "isRead": False,
                "hasAttachments": False,
                "bodyPreview": "Quick note...",
                "conversationId": "conv-1",
            }
        ]
    )

    result = list_emails.fn(limit=10)
    assert "body" not in result[0]
    assert result[0]["from"] == "JP <jp@x.com>"
    assert "conversation_url" in result[0]


@patch("microsoft_mcp.tools.graph")
def test_list_emails_with_body_returns_detail_shape(mock_graph):
    from microsoft_mcp.tools import list_emails

    mock_graph.request_paginated.return_value = iter(
        [
            {
                "id": "e-2",
                "subject": "Details",
                "from": {"emailAddress": {"name": "A", "address": "a@x.com"}},
                "body": {"contentType": "text", "content": "Full body content here"},
                "receivedDateTime": "2026-03-23T10:00:00Z",
                "isRead": True,
                "hasAttachments": False,
                "conversationId": "conv-2",
            }
        ]
    )

    result = list_emails.fn(limit=10, include_body=True)
    assert "body" in result[0]


@patch("microsoft_mcp.tools.graph")
def test_list_emails_resolves_custom_folder_name_before_listing_messages(mock_graph):
    from microsoft_mcp.tools import list_emails

    mock_graph.request_paginated.side_effect = [
        iter(
            [
                {
                    "id": "folder-1",
                    "displayName": "Cresa Deals of the Week",
                    "childFolderCount": 0,
                    "totalItemCount": 12,
                    "unreadItemCount": 2,
                }
            ]
        ),
        iter(
            [
                {
                    "id": "e-9",
                    "subject": "Deals of the Week 3/27",
                    "from": {
                        "emailAddress": {
                            "name": "Cresa Communications",
                            "address": "communications@cresa.email",
                        }
                    },
                    "receivedDateTime": "2026-03-27T21:04:56Z",
                    "isRead": False,
                    "conversationId": "conv-9",
                }
            ]
        ),
    ]

    result = list_emails.fn(folder="Cresa Deals of the Week", limit=1)

    assert mock_graph.request_paginated.call_args_list == [
        call(
            "/me/mailFolders",
            params={
                "$top": 100,
                "$select": "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,isHidden",
            },
            limit=500,
        ),
        call(
            "/me/mailFolders/folder-1/messages",
            params={
                "$top": 1,
                "$select": "id,subject,from,toRecipients,receivedDateTime,hasAttachments,bodyPreview,conversationId,isRead",
                "$orderby": "receivedDateTime desc",
            },
            limit=1,
        ),
    ]
    assert result[0]["id"] == "e-9"


@patch("microsoft_mcp.tools.graph")
def test_get_email_drops_body_preview_when_body_present(mock_graph):
    from microsoft_mcp.tools import get_email

    mock_graph.request.return_value = {
        "id": "e-3",
        "subject": "Test",
        "from": {"emailAddress": {"name": "X", "address": "x@x.com"}},
        "body": {"contentType": "text", "content": "Full body"},
        "bodyPreview": "Full body",
        "conversationId": "conv-3",
        "receivedDateTime": "2026-03-23T10:00:00Z",
        "isRead": True,
        "hasAttachments": False,
        "@odata.context": "junk",
    }

    result = get_email.fn("e-3")
    assert "bodyPreview" not in result
    assert "@odata.context" not in result
    assert "body" in result
