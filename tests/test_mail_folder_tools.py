from unittest.mock import call, patch


def _folder(
    folder_id: str,
    display_name: str,
    *,
    parent_folder_id: str | None = None,
    child_folder_count: int = 0,
    total_item_count: int = 0,
    unread_item_count: int = 0,
):
    folder = {
        "id": folder_id,
        "displayName": display_name,
        "childFolderCount": child_folder_count,
        "totalItemCount": total_item_count,
        "unreadItemCount": unread_item_count,
    }
    if parent_folder_id:
        folder["parentFolderId"] = parent_folder_id
    return folder


@patch("microsoft_mcp.tools.graph")
def test_list_mail_folders_returns_normalized_shape(mock_graph):
    from microsoft_mcp.tools import list_mail_folders

    mock_graph.request_paginated.return_value = iter(
        [
            _folder(
                "folder-1",
                "Cresa Deals of the Week",
                total_item_count=12,
                unread_item_count=2,
            )
        ]
    )

    result = list_mail_folders.fn()

    mock_graph.request_paginated.assert_called_once_with(
        "/me/mailFolders",
        params={
            "$top": 100,
            "$select": "id,displayName,parentFolderId,childFolderCount,totalItemCount,unreadItemCount,isHidden",
        },
        limit=100,
    )
    assert result == [
        {
            "id": "folder-1",
            "display_name": "Cresa Deals of the Week",
            "parent_folder_id": None,
            "child_folder_count": 0,
            "total_item_count": 12,
            "unread_item_count": 2,
            "is_hidden": False,
        }
    ]


@patch("microsoft_mcp.tools.graph")
def test_get_mail_folder_resolves_display_name(mock_graph):
    from microsoft_mcp.tools import get_mail_folder

    mock_graph.request_paginated.return_value = iter(
        [_folder("folder-1", "Cresa Deals of the Week", total_item_count=12)]
    )

    result = get_mail_folder.fn("Cresa Deals of the Week")

    assert result["id"] == "folder-1"
    assert result["display_name"] == "Cresa Deals of the Week"


@patch("microsoft_mcp.tools.graph")
def test_create_mail_folder_under_root(mock_graph):
    from microsoft_mcp.tools import create_mail_folder

    mock_graph.request.return_value = _folder("folder-2", "Weekly Digest")

    result = create_mail_folder.fn("Weekly Digest")

    mock_graph.request.assert_called_once_with(
        "POST",
        "/me/mailFolders",
        json={"displayName": "Weekly Digest"},
    )
    assert result["id"] == "folder-2"
    assert result["display_name"] == "Weekly Digest"


@patch("microsoft_mcp.tools.graph")
def test_rename_mail_folder_updates_display_name(mock_graph):
    from microsoft_mcp.tools import rename_mail_folder

    mock_graph.request_paginated.return_value = iter([_folder("folder-3", "Old Name")])
    mock_graph.request.side_effect = [
        _folder("folder-3", "New Name"),
    ]

    result = rename_mail_folder.fn("Old Name", "New Name")

    assert mock_graph.request.call_args_list == [
        call(
            "PATCH",
            "/me/mailFolders/folder-3",
            json={"displayName": "New Name"},
        )
    ]
    assert result["display_name"] == "New Name"


@patch("microsoft_mcp.tools.graph")
def test_delete_mail_folder_uses_resolved_folder_id(mock_graph):
    from microsoft_mcp.tools import delete_mail_folder

    mock_graph.request_paginated.return_value = iter([_folder("folder-4", "Trash Me")])
    mock_graph.request.return_value = None

    result = delete_mail_folder.fn("Trash Me")

    mock_graph.request.assert_called_once_with("DELETE", "/me/mailFolders/folder-4")
    assert result == {"status": "deleted", "folder_id": "folder-4"}
