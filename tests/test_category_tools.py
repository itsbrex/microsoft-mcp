from unittest.mock import call, patch


def _master_category(
    category_id: str,
    display_name: str,
    color: str = "preset0",
):
    return {
        "id": category_id,
        "displayName": display_name,
        "color": color,
    }


@patch("microsoft_mcp.tools.graph")
def test_list_master_categories_returns_normalized_shape(mock_graph):
    from microsoft_mcp.tools import list_master_categories

    mock_graph.request_paginated.return_value = iter(
        [
            _master_category("cat-1", "Deal Correspondence", "preset0"),
            _master_category("cat-2", "EOD Update", "preset1"),
        ]
    )

    result = list_master_categories.fn()

    mock_graph.request_paginated.assert_called_once_with(
        "/me/outlook/masterCategories",
        params={
            "$top": 100,
            "$select": "id,displayName,color",
        },
        limit=100,
    )
    assert result == [
        {"id": "cat-1", "display_name": "Deal Correspondence", "color": "preset0"},
        {"id": "cat-2", "display_name": "EOD Update", "color": "preset1"},
    ]


@patch("microsoft_mcp.tools.graph")
def test_get_master_category_resolves_by_display_name(mock_graph):
    from microsoft_mcp.tools import get_master_category

    mock_graph.request_paginated.return_value = iter(
        [
            _master_category("cat-1", "Deal Correspondence", "preset0"),
        ]
    )

    result = get_master_category.fn("Deal Correspondence")

    assert result == {
        "id": "cat-1",
        "display_name": "Deal Correspondence",
        "color": "preset0",
    }


@patch("microsoft_mcp.tools.graph")
def test_create_master_category_uses_graph_endpoint(mock_graph):
    from microsoft_mcp.tools import create_master_category

    mock_graph.request.return_value = _master_category(
        "cat-1", "Deal Correspondence", "preset0"
    )

    result = create_master_category.fn(
        display_name="Deal Correspondence",
        color="preset0",
    )

    mock_graph.request.assert_called_once_with(
        "POST",
        "/me/outlook/masterCategories",
        json={"displayName": "Deal Correspondence", "color": "preset0"},
    )
    assert result == {
        "id": "cat-1",
        "display_name": "Deal Correspondence",
        "color": "preset0",
    }


@patch("microsoft_mcp.tools.graph")
def test_update_master_category_can_recolor(mock_graph):
    from microsoft_mcp.tools import update_master_category

    mock_graph.request_paginated.return_value = iter(
        [_master_category("cat-1", "Deal Correspondence", "preset0")]
    )
    mock_graph.request.return_value = _master_category(
        "cat-1", "Deal Correspondence", "preset7"
    )

    result = update_master_category.fn(
        category="Deal Correspondence",
        color="preset7",
    )

    mock_graph.request.assert_called_once_with(
        "PATCH",
        "/me/outlook/masterCategories/cat-1",
        json={"color": "preset7"},
    )
    assert result == {
        "id": "cat-1",
        "display_name": "Deal Correspondence",
        "color": "preset7",
    }


@patch("microsoft_mcp.tools.graph")
def test_delete_master_category_uses_category_id(mock_graph):
    from microsoft_mcp.tools import delete_master_category

    mock_graph.request_paginated.return_value = iter(
        [_master_category("cat-1", "Deal Correspondence", "preset0")]
    )
    mock_graph.request.return_value = None

    result = delete_master_category.fn("Deal Correspondence")

    mock_graph.request.assert_called_once_with(
        "DELETE",
        "/me/outlook/masterCategories/cat-1",
    )
    assert result == {
        "status": "deleted",
        "id": "cat-1",
        "display_name": "Deal Correspondence",
    }


@patch("microsoft_mcp.tools.graph")
def test_ensure_master_categories_creates_missing_categories(mock_graph):
    from microsoft_mcp.tools import ensure_master_categories

    mock_graph.request_paginated.return_value = iter(
        [_master_category("cat-2", "EOD Update", "preset1")]
    )
    mock_graph.request.side_effect = [
        _master_category("cat-1", "Deal Correspondence", "preset0"),
    ]

    result = ensure_master_categories.fn(
        categories=[
            {"display_name": "Deal Correspondence", "color": "preset0"},
            {"display_name": "EOD Update", "color": "preset1"},
        ]
    )

    assert mock_graph.request.call_args_list == [
        call(
            "POST",
            "/me/outlook/masterCategories",
            json={"displayName": "Deal Correspondence", "color": "preset0"},
        )
    ]
    assert result["created"] == [
        {"id": "cat-1", "display_name": "Deal Correspondence", "color": "preset0"}
    ]
    assert result["existing"] == [
        {"id": "cat-2", "display_name": "EOD Update", "color": "preset1"}
    ]
