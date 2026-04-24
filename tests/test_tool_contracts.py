from unittest.mock import patch


@patch("microsoft_mcp.tools.graph")
def test_get_user_details_strips_odata_and_nulls(mock_graph):
    from microsoft_mcp.tools import get_user_details

    mock_graph.request.return_value = {
        "@odata.context": "https://graph.microsoft.com/v1.0/$metadata#users/$entity",
        "id": "u-1",
        "displayName": "John",
        "mail": "john@example.com",
        "mobilePhone": None,
        "businessPhones": [],
    }

    result = get_user_details.fn()
    assert "@odata.context" not in result
    assert "mobilePhone" not in result
    assert result["displayName"] == "John"


@patch("microsoft_mcp.tools.graph")
def test_check_availability_returns_compact_schedule(mock_graph):
    from microsoft_mcp.tools import check_availability

    mock_graph.request.side_effect = [
        {"mail": "me@example.com"},  # GET /me
        {
            "@odata.context": "https://graph.microsoft.com/...",
            "value": [
                {
                    "scheduleId": "me@example.com",
                    "availabilityView": "0000220000",
                    "scheduleItems": [
                        {
                            "status": "busy",
                            "subject": "Meeting",
                            "start": {"dateTime": "2026-03-23T14:00:00"},
                            "end": {"dateTime": "2026-03-23T15:00:00"},
                        }
                    ],
                    "workingHours": {
                        "daysOfWeek": [
                            "monday",
                            "tuesday",
                            "wednesday",
                            "thursday",
                            "friday",
                        ],
                        "startTime": "09:00:00",
                        "endTime": "17:00:00",
                    },
                }
            ],
        },
    ]

    result = check_availability.fn("2026-03-23T09:00:00Z", "2026-03-23T17:00:00Z")
    assert "participants" in result
    assert "@odata.context" not in result
    assert "value" not in result


@patch("microsoft_mcp.tools.graph")
def test_list_contacts_summary_mode_is_compact(mock_graph):
    from microsoft_mcp.tools import list_contacts

    mock_graph.request_paginated.return_value = iter(
        [
            {
                "@odata.etag": "abc",
                "id": "c-1",
                "displayName": "Brian",
                "jobTitle": "Dev",
                "companyName": "Acme",
                "emailAddresses": [{"address": "brian@acme.com"}],
                "businessPhones": ["+1234"],
                "mobilePhone": "+5678",
                "homePhones": [],
                "personalNotes": None,
                "changeKey": "xyz",
            }
        ]
    )

    result = list_contacts.fn(limit=5)
    first = result[0]
    assert "@odata.etag" not in first
    assert "changeKey" not in first
    assert "personalNotes" not in first
    assert first["email_addresses"] == ["brian@acme.com"]
    assert first["displayName"] == "Brian"


@patch("microsoft_mcp.tools.graph")
def test_get_contact_returns_detailed_shape(mock_graph):
    from microsoft_mcp.tools import get_contact

    mock_graph.request.return_value = {
        "@odata.context": "https://graph.microsoft.com/...",
        "id": "c-2",
        "displayName": "Alice",
        "emailAddresses": [{"address": "alice@x.com"}],
        "jobTitle": "Manager",
        "businessAddress": {"street": "123 Main", "city": "NYC"},
        "changeKey": "abc",
    }

    result = get_contact.fn("c-2")
    assert "@odata.context" not in result
    assert result["email_addresses"] == ["alice@x.com"]
    assert result["businessAddress"]["city"] == "NYC"


def test_convert_to_markdown_is_private():
    """Helper that isn't @mcp.tool-decorated must have a leading underscore to signal intent."""
    from microsoft_mcp import tools as tools_mod

    assert not hasattr(tools_mod, "convert_to_markdown"), (
        "convert_to_markdown should be renamed to _convert_to_markdown (no leading underscore implies MCP tool)"
    )
    assert hasattr(tools_mod, "_convert_to_markdown"), (
        "_convert_to_markdown should exist as the renamed helper"
    )
