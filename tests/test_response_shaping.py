from microsoft_mcp.response_shaping import ResponseProfile, BudgetHints, cleanup_graph_payload


def test_response_profile_defaults_to_assistant_summary():
    assert ResponseProfile.default_for_operation("list") == ResponseProfile.SUMMARY


def test_response_profile_defaults_to_detail_for_get():
    assert ResponseProfile.default_for_operation("get") == ResponseProfile.DETAIL


def test_response_profile_defaults_to_summary_for_search():
    assert ResponseProfile.default_for_operation("search") == ResponseProfile.SUMMARY


def test_budget_hints_exposes_body_and_item_limits():
    hints = BudgetHints.for_operation("list_emails")
    assert hints.include_body is False
    assert hints.max_items <= 25


# --- Task 2: cleanup_graph_payload ---


def test_cleanup_graph_payload_strips_odata_and_empty_values():
    raw = {
        "@odata.context": "x",
        "@odata.etag": "y",
        "displayName": "John",
        "mobilePhone": None,
        "otherAddress": {},
        "businessPhones": [],
    }
    assert cleanup_graph_payload(raw) == {"displayName": "John"}


def test_cleanup_graph_payload_keeps_false_and_zero():
    raw = {"isRead": False, "size": 0, "subject": "Test"}
    assert cleanup_graph_payload(raw) == raw


def test_cleanup_graph_payload_strips_nested_odata():
    raw = {
        "value": [
            {"@odata.type": "#microsoft.graph.message", "id": "1", "extra": None},
            {"id": "2", "tags": []},
        ]
    }
    assert cleanup_graph_payload(raw) == {"value": [{"id": "1"}, {"id": "2"}]}


def test_cleanup_graph_payload_strips_noise_keys():
    raw = {
        "id": "1",
        "changeKey": "abc",
        "parentFolderId": "xyz",
        "subject": "Hello",
    }
    assert cleanup_graph_payload(raw) == {"id": "1", "subject": "Hello"}
