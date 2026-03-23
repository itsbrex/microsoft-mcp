from microsoft_mcp.response_shaping import ResponseProfile, BudgetHints


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
