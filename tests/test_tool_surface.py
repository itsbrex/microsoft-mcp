"""Tool-surface regression for the outlook-creds mail port.

Guards against accidental de-registration of the MCP tools added across the
mail-port waves (inbox rules, Focused Inbox overrides, reply/forward drafts,
MailTips, attachments, Microsoft To-Do, email templates, signatures, intel
reports, bounce scanning). These tool names are auth-method independent (none
are Teams tools), so they must be registered under both the azure and msal
auth paths.

Import convention follows the repo pattern:
    from src.microsoft_mcp import <module>
"""

import asyncio

from src.microsoft_mcp import tools

# Mail-port tools, grouped by wave/family. Each name MUST stay registered.
EXPECTED_MAIL_PORT_TOOLS = {
    # Wave 2-3: inbox rules
    "list_inbox_rules",
    "get_inbox_rule",
    "create_inbox_rule",
    "update_inbox_rule",
    "delete_inbox_rule",
    "toggle_inbox_rule",
    "reorder_inbox_rules",
    "export_inbox_rules",
    "import_inbox_rules",
    # Wave 4: Focused Inbox overrides
    "list_focused_overrides",
    "create_focused_override",
    "update_focused_override",
    "delete_focused_override",
    # Wave 5: reply/forward drafts + explicit send + mailtips + attachments
    "reply_email_draft",
    "reply_all_email_draft",
    "forward_email_draft",
    "send_email_draft",
    "get_mailtips",
    "list_attachments",
    "download_attachments",
    # Wave 6: Microsoft To-Do
    "list_todo_lists",
    "create_todo_list",
    "list_tasks",
    "create_task",
    "update_task",
    "complete_task",
    "delete_task",
    "create_task_from_email",
    # Wave 7: email templates
    "list_email_templates",
    "render_email_template",
    "find_template_variables",
    "get_template_placeholders",
    "substitute_template_variables",
    # Wave 8: signatures + intel
    "parse_email_signature",
    "normalize_phone_number",
    "generate_morning_briefing",
    "get_priority_signals",
    "get_contact_intelligence",
    "get_end_of_day_recap",
    # Wave 9: bounces
    "scan_bounces",
}


def _registered_tool_names() -> set[str]:
    # FastMCP 3.x removed mcp._tool_manager. The local provider's list_tools()
    # returns every registered tool (including those disabled for direct MCP
    # exposure), matching the registration-coverage intent of this guard.
    return {tool.name for tool in asyncio.run(tools.mcp._local_provider.list_tools())}


def test_all_mail_port_tools_registered() -> None:
    registered = _registered_tool_names()
    missing = EXPECTED_MAIL_PORT_TOOLS - registered
    assert not missing, f"mail-port tools not registered: {sorted(missing)}"


def test_tool_surface_has_no_unexpected_dropouts() -> None:
    # A coarse floor: the mail port pushed the registered tool count well past
    # the pre-port baseline. Guard against a wholesale registration regression.
    registered = _registered_tool_names()
    assert len(registered) >= len(EXPECTED_MAIL_PORT_TOOLS) + 30
