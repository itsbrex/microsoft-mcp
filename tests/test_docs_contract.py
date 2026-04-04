from pathlib import Path


def test_readme_mentions_integrated_code_mode_surface():
    text = Path("README.md").read_text()
    assert "list_inbox_items" in text
    assert "list_invite_messages" in text
    assert "rsvp_to_invite_message" in text
    assert "call_tool_chain" in text
    assert "search_tools" in text
    assert "Integrated Code Mode Surface" in text
    assert "MICROSOFT_MCP_TOOL_MODE" in text
    assert "UTCP Bridge Config Generator" in text


def test_code_mode_docs_and_examples_exist():
    assert Path("docs/code-mode-inbox-orchestration.md").exists()
    assert Path("examples/code-mode/inbox_triage.py").exists()
