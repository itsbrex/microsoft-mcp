from pathlib import Path


def test_readme_mentions_inbox_tools_and_code_mode_usage():
    text = Path("README.md").read_text()
    assert "list_inbox_items" in text
    assert "Code Mode" in text
