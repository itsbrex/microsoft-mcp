import json
from pathlib import Path

from microsoft_mcp.utcp_bridge_config import (
    build_bridge_claude_config,
    build_utcp_config,
    convert_claude_config,
    derive_manual_name,
)


def sample_claude_config():
    return {
        "mcpServers": {
            "microsoft-mcp": {
                "command": "/path/to/uv",
                "args": ["run", "microsoft-mcp"],
            },
            "github": {
                "command": "docker",
                "args": ["run", "-i", "--rm", "mcp/github"],
            },
            "notion-mcp": {
                "command": "npx",
                "args": ["@notionhq/notion-mcp-server"],
            },
        }
    }


def sample_claude_config_with_bridge():
    payload = sample_claude_config()
    payload["mcpServers"]["code-mode-mcp"] = {
        "command": "npx",
        "args": ["@utcp/code-mode-mcp"],
    }
    return payload


def test_build_utcp_config_supports_include_and_exclude():
    utcp_config, mappings = build_utcp_config(
        sample_claude_config(),
        include_servers=["microsoft-mcp", "github"],
        exclude_servers=["github"],
    )

    assert [mapping.source_server_name for mapping in mappings] == ["microsoft-mcp"]
    assert [manual["name"] for manual in utcp_config["manual_call_templates"]] == [
        "microsoft_mcp"
    ]


def test_build_utcp_config_skips_existing_bridge_server_by_default():
    utcp_config, mappings = build_utcp_config(sample_claude_config_with_bridge())

    assert [mapping.source_server_name for mapping in mappings] == [
        "microsoft-mcp",
        "github",
        "notion-mcp",
    ]
    manual_names = [manual["name"] for manual in utcp_config["manual_call_templates"]]
    assert "code_mode_mcp" not in manual_names


def test_build_utcp_config_can_explicitly_include_bridge_server():
    utcp_config, mappings = build_utcp_config(
        sample_claude_config_with_bridge(),
        include_servers=["code-mode-mcp"],
    )

    assert [mapping.source_server_name for mapping in mappings] == ["code-mode-mcp"]
    assert [manual["name"] for manual in utcp_config["manual_call_templates"]] == [
        "code_mode_mcp"
    ]


def test_derive_manual_name_prefers_concise_names():
    assert derive_manual_name("google_sheets") == "google_sheets"
    assert derive_manual_name("microsoft-mcp") == "microsoft_mcp"
    assert derive_manual_name("code-mode-mcp") == "code_mode_mcp"


def test_build_bridge_config_points_to_generated_utcp_file(tmp_path: Path):
    utcp_config_path = tmp_path / ".utcp_config.json"

    bridge_config = build_bridge_claude_config(
        utcp_config_path,
        bridge_name="code-mode-mcp",
    )

    assert list(bridge_config["mcpServers"]) == ["code-mode-mcp"]
    assert (
        bridge_config["mcpServers"]["code-mode-mcp"]["command"]
        == "/Users/hack/.local/share/mise/shims/npx"
    )
    assert bridge_config["mcpServers"]["code-mode-mcp"]["env"][
        "UTCP_CONFIG_FILE"
    ] == str(utcp_config_path.resolve())


def test_convert_claude_config_is_non_destructive_and_writes_outputs(tmp_path: Path):
    source = tmp_path / "claude_desktop_config.json"
    source_payload = sample_claude_config()
    source.write_text(json.dumps(source_payload, indent=2) + "\n")

    outputs = convert_claude_config(
        source,
        output_dir=tmp_path / "generated",
        include_servers=["microsoft-mcp", "github"],
    )

    assert json.loads(source.read_text()) == source_payload
    assert outputs["utcp_config"].exists()
    assert outputs["bridge_config"].exists()
    assert outputs["manual_map"].exists()

    utcp_payload = json.loads(outputs["utcp_config"].read_text())
    assert len(utcp_payload["manual_call_templates"]) == 2
    assert [
        entry["source_server_name"]
        for entry in json.loads(outputs["manual_map"].read_text())
    ] == ["microsoft-mcp", "github"]


def test_build_utcp_config_can_override_server_env():
    utcp_config, _ = build_utcp_config(
        sample_claude_config(),
        include_servers=["microsoft-mcp"],
        server_env_overrides={
            "microsoft-mcp": {
                "MICROSOFT_MCP_TOOL_MODE": "hybrid",
            }
        },
    )

    wrapped_env = utcp_config["manual_call_templates"][0]["config"]["mcpServers"][
        "microsoft-mcp"
    ]["env"]
    assert wrapped_env["MICROSOFT_MCP_TOOL_MODE"] == "hybrid"


def test_build_utcp_config_sets_default_stdio_transport_for_command_servers():
    utcp_config, _ = build_utcp_config(
        sample_claude_config(),
        include_servers=["microsoft-mcp"],
    )

    wrapped_server = utcp_config["manual_call_templates"][0]["config"]["mcpServers"][
        "microsoft-mcp"
    ]
    assert wrapped_server["transport"] == "stdio"


def test_build_utcp_config_preserves_explicit_transport():
    payload = sample_claude_config()
    payload["mcpServers"]["microsoft-mcp"]["transport"] = "stdio"

    utcp_config, _ = build_utcp_config(
        payload,
        include_servers=["microsoft-mcp"],
    )

    wrapped_server = utcp_config["manual_call_templates"][0]["config"]["mcpServers"][
        "microsoft-mcp"
    ]
    assert wrapped_server["transport"] == "stdio"
