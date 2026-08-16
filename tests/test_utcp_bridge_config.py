import json
import re
from pathlib import Path

import pytest

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


def sample_claude_config_with_bridge(
    bridge_name: str = "code-mode-mcp",
    *,
    config_env_name: str = "UTCP_CONFIG_PATH",
    config_path: str = "/tmp/selected.utcp.json",
):
    payload = sample_claude_config()
    payload["mcpServers"][bridge_name] = {
        "command": "npx",
        "args": ["@utcp/code-mode-mcp@1.2.1"],
        "env": {config_env_name: config_path},
    }
    return payload


def write_extension_utcp_config(
    tmp_path: Path,
    extension_key: str,
    extension_value,
) -> Path:
    utcp_config_path = tmp_path / ".utcp_config.json"
    utcp_config_path.write_text(
        json.dumps(
            {
                "manual_call_templates": [
                    {
                        "name": "microsoft_mcp",
                        "call_template_type": "mcp",
                        "config": {"mcpServers": {}},
                        extension_key: extension_value,
                    }
                ]
            }
        )
    )
    return utcp_config_path


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


@pytest.mark.parametrize("bridge_name", ["code-mode", "code-mode-mcp"])
def test_build_utcp_config_skips_canonical_and_legacy_bridge_inputs(bridge_name):
    _, mappings = build_utcp_config(
        sample_claude_config_with_bridge(bridge_name),
    )

    assert bridge_name not in {mapping.source_server_name for mapping in mappings}


@pytest.mark.parametrize("bridge_name", ["code-mode", "code-mode-mcp"])
def test_build_utcp_config_does_not_explicitly_import_bridge_server(bridge_name):
    with pytest.raises(ValueError, match="No MCP servers remain"):
        build_utcp_config(
            {
                "mcpServers": {
                    bridge_name: sample_claude_config_with_bridge(bridge_name)[
                        "mcpServers"
                    ][bridge_name]
                }
            },
            include_servers=[bridge_name],
        )


def test_build_utcp_config_accepts_equivalent_canonical_and_legacy_registrations():
    payload = sample_claude_config_with_bridge(
        "code-mode",
        config_env_name="UTCP_CONFIG_FILE",
    )
    payload["mcpServers"]["code-mode-mcp"] = {
        "command": "npx",
        "args": ["@utcp/code-mode-mcp@1.2.1"],
        "env": {"UTCP_CONFIG_PATH": "/tmp/selected.utcp.json"},
    }

    _, mappings = build_utcp_config(payload)

    assert {mapping.source_server_name for mapping in mappings} == {
        "microsoft-mcp",
        "github",
        "notion-mcp",
    }


def test_build_utcp_config_rejects_divergent_duplicate_bridge_registrations():
    payload = sample_claude_config_with_bridge(
        "code-mode",
        config_env_name="UTCP_CONFIG_FILE",
    )
    payload["mcpServers"]["code-mode-mcp"] = {
        "command": "node",
        "args": ["/different/local/bridge.js"],
        "env": {"UTCP_CONFIG_PATH": "/tmp/selected.utcp.json"},
    }

    with pytest.raises(ValueError, match="Divergent bridge registrations"):
        build_utcp_config(payload)


def test_build_utcp_config_accepts_equal_config_path_aliases():
    payload = sample_claude_config_with_bridge(
        "code-mode",
        config_env_name="UTCP_CONFIG_FILE",
    )
    payload["mcpServers"]["code-mode"]["env"][
        "UTCP_CONFIG_PATH"
    ] = "/tmp/selected.utcp.json"

    _, mappings = build_utcp_config(payload)

    assert len(mappings) == 3


def test_build_utcp_config_rejects_conflicting_config_path_aliases():
    payload = sample_claude_config_with_bridge(
        "code-mode",
        config_env_name="UTCP_CONFIG_FILE",
    )
    payload["mcpServers"]["code-mode"]["env"][
        "UTCP_CONFIG_PATH"
    ] = "/tmp/different.utcp.json"

    with pytest.raises(ValueError, match="Conflicting UTCP config paths"):
        build_utcp_config(payload)


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

    from microsoft_mcp.utcp_bridge_config import DEFAULT_BRIDGE_COMMAND

    assert list(bridge_config["mcpServers"]) == ["code-mode"]
    assert bridge_config["mcpServers"]["code-mode"]["command"] == DEFAULT_BRIDGE_COMMAND
    assert bridge_config["mcpServers"]["code-mode"]["env"]["UTCP_CONFIG_FILE"] == str(
        utcp_config_path.resolve()
    )
    assert "UTCP_CONFIG_PATH" not in bridge_config["mcpServers"]["code-mode"]["env"]


def test_build_bridge_config_rejects_non_bridge_host_name(tmp_path: Path):
    with pytest.raises(ValueError, match="Unsupported bridge name"):
        build_bridge_claude_config(
            tmp_path / ".utcp_config.json",
            bridge_name="custom-code-mode",
        )


def test_build_bridge_config_pins_upstream_fallback(monkeypatch, tmp_path: Path):
    monkeypatch.delenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_CODE_MODE_DIR", raising=False)
    monkeypatch.setattr(
        "microsoft_mcp.utcp_bridge_config.shutil.which",
        lambda command: f"/usr/bin/{command}",
    )

    bridge_config = build_bridge_claude_config(tmp_path / ".utcp_config.json")
    bridge = bridge_config["mcpServers"]["code-mode"]

    assert bridge["command"] == "/usr/bin/npx"
    assert bridge["args"] == ["@utcp/code-mode-mcp@1.2.1"]


def test_build_bridge_config_prefers_local_dist(monkeypatch, tmp_path: Path):
    local_root = tmp_path / "code-mode"
    local_dist = local_root / "code-mode-mcp" / "dist" / "index.js"
    local_dist.parent.mkdir(parents=True)
    local_dist.write_text("// built bridge\n")
    monkeypatch.delenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", raising=False)
    monkeypatch.setenv("MICROSOFT_MCP_CODE_MODE_DIR", str(local_root))
    monkeypatch.setattr(
        "microsoft_mcp.utcp_bridge_config.shutil.which",
        lambda command: f"/usr/bin/{command}",
    )

    bridge_config = build_bridge_claude_config(tmp_path / ".utcp_config.json")
    bridge = bridge_config["mcpServers"]["code-mode"]

    assert bridge["command"] == "/usr/bin/node"
    assert bridge["args"] == [str(local_dist)]


def test_extension_config_rejects_upstream_fallback(monkeypatch, tmp_path: Path):
    utcp_config_path = write_extension_utcp_config(
        tmp_path,
        "exclude_tools",
        ["delete_email"],
    )
    monkeypatch.delenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_CODE_MODE_DIR", raising=False)

    with pytest.raises(ValueError, match="Bridge Extensions"):
        build_bridge_claude_config(utcp_config_path)


def test_extension_config_accepts_explicit_bridge_command(monkeypatch, tmp_path: Path):
    utcp_config_path = write_extension_utcp_config(tmp_path, "default_disabled", True)
    monkeypatch.delenv("MICROSOFT_MCP_CODE_MODE_DIR", raising=False)
    monkeypatch.setenv(
        "MICROSOFT_MCP_UTCP_BRIDGE_COMMAND",
        "/opt/local/bin/code-mode-mcp",
    )

    bridge_config = build_bridge_claude_config(utcp_config_path)
    bridge = bridge_config["mcpServers"]["code-mode"]

    assert bridge["command"] == "/opt/local/bin/code-mode-mcp"
    assert bridge["args"] == []


def test_extension_config_accepts_explicit_bridge_command_argument(
    monkeypatch, tmp_path: Path
):
    utcp_config_path = write_extension_utcp_config(tmp_path, "default_disabled", True)
    monkeypatch.delenv("MICROSOFT_MCP_CODE_MODE_DIR", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", raising=False)

    bridge_config = build_bridge_claude_config(
        utcp_config_path,
        bridge_command="/opt/local/bin/code-mode-mcp",
    )
    bridge = bridge_config["mcpServers"]["code-mode"]

    assert bridge["command"] == "/opt/local/bin/code-mode-mcp"
    assert bridge["args"] == []


def test_extension_config_rejects_explicit_npx_upstream_fallback(
    monkeypatch, tmp_path: Path
):
    utcp_config_path = write_extension_utcp_config(
        tmp_path,
        "exclude_tools",
        ["delete_email"],
    )
    monkeypatch.delenv("MICROSOFT_MCP_CODE_MODE_DIR", raising=False)
    monkeypatch.setenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", "/usr/local/bin/npx")

    with pytest.raises(ValueError, match="Bridge Extensions"):
        build_bridge_claude_config(utcp_config_path)


def test_extension_config_accepts_local_dist(monkeypatch, tmp_path: Path):
    utcp_config_path = write_extension_utcp_config(
        tmp_path,
        "include_tools",
        ["list_emails"],
    )
    local_root = tmp_path / "code-mode"
    local_dist = local_root / "code-mode-mcp" / "dist" / "index.js"
    local_dist.parent.mkdir(parents=True)
    local_dist.write_text("// built bridge\n")
    monkeypatch.delenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", raising=False)
    monkeypatch.setenv("MICROSOFT_MCP_CODE_MODE_DIR", str(local_root))
    monkeypatch.setattr(
        "microsoft_mcp.utcp_bridge_config.shutil.which",
        lambda command: f"/usr/bin/{command}",
    )

    bridge_config = build_bridge_claude_config(utcp_config_path)
    bridge = bridge_config["mcpServers"]["code-mode"]

    assert bridge["command"] == "/usr/bin/node"
    assert bridge["args"] == [str(local_dist)]


def test_extension_config_rejects_args_without_explicit_command(
    monkeypatch, tmp_path: Path
):
    utcp_config_path = write_extension_utcp_config(
        tmp_path,
        "exclude_tools",
        ["delete_email"],
    )
    monkeypatch.delenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_CODE_MODE_DIR", raising=False)
    monkeypatch.setattr(
        "microsoft_mcp.utcp_bridge_config.shutil.which",
        lambda command: f"/usr/bin/{command}",
    )

    with pytest.raises(ValueError, match="Bridge Extensions"):
        build_bridge_claude_config(
            utcp_config_path,
            bridge_args=["@itsbrex/code-mode-mcp"],
        )


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

    bridge_payload = json.loads(outputs["bridge_config"].read_text())
    assert list(bridge_payload["mcpServers"]) == ["code-mode"]
    bridge_env = bridge_payload["mcpServers"]["code-mode"]["env"]
    assert bridge_env == {"UTCP_CONFIG_FILE": str(outputs["utcp_config"].resolve())}


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


def test_default_bridge_command_does_not_hardcode_user_path(monkeypatch):
    """The default bridge command must not be a hardcoded user-specific path.

    Simulate a clean environment (no override, npx not on PATH) so the
    constant cannot accidentally pick up a developer-specific path.
    """
    monkeypatch.delenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", raising=False)
    monkeypatch.setenv("PATH", "")
    import importlib
    import microsoft_mcp.utcp_bridge_config as mod

    importlib.reload(mod)
    try:
        # Match any hardcoded user home (/Users/<name>/ or /home/<name>/) so
        # the guard catches leaks from any contributor's machine.
        hardcoded_home = re.search(
            r"/(?:Users|home)/[A-Za-z0-9_.-]+/", mod.DEFAULT_BRIDGE_COMMAND
        )
        assert hardcoded_home is None, (
            f"DEFAULT_BRIDGE_COMMAND contains a hardcoded user path "
            f"({hardcoded_home.group(0)!r}); use shutil.which"
        )
        assert mod.DEFAULT_BRIDGE_COMMAND, "DEFAULT_BRIDGE_COMMAND must not be empty"
    finally:
        importlib.reload(mod)


def test_bridge_command_can_be_overridden_by_env(monkeypatch):
    """An env override should win over auto-discovery."""
    monkeypatch.setenv(
        "MICROSOFT_MCP_UTCP_BRIDGE_COMMAND",
        "/custom/path/to/code-mode-mcp",
    )
    # Re-import to re-evaluate the module-level constant.
    import importlib
    import microsoft_mcp.utcp_bridge_config as mod

    importlib.reload(mod)
    try:
        assert mod.DEFAULT_BRIDGE_COMMAND == "/custom/path/to/code-mode-mcp"
        assert mod.DEFAULT_BRIDGE_ARGS == []
    finally:
        # Restore original module state for other tests.
        monkeypatch.delenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", raising=False)
        importlib.reload(mod)


def test_npx_bridge_command_override_keeps_pinned_upstream_arg(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", "/custom/path/to/npx")
    import importlib
    from src.microsoft_mcp import utcp_bridge_config as mod

    importlib.reload(mod)
    try:
        assert mod.DEFAULT_BRIDGE_COMMAND == "/custom/path/to/npx"
        assert mod.DEFAULT_BRIDGE_ARGS == ["@utcp/code-mode-mcp@1.2.1"]
    finally:
        monkeypatch.delenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND", raising=False)
        importlib.reload(mod)
