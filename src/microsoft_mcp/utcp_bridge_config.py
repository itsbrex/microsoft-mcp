from __future__ import annotations

import argparse
import copy
import json
import os
import re
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Sequence

DEFAULT_BRIDGE_NAME = "code-mode-mcp"


def _resolve_default_bridge_command() -> str:
    """Resolve the npx command path to use for the UTCP bridge.

    Resolution order:
    1. MICROSOFT_MCP_UTCP_BRIDGE_COMMAND environment variable.
    2. shutil.which("npx") — looks up npx on the current PATH.
    3. The literal string "npx" — relies on PATH at exec time.
    """
    override = os.getenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND")
    if override:
        return override
    discovered = shutil.which("npx")
    if discovered:
        return discovered
    return "npx"


DEFAULT_BRIDGE_COMMAND = _resolve_default_bridge_command()
DEFAULT_BRIDGE_ARGS = ["@utcp/code-mode-mcp"]


@dataclass(frozen=True, slots=True)
class ManualMapping:
    source_server_name: str
    manual_name: str


def load_json(path: Path) -> dict[str, Any]:
    data = json.loads(path.read_text())
    if not isinstance(data, dict):
        raise ValueError(f"Expected top-level JSON object in {path}")
    return data


def write_json(path: Path, payload: dict[str, Any] | list[dict[str, Any]]) -> None:
    path.write_text(json.dumps(payload, indent=2) + "\n")


def list_server_names(claude_config: dict[str, Any]) -> list[str]:
    mcp_servers = claude_config.get("mcpServers")
    if not isinstance(mcp_servers, dict) or not mcp_servers:
        raise ValueError(
            "Expected a non-empty 'mcpServers' object in the Claude config"
        )
    return list(mcp_servers)


def sanitize_manual_name(name: str) -> str:
    sanitized = re.sub(r"[^0-9a-zA-Z_]+", "_", name).strip("_")
    if not sanitized:
        sanitized = "manual"
    if sanitized[0].isdigit():
        sanitized = f"manual_{sanitized}"
    return sanitized


def derive_manual_name(source_server_name: str) -> str:
    """
    Build a stable manual name from a source server name.

    Examples:
    - google_sheets      -> google_sheets
    - microsoft-mcp      -> microsoft_mcp
    - code-mode-mcp      -> code_mode_mcp
    - browser-tools-mcp  -> browser_tools_mcp
    """
    return sanitize_manual_name(source_server_name)


def _infer_transport(server_config: dict[str, Any]) -> str | None:
    transport = server_config.get("transport")
    if isinstance(transport, str) and transport.strip():
        return transport.strip().lower()

    if isinstance(server_config.get("url"), str):
        return "http"
    if isinstance(server_config.get("command"), str):
        return "stdio"
    return None


def _select_server_names(
    claude_config: dict[str, Any],
    *,
    include_servers: list[str] | None = None,
    exclude_servers: list[str] | None = None,
    bridge_servers_to_skip: Sequence[str] | None = None,
) -> list[str]:
    available_names = list_server_names(claude_config)
    available_lookup = set(available_names)
    skip_set = {
        name
        for name in (bridge_servers_to_skip or ())
        if isinstance(name, str) and name
    }

    selected_names = available_names
    if include_servers:
        missing = [name for name in include_servers if name not in available_lookup]
        if missing:
            raise ValueError(
                f"Included server names were not found in the Claude config: {missing}"
            )
        include_lookup = set(include_servers)
        selected_names = [name for name in available_names if name in include_lookup]

    if exclude_servers:
        missing = [name for name in exclude_servers if name not in available_lookup]
        if missing:
            raise ValueError(
                f"Excluded server names were not found in the Claude config: {missing}"
            )
        exclude_lookup = set(exclude_servers)
        selected_names = [name for name in selected_names if name not in exclude_lookup]

    if include_servers is None and skip_set:
        selected_names = [name for name in selected_names if name not in skip_set]

    if not selected_names:
        raise ValueError("No MCP servers remain after include/exclude filtering")

    return selected_names


def build_manual_name_map(
    claude_config: dict[str, Any],
    *,
    include_servers: list[str] | None = None,
    exclude_servers: list[str] | None = None,
    bridge_servers_to_skip: Sequence[str] | None = None,
) -> list[ManualMapping]:
    selected_names = _select_server_names(
        claude_config,
        include_servers=include_servers,
        exclude_servers=exclude_servers,
        bridge_servers_to_skip=bridge_servers_to_skip,
    )

    mappings: list[ManualMapping] = []
    seen: set[str] = set()

    for source_server_name in selected_names:
        base_name = derive_manual_name(source_server_name)
        manual_name = base_name
        suffix = 2
        while manual_name in seen:
            manual_name = f"{base_name}_{suffix}"
            suffix += 1
        seen.add(manual_name)
        mappings.append(
            ManualMapping(
                source_server_name=source_server_name,
                manual_name=manual_name,
            )
        )

    return mappings


def build_utcp_config(
    claude_config: dict[str, Any],
    *,
    include_servers: list[str] | None = None,
    exclude_servers: list[str] | None = None,
    env_file_path: str | None = None,
    server_env_overrides: dict[str, dict[str, str]] | None = None,
    bridge_servers_to_skip: Sequence[str] | None = (DEFAULT_BRIDGE_NAME,),
) -> tuple[dict[str, Any], list[ManualMapping]]:
    mcp_servers = claude_config.get("mcpServers")
    if not isinstance(mcp_servers, dict) or not mcp_servers:
        raise ValueError(
            "Expected a non-empty 'mcpServers' object in the Claude config"
        )

    mappings = build_manual_name_map(
        claude_config,
        include_servers=include_servers,
        exclude_servers=exclude_servers,
        bridge_servers_to_skip=bridge_servers_to_skip,
    )

    load_variables_from: list[dict[str, Any]] = []
    if env_file_path:
        load_variables_from.append(
            {
                "variable_loader_type": "dotenv",
                "env_file_path": env_file_path,
            }
        )

    manual_call_templates: list[dict[str, Any]] = []
    overrides = server_env_overrides or {}

    for mapping in mappings:
        server_config = copy.deepcopy(mcp_servers[mapping.source_server_name])
        env_override = overrides.get(mapping.source_server_name)
        if env_override:
            env = server_config.setdefault("env", {})
            if not isinstance(env, dict):
                raise ValueError(
                    f"Expected 'env' for server '{mapping.source_server_name}' to be an object"
                )
            env.update(env_override)

        inferred_transport = _infer_transport(server_config)
        if inferred_transport and "transport" not in server_config:
            server_config["transport"] = inferred_transport

        manual_call_templates.append(
            {
                "name": mapping.manual_name,
                "call_template_type": "mcp",
                "config": {
                    "mcpServers": {
                        mapping.source_server_name: server_config,
                    }
                },
            }
        )

    utcp_config = {
        "load_variables_from": load_variables_from,
        "tool_repository": {"tool_repository_type": "in_memory"},
        "tool_search_strategy": {
            "tool_search_strategy_type": "tag_and_description_word_match"
        },
        "manual_call_templates": manual_call_templates,
        "post_processing": [],
    }
    return utcp_config, mappings


def build_bridge_claude_config(
    utcp_config_path: Path,
    *,
    bridge_name: str = DEFAULT_BRIDGE_NAME,
    bridge_command: str = DEFAULT_BRIDGE_COMMAND,
    bridge_args: list[str] | None = None,
) -> dict[str, Any]:
    args = bridge_args or DEFAULT_BRIDGE_ARGS
    return {
        "mcpServers": {
            bridge_name: {
                "command": bridge_command,
                "args": args,
                "env": {
                    "UTCP_CONFIG_FILE": str(utcp_config_path.resolve()),
                },
            }
        }
    }


def convert_claude_config(
    source_path: Path,
    *,
    output_dir: Path,
    include_servers: list[str] | None = None,
    exclude_servers: list[str] | None = None,
    env_file_path: str | None = None,
    server_env_overrides: dict[str, dict[str, str]] | None = None,
    utcp_config_name: str = ".utcp_config.json",
    bridge_config_name: str = "claude_desktop_config.utcp.json",
    manual_map_name: str = "manual_map.json",
    bridge_name: str = DEFAULT_BRIDGE_NAME,
    bridge_command: str = DEFAULT_BRIDGE_COMMAND,
    bridge_args: list[str] | None = None,
) -> dict[str, Path]:
    claude_config = load_json(source_path)
    utcp_config, mappings = build_utcp_config(
        claude_config,
        include_servers=include_servers,
        exclude_servers=exclude_servers,
        env_file_path=env_file_path,
        server_env_overrides=server_env_overrides,
        bridge_servers_to_skip=(bridge_name,),
    )

    output_dir.mkdir(parents=True, exist_ok=True)

    utcp_config_path = output_dir / utcp_config_name
    bridge_config_path = output_dir / bridge_config_name
    manual_map_path = output_dir / manual_map_name

    write_json(utcp_config_path, utcp_config)
    write_json(
        bridge_config_path,
        build_bridge_claude_config(
            utcp_config_path,
            bridge_name=bridge_name,
            bridge_command=bridge_command,
            bridge_args=bridge_args,
        ),
    )
    write_json(
        manual_map_path,
        [
            {
                "source_server_name": mapping.source_server_name,
                "manual_name": mapping.manual_name,
            }
            for mapping in mappings
        ],
    )

    return {
        "utcp_config": utcp_config_path,
        "bridge_config": bridge_config_path,
        "manual_map": manual_map_path,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Wrap a Claude Desktop mcpServers config into a UTCP "
            "code-mode bridge configuration."
        )
    )
    parser.add_argument(
        "source",
        type=Path,
        help="Path to the source Claude Desktop config JSON file.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        help="Directory where generated UTCP and bridge configs should be written.",
    )
    parser.add_argument(
        "--env-file",
        default=None,
        help="Optional dotenv path to add to load_variables_from in the UTCP config.",
    )
    parser.add_argument(
        "--bridge-name",
        default=DEFAULT_BRIDGE_NAME,
        help="Name of the single bridge server entry in the generated Claude config.",
    )
    parser.add_argument(
        "--bridge-command",
        default=DEFAULT_BRIDGE_COMMAND,
        help="Command used to launch the UTCP code-mode bridge.",
    )
    parser.add_argument(
        "--bridge-arg",
        action="append",
        dest="bridge_args",
        default=None,
        help=(
            "Argument passed to the bridge command. Repeat for multiple args. "
            f"Defaults to {DEFAULT_BRIDGE_ARGS[0]}."
        ),
    )
    parser.add_argument(
        "--include-server",
        action="append",
        dest="include_servers",
        default=None,
        help="Only wrap the named server. Repeat to include multiple servers.",
    )
    parser.add_argument(
        "--exclude-server",
        action="append",
        dest="exclude_servers",
        default=None,
        help="Skip wrapping the named server. Repeat to exclude multiple servers.",
    )
    parser.add_argument(
        "--set-env",
        nargs=3,
        metavar=("SERVER", "KEY", "VALUE"),
        action="append",
        dest="env_overrides",
        default=None,
        help=(
            "Override or add an env var in the wrapped server config. "
            "Repeat for multiple overrides."
        ),
    )
    parser.add_argument(
        "--list-servers",
        action="store_true",
        help="Print server names from the source config and exit.",
    )
    return parser


def _parse_env_overrides(
    entries: list[list[str]] | None,
) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    for server, key, value in entries or []:
        result.setdefault(server, {})[key] = value
    return result


def main(argv: list[str] | None = None) -> int:
    parser = build_parser()
    args = parser.parse_args(argv)

    source_config = load_json(args.source)

    if args.list_servers:
        for name in list_server_names(source_config):
            print(name)
        return 0

    if args.output_dir is None:
        parser.error("--output-dir is required unless --list-servers is used")

    outputs = convert_claude_config(
        args.source,
        output_dir=args.output_dir,
        include_servers=args.include_servers,
        exclude_servers=args.exclude_servers,
        env_file_path=args.env_file,
        server_env_overrides=_parse_env_overrides(args.env_overrides),
        bridge_name=args.bridge_name,
        bridge_command=args.bridge_command,
        bridge_args=args.bridge_args,
    )

    print(f"Generated UTCP config: {outputs['utcp_config']}")
    print(f"Generated bridge Claude config: {outputs['bridge_config']}")
    print(f"Generated manual map: {outputs['manual_map']}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
