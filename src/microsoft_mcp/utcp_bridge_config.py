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

DEFAULT_BRIDGE_NAME = "code-mode"
LEGACY_BRIDGE_NAME = "code-mode-mcp"
BRIDGE_SERVER_NAMES = (DEFAULT_BRIDGE_NAME, LEGACY_BRIDGE_NAME)
UPSTREAM_BRIDGE_PACKAGE = "@utcp/code-mode-mcp@1.2.1"

_CANONICAL_CONFIG_ENV = "UTCP_CONFIG_FILE"
_LEGACY_CONFIG_ENV = "UTCP_CONFIG_PATH"
_BRIDGE_EXTENSION_CONFIG_KEYS = frozenset(
    {"exclude_tools", "include_tools", "default_disabled"}
)

# Relative path from a code-mode repo root to the compiled MCP entry point.
_CODE_MODE_DIST_REL = Path("code-mode-mcp") / "dist" / "index.js"


def _local_code_mode_dist() -> Path | None:
    """Return the resolved dist/index.js path when MICROSOFT_MCP_CODE_MODE_DIR is set.

    Returns None if the env var is unset or the file does not exist.
    """
    local_dir = os.getenv("MICROSOFT_MCP_CODE_MODE_DIR")
    if not local_dir:
        return None
    candidate = Path(local_dir) / _CODE_MODE_DIST_REL
    return candidate if candidate.exists() else None


def _resolve_default_bridge_command() -> str:
    """Resolve the command used to launch the UTCP code-mode bridge.

    Resolution order:
    1. MICROSOFT_MCP_UTCP_BRIDGE_COMMAND environment variable.
    2. MICROSOFT_MCP_CODE_MODE_DIR set and dist/index.js present → node (via shutil.which).
    3. shutil.which("npx") — looks up npx on the current PATH.
    4. The literal string "npx" — relies on PATH at exec time.
    """
    override = os.getenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND")
    if override:
        return override
    if _local_code_mode_dist() is not None:
        return shutil.which("node") or "node"
    discovered = shutil.which("npx")
    if discovered:
        return discovered
    return "npx"


def _is_npx_command(command: str) -> bool:
    return Path(command).name.lower() in {"npx", "npx.cmd"}


def _resolve_bridge_args(
    command: str,
    *,
    explicit_command: bool,
    bridge_args: Sequence[str] | None = None,
) -> list[str]:
    """Resolve args for an explicit or automatically selected bridge command.

    When MICROSOFT_MCP_CODE_MODE_DIR points to a local code-mode checkout
    whose dist/index.js exists, the args contain that path so the bridge
    is launched without a network round-trip to npm. An explicit bridge
    command is treated as a self-contained launcher. Otherwise fall back
    to the exact upstream package release matching the canonical seven tools.
    """
    if bridge_args is not None:
        return list(bridge_args)
    if explicit_command:
        return [UPSTREAM_BRIDGE_PACKAGE] if _is_npx_command(command) else []
    dist = _local_code_mode_dist()
    if dist is not None:
        return [str(dist)]
    return [UPSTREAM_BRIDGE_PACKAGE]


def _resolve_default_bridge_args() -> list[str]:
    command_override = os.getenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND")
    command = command_override or _resolve_default_bridge_command()
    return _resolve_bridge_args(
        command,
        explicit_command=command_override is not None,
    )


DEFAULT_BRIDGE_COMMAND = _resolve_default_bridge_command()
DEFAULT_BRIDGE_ARGS = _resolve_default_bridge_args()


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


def _normalized_config_path(value: Any, *, bridge_name: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(
            f"Expected UTCP config path for bridge '{bridge_name}' to be "
            "a non-empty string"
        )
    return str(Path(value).expanduser().resolve())


def _normalize_bridge_registration(
    bridge_name: str,
    server_config: Any,
) -> dict[str, Any]:
    if not isinstance(server_config, dict):
        raise ValueError(
            f"Expected bridge registration '{bridge_name}' to be an object"
        )

    normalized = copy.deepcopy(server_config)
    env = normalized.get("env")
    if env is None:
        return normalized
    if not isinstance(env, dict):
        raise ValueError(f"Expected 'env' for bridge '{bridge_name}' to be an object")

    canonical_value = env.get(_CANONICAL_CONFIG_ENV)
    legacy_value = env.get(_LEGACY_CONFIG_ENV)
    canonical_path = (
        _normalized_config_path(canonical_value, bridge_name=bridge_name)
        if canonical_value is not None
        else None
    )
    legacy_path = (
        _normalized_config_path(legacy_value, bridge_name=bridge_name)
        if legacy_value is not None
        else None
    )
    if (
        canonical_path is not None
        and legacy_path is not None
        and canonical_path != legacy_path
    ):
        raise ValueError(
            f"Conflicting UTCP config paths for bridge '{bridge_name}': "
            f"{_CANONICAL_CONFIG_ENV} and {_LEGACY_CONFIG_ENV} differ"
        )

    selected_path = canonical_path or legacy_path
    if selected_path is not None:
        env[_CANONICAL_CONFIG_ENV] = selected_path
    env.pop(_LEGACY_CONFIG_ENV, None)
    return normalized


def _validate_bridge_registrations(claude_config: dict[str, Any]) -> None:
    mcp_servers = claude_config.get("mcpServers")
    if not isinstance(mcp_servers, dict):
        return

    normalized: dict[str, dict[str, Any]] = {}
    for bridge_name in BRIDGE_SERVER_NAMES:
        if bridge_name in mcp_servers:
            normalized[bridge_name] = _normalize_bridge_registration(
                bridge_name,
                mcp_servers[bridge_name],
            )

    if len(normalized) == 2:
        canonical = normalized[DEFAULT_BRIDGE_NAME]
        legacy = normalized[LEGACY_BRIDGE_NAME]
        if canonical != legacy:
            raise ValueError(
                "Divergent bridge registrations found for "
                f"'{DEFAULT_BRIDGE_NAME}' and '{LEGACY_BRIDGE_NAME}'"
            )


def _canonical_bridge_name(bridge_name: str) -> str:
    if bridge_name not in BRIDGE_SERVER_NAMES:
        raise ValueError(
            f"Unsupported bridge name '{bridge_name}'; expected "
            f"'{DEFAULT_BRIDGE_NAME}' or legacy '{LEGACY_BRIDGE_NAME}'"
        )
    return DEFAULT_BRIDGE_NAME


def _utcp_config_requires_bridge_extensions(utcp_config_path: Path) -> bool:
    if not utcp_config_path.exists():
        return False

    config = load_json(utcp_config_path)
    templates = config.get("manual_call_templates")
    if not isinstance(templates, list):
        return False
    return any(
        isinstance(template, dict)
        and bool(_BRIDGE_EXTENSION_CONFIG_KEYS.intersection(template))
        for template in templates
    )


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
    _validate_bridge_registrations(claude_config)
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

    if skip_set:
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
    bridge_servers_to_skip: Sequence[str] | None = BRIDGE_SERVER_NAMES,
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
    bridge_command: str | None = None,
    bridge_args: list[str] | None = None,
) -> dict[str, Any]:
    canonical_bridge_name = _canonical_bridge_name(bridge_name)
    environment_command = os.getenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND")
    explicit_command = bridge_command or environment_command
    command = explicit_command or _resolve_default_bridge_command()
    args = _resolve_bridge_args(
        command,
        explicit_command=explicit_command is not None,
        bridge_args=bridge_args,
    )

    uses_upstream_fallback = _is_npx_command(command) and args == [
        UPSTREAM_BRIDGE_PACKAGE
    ]
    uses_automatic_local_dist = (
        explicit_command is None
        and bridge_args is None
        and _local_code_mode_dist() is not None
    )
    uses_explicit_extension_bridge = (
        explicit_command is not None and not uses_upstream_fallback
    )
    if _utcp_config_requires_bridge_extensions(utcp_config_path) and not (
        uses_automatic_local_dist or uses_explicit_extension_bridge
    ):
        raise ValueError(
            "UTCP config requires Local Bridge Extensions; set "
            "MICROSOFT_MCP_CODE_MODE_DIR or "
            "an explicit bridge command (and arguments when needed)"
        )

    return {
        "mcpServers": {
            canonical_bridge_name: {
                "command": command,
                "args": args,
                "env": {
                    _CANONICAL_CONFIG_ENV: str(utcp_config_path.resolve()),
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
    bridge_command: str | None = None,
    bridge_args: list[str] | None = None,
) -> dict[str, Path]:
    claude_config = load_json(source_path)
    utcp_config, mappings = build_utcp_config(
        claude_config,
        include_servers=include_servers,
        exclude_servers=exclude_servers,
        env_file_path=env_file_path,
        server_env_overrides=server_env_overrides,
        bridge_servers_to_skip=BRIDGE_SERVER_NAMES,
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
        help=(
            "Bridge host name to accept. Both code-mode and legacy code-mode-mcp "
            "emit the canonical code-mode key."
        ),
    )
    parser.add_argument(
        "--bridge-command",
        default=None,
        help=(
            "Explicit command used to launch the UTCP code-mode bridge. "
            "Defaults to a local dist when configured, then the pinned upstream release."
        ),
    )
    parser.add_argument(
        "--bridge-arg",
        action="append",
        dest="bridge_args",
        default=None,
        help=(
            "Argument passed to the bridge command. Repeat for multiple args. "
            f"Automatic upstream fallback is {UPSTREAM_BRIDGE_PACKAGE}."
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
