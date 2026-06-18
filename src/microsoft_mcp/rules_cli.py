"""CLI for managing Outlook inbox rules (server-side message rules).

Exposed two ways (mirrors auth_cli.py / signatures_cli.py):
- ``microsoft-mcp-rules <cmd> ...``   (standalone console script)
- ``microsoft-mcp rules <cmd> ...``   (subcommand of the main entry point)

Both routes share ``main()``. Handlers import ``microsoft_mcp.tools`` lazily
so the heavy Graph/auth stack is not loaded for ``--help`` or dispatch.
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from typing import Any

# ---------------------------------------------------------------------------
# ANSI color (zero-dependency, TTY-gated, honors NO_COLOR)

_RESET = "\033[0m"
_COLORS = {
    "green": "\033[32m",
    "red": "\033[31m",
    "yellow": "\033[33m",
    "dim": "\033[2m",
}


def _color_enabled(stream: Any = None) -> bool:
    stream = stream if stream is not None else sys.stdout
    if os.environ.get("NO_COLOR") is not None:
        return False
    if os.environ.get("MICROSOFT_MCP_FORCE_COLOR") == "1":
        return True
    if os.environ.get("TERM", "") == "dumb":
        return False
    isatty = getattr(stream, "isatty", None)
    return bool(isatty()) if callable(isatty) else False


def _c(text: str, color: str) -> str:
    if not _color_enabled():
        return text
    return f"{_COLORS.get(color, '')}{text}{_RESET}"


# ---------------------------------------------------------------------------
# Helpers


def _print_json(payload: Any) -> None:
    print(json.dumps(payload, indent=2, default=str))


def _print_rules_table(rules: list[dict[str, Any]]) -> None:
    if not rules:
        print("(no inbox rules found)")
        return
    id_w = max(len("ID"), max(len(str(r.get("id", ""))) for r in rules))
    name_w = max(
        len("NAME"),
        max(len(str(r.get("display_name", r.get("displayName", "")))) for r in rules),
    )
    seq_w = max(len("SEQ"), max(len(str(r.get("sequence", ""))) for r in rules))
    header = (
        f"{'ID'.ljust(id_w)}  {'SEQ'.rjust(seq_w)}  "
        f"{'ENABLED'.ljust(7)}  {'NAME'.ljust(name_w)}"
    )
    print(header)
    print("-" * len(header))
    for r in rules:
        rid = str(r.get("id", ""))
        name = str(r.get("display_name", r.get("displayName", "")))
        seq = str(r.get("sequence", ""))
        enabled = r.get("is_enabled", r.get("isEnabled"))
        enabled_str = "yes" if enabled else "no"
        print(
            f"{rid.ljust(id_w)}  {seq.rjust(seq_w)}  "
            f"{enabled_str.ljust(7)}  {name.ljust(name_w)}"
        )


def _print_rule_detail(rule: dict[str, Any]) -> None:
    lines = [
        ("ID", rule.get("id", "")),
        ("Name", rule.get("display_name", rule.get("displayName", ""))),
        ("Sequence", rule.get("sequence", "")),
        ("Enabled", "yes" if rule.get("is_enabled", rule.get("isEnabled")) else "no"),
        ("Conditions", rule.get("conditions_summary", "")),
        ("Actions", rule.get("actions_summary", "")),
        ("Exceptions", rule.get("exceptions_summary", "")),
    ]
    for label, value in lines:
        if value:
            print(f"{label + ':':13} {value}")


# ---------------------------------------------------------------------------
# Subcommand handlers


def _cmd_list(args: argparse.Namespace) -> int:
    import microsoft_mcp.tools as tools

    rules = tools.list_inbox_rules.fn()
    if args.json:
        _print_json(rules)
    else:
        _print_rules_table(rules)
    return 0


def _cmd_get(args: argparse.Namespace) -> int:
    import microsoft_mcp.tools as tools

    rule = tools.get_inbox_rule.fn(args.id)
    if args.json:
        _print_json(rule)
    else:
        _print_rule_detail(rule)
    return 0


def _cmd_export(args: argparse.Namespace) -> int:
    import microsoft_mcp.tools as tools

    result = tools.export_inbox_rules.fn(path=args.output)
    if args.json:
        _print_json(result)
    else:
        if args.output:
            print(f"Exported {result.get('count', 0)} rule(s) to {args.output}")
        else:
            yaml_str = result.get("yaml", "")
            print(yaml_str, end="")
    return 0


def _cmd_import(args: argparse.Namespace) -> int:
    import microsoft_mcp.tools as tools

    result = tools.import_inbox_rules.fn(
        path=args.file,
        mode=args.mode,
        dry_run=args.dry_run,
    )
    if args.json:
        _print_json(result)
    else:
        imported = result.get("imported", 0)
        skipped = result.get("skipped", 0)
        errors = result.get("errors", [])
        dry = " (dry run)" if result.get("dry_run") else ""
        print(f"imported={imported} skipped={skipped} errors={len(errors)}{dry}")
        for err in errors:
            print(_c(f"  error: {err}", "red"), file=sys.stderr)
    return 0 if not result.get("errors") else 1


def _cmd_create(args: argparse.Namespace) -> int:
    import microsoft_mcp.tools as tools

    kwargs: dict[str, Any] = {"display_name": args.name}
    if args.from_contains:
        kwargs["from_addresses"] = args.from_contains
    if args.subject_contains:
        kwargs["subject_contains"] = args.subject_contains
    if args.move_to:
        kwargs["move_to_folder"] = args.move_to
    if args.mark_read:
        kwargs["mark_as_read"] = True
    if args.stop:
        kwargs["stop_processing_rules"] = True

    rule = tools.create_inbox_rule.fn(**kwargs)
    if args.json:
        _print_json(rule)
    else:
        rid = rule.get("id", "")
        name = rule.get("display_name", rule.get("displayName", ""))
        print(_c(f"  created rule '{name}' (id={rid})", "green"))
    return 0


def _cmd_delete(args: argparse.Namespace) -> int:
    if not args.confirm:
        print(
            "error: pass --confirm to delete a rule (this is destructive)",
            file=sys.stderr,
        )
        return 1

    import microsoft_mcp.tools as tools

    result = tools.delete_inbox_rule.fn(args.id)
    if args.json:
        _print_json(result)
    else:
        print(_c(f"  deleted rule {args.id}", "green"))
    return 0


def _cmd_toggle(args: argparse.Namespace) -> int:
    import microsoft_mcp.tools as tools

    result = tools.toggle_inbox_rule.fn(args.id)
    if args.json:
        _print_json(result)
    else:
        new_state = "enabled" if result.get("is_enabled") else "disabled"
        print(_c(f"  rule {args.id} is now {new_state}", "green"))
    return 0


# ---------------------------------------------------------------------------
# Argparse wiring


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="microsoft-mcp-rules",
        description="Manage Outlook inbox rules via microsoft-mcp.",
    )
    sub = parser.add_subparsers(dest="cmd", required=True, metavar="COMMAND")

    p_list = sub.add_parser("list", help="list all inbox rules")
    p_list.add_argument("--json", action="store_true")
    p_list.set_defaults(func=_cmd_list)

    p_get = sub.add_parser("get", help="get a single inbox rule by ID")
    p_get.add_argument("id", help="rule ID")
    p_get.add_argument("--json", action="store_true")
    p_get.set_defaults(func=_cmd_get)

    p_export = sub.add_parser("export", help="export inbox rules to YAML")
    p_export.add_argument(
        "--output", metavar="FILE", default=None, help="write YAML to FILE"
    )
    p_export.add_argument("--json", action="store_true")
    p_export.set_defaults(func=_cmd_export)

    p_import = sub.add_parser("import", help="import inbox rules from a YAML file")
    p_import.add_argument("file", help="YAML file to import")
    p_import.add_argument(
        "--mode",
        choices=["create", "sync"],
        default="create",
        help="import mode (default: create)",
    )
    p_import.add_argument(
        "--dry-run", action="store_true", help="validate without making changes"
    )
    p_import.add_argument("--json", action="store_true")
    p_import.set_defaults(func=_cmd_import)

    p_create = sub.add_parser("create", help="create a new inbox rule")
    p_create.add_argument("--name", required=True, help="display name for the rule")
    p_create.add_argument(
        "--from-contains",
        dest="from_contains",
        nargs="+",
        metavar="ADDR",
        help="match if sender contains these addresses",
    )
    p_create.add_argument(
        "--subject-contains",
        dest="subject_contains",
        nargs="+",
        metavar="STR",
        help="match if subject contains these strings",
    )
    p_create.add_argument("--move-to", dest="move_to", metavar="FOLDER")
    p_create.add_argument(
        "--mark-read",
        dest="mark_read",
        action="store_true",
        help="mark matching mail as read",
    )
    p_create.add_argument(
        "--stop", action="store_true", help="stop processing further rules"
    )
    p_create.add_argument("--json", action="store_true")
    p_create.set_defaults(func=_cmd_create)

    p_delete = sub.add_parser("delete", help="delete an inbox rule")
    p_delete.add_argument("id", help="rule ID")
    p_delete.add_argument(
        "--confirm",
        action="store_true",
        help="required: confirm the destructive delete",
    )
    p_delete.add_argument("--json", action="store_true")
    p_delete.set_defaults(func=_cmd_delete)

    p_toggle = sub.add_parser(
        "toggle", help="toggle a rule between enabled and disabled"
    )
    p_toggle.add_argument("id", help="rule ID")
    p_toggle.add_argument("--json", action="store_true")
    p_toggle.set_defaults(func=_cmd_toggle)

    return parser


def main(argv: list[str] | None = None) -> int:
    try:
        from dotenv import load_dotenv

        load_dotenv()
    except Exception:
        pass
    parser = _build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


def cli_main() -> None:
    """Console-script entry point."""
    sys.exit(main(sys.argv[1:]))


if __name__ == "__main__":
    cli_main()
