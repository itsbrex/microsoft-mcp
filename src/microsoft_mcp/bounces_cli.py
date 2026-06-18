"""CLI for scanning bounce/NDR messages from a Microsoft 365 mailbox.

Exposed two ways (mirrors intel_cli.py / auth_cli.py / signatures_cli.py):
- ``microsoft-mcp-bounces <cmd> ...``   (standalone console script)
- ``microsoft-mcp bounces <cmd> ...``   (subcommand of the main entry point)

Both routes share ``main()``. The ``scan`` subcommand makes LIVE Graph calls;
``patterns`` is read-only and needs no Graph bootstrap.
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
    "cyan": "\033[36m",
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
# Folder resolver (lightweight — no tools.py import)

_WELL_KNOWN_FOLDERS = frozenset(
    {
        "inbox",
        "sentitems",
        "drafts",
        "deleteditems",
        "junkemail",
        "archive",
        "outbox",
        "clutter",
        "conversationhistory",
        "recoverableitemsdeletions",
        "scheduled",
        "searchfolders",
        "serverfailures",
        "syncissues",
        "msgfolderroot",
    }
)


def _resolve_folder(graph: Any, folder: str) -> str:
    """Resolve a folder name to a Graph folder id (or well-known name).

    Well-known names (case-insensitive) pass through canonicalised; a display
    name is looked up via /me/mailFolders; anything else (e.g. an opaque id) is
    returned unchanged so it is never corrupted.
    """
    raw = folder.strip()
    key = raw.casefold().replace(" ", "")
    if key in _WELL_KNOWN_FOLDERS:
        return key
    data = graph.request(
        "GET", "/me/mailFolders", params={"$top": 100, "$select": "id,displayName"}
    )
    for f in (data or {}).get("value", []):
        if (f.get("displayName") or "").casefold() == raw.casefold():
            return f["id"]
    return raw  # assume caller passed a folder id


# ---------------------------------------------------------------------------
# Bootstrap helpers


def _bootstrap_graph() -> Any:
    """Import and wire up the graph module with an auth instance.

    Mirrors the bootstrap that intel_cli.py performs.
    """
    import pathlib as pl

    from microsoft_mcp import graph
    from microsoft_mcp.auth_base import AuthProvider  # noqa: F401

    def _env_path(name: str) -> pl.Path | None:
        value = os.getenv(name)
        return pl.Path(value) if value else None

    auth_method = os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower()

    if auth_method == "msal":
        from microsoft_mcp.auth_msal import MSALRefreshTokenAuth

        auth = MSALRefreshTokenAuth(
            tokens_dir=_env_path("MICROSOFT_MCP_TOKENS_DIR"),
            client_id=os.getenv("MICROSOFT_MCP_CLIENT_ID"),
            tenant_id=os.getenv("MICROSOFT_MCP_TENANT_ID"),
            account_identifier=os.getenv("MICROSOFT_MCP_ACCOUNT_ID"),
        )
    else:
        from microsoft_mcp.auth import AzureAuthentication

        auth = AzureAuthentication(
            auth_record_file=_env_path("AZURE_CRED_CACHE_FILE"),
            token_cache_file=_env_path("AZURE_TOKEN_CACHE_FILE"),
        )

    graph.set_auth_instance(auth)
    return graph


# ---------------------------------------------------------------------------
# Subcommand handlers


def _cmd_scan(args: argparse.Namespace) -> int:
    from microsoft_mcp import bounces

    graph = _bootstrap_graph()
    folder_id = _resolve_folder(graph, args.folder)
    rows = bounces.scan_folder(graph.request, folder_id, limit=args.limit)

    if args.output:
        bounces.write_csv(rows, args.output)

    if args.json:
        print(json.dumps({"count": len(rows), "rows": rows}, indent=2, default=str))
    else:
        print(f"Bounces found: {len(rows)}")
        if rows:
            from collections import Counter

            reasons = Counter(r["reason"] for r in rows)
            print("\nReason breakdown:")
            for reason, count in reasons.most_common():
                print(f"  {_c(reason, 'yellow')}: {count}")
        if args.output:
            print(_c(f"\nCSV written to: {args.output}", "cyan"))
    return 0


def _cmd_patterns(args: argparse.Namespace) -> int:
    from microsoft_mcp import bounces

    catalogs: dict[str, Any] = {
        "SUBJECT_KEYWORDS": bounces.SUBJECT_KEYWORDS,
        "SENDER_PATTERNS": bounces.SENDER_PATTERNS,
        "BODY_PATTERNS": bounces.BODY_PATTERNS,
        "BOUNCE_REASONS": [
            {"pattern": p, "reason": r} for p, r in bounces.BOUNCE_REASONS
        ],
        "STRONG_SUBJECT_INDICATORS": list(bounces.STRONG_SUBJECT_INDICATORS),
        "EXCLUDED_SUBJECT_PREFIXES": list(bounces.EXCLUDED_SUBJECT_PREFIXES),
    }

    if args.json:
        print(json.dumps(catalogs, indent=2))
    else:
        print(f"SUBJECT_KEYWORDS ({len(bounces.SUBJECT_KEYWORDS)}):")
        for kw in bounces.SUBJECT_KEYWORDS:
            print(f"  {kw}")

        print(f"\nSENDER_PATTERNS ({len(bounces.SENDER_PATTERNS)}):")
        for sp in bounces.SENDER_PATTERNS:
            print(f"  {sp}")

        print(f"\nBODY_PATTERNS ({len(bounces.BODY_PATTERNS)}):")
        for bp in bounces.BODY_PATTERNS:
            print(f"  {bp}")

        print(f"\nBOUNCE_REASONS ({len(bounces.BOUNCE_REASONS)}):")
        for pat, reason in bounces.BOUNCE_REASONS:
            print(f"  {_c(pat, 'dim')} -> {reason}")

        print(
            f"\nSTRONG_SUBJECT_INDICATORS ({len(bounces.STRONG_SUBJECT_INDICATORS)}):"
        )
        for ind in bounces.STRONG_SUBJECT_INDICATORS:
            print(f"  {ind}")

        print(
            f"\nEXCLUDED_SUBJECT_PREFIXES ({len(bounces.EXCLUDED_SUBJECT_PREFIXES)}):"
        )
        for pfx in bounces.EXCLUDED_SUBJECT_PREFIXES:
            print(f"  {pfx}")
    return 0


# ---------------------------------------------------------------------------
# Argparse wiring


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="microsoft-mcp-bounces",
        description="Scan Outlook folders for bounce/NDR messages.",
    )
    sub = parser.add_subparsers(dest="cmd", required=True, metavar="COMMAND")

    p_scan = sub.add_parser("scan", help="scan a folder for bounce messages")
    p_scan.add_argument(
        "--folder",
        default="inbox",
        metavar="FOLDER",
        help="folder alias or ID to scan (default: inbox)",
    )
    p_scan.add_argument(
        "--limit",
        type=int,
        default=200,
        metavar="N",
        help="max messages to scan (default: 200)",
    )
    p_scan.add_argument(
        "--output",
        metavar="CSV_PATH",
        default=None,
        help="write bounce rows to this CSV file",
    )
    p_scan.add_argument("--json", action="store_true", help="emit JSON output")
    p_scan.set_defaults(func=_cmd_scan)

    p_patterns = sub.add_parser(
        "patterns", help="show bounce detection pattern catalogs (read-only)"
    )
    p_patterns.add_argument("--json", action="store_true", help="emit JSON output")
    p_patterns.set_defaults(func=_cmd_patterns)

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
