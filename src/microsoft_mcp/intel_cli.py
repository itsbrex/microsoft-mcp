"""CLI for generating intelligence reports from the intel engine.

Exposed two ways (mirrors auth_cli.py / signatures_cli.py):
- ``microsoft-mcp-intel <cmd> ...``   (standalone console script)
- ``microsoft-mcp intel <cmd> ...``   (subcommand of the main entry point)

Both routes share ``main()``. The CLI makes LIVE Graph calls, so it imports
``graph`` and the engine and calls ``engine.generate_*(graph.request, ...)``.
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
# Bootstrap helpers


def _resolve_account() -> str:
    """Return the active account id from env (fall back to 'default')."""
    return os.getenv("MICROSOFT_MCP_ACCOUNT_ID") or "default"


def _bootstrap_graph() -> Any:
    """Import and wire up the graph module with an auth instance.

    Mirrors the bootstrap that tools.py performs at import time.
    """
    from microsoft_mcp import graph
    from microsoft_mcp.auth_base import AuthProvider

    auth_method = os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower()

    auth: AuthProvider
    if auth_method == "msal":
        from microsoft_mcp.auth_msal import MSALRefreshTokenAuth
        import pathlib as pl

        def _env_path(name: str) -> pl.Path | None:
            value = os.getenv(name)
            return pl.Path(value) if value else None

        auth = MSALRefreshTokenAuth(
            tokens_dir=_env_path("MICROSOFT_MCP_TOKENS_DIR"),
            client_id=os.getenv("MICROSOFT_MCP_CLIENT_ID"),
            tenant_id=os.getenv("MICROSOFT_MCP_TENANT_ID"),
            account_identifier=os.getenv("MICROSOFT_MCP_ACCOUNT_ID"),
        )
    else:
        from microsoft_mcp.auth import AzureAuthentication
        import pathlib as pl

        def _env_path(name: str) -> pl.Path | None:  # type: ignore[no-redef]
            value = os.getenv(name)
            return pl.Path(value) if value else None

        auth = AzureAuthentication(
            auth_record_file=_env_path("AZURE_CRED_CACHE_FILE"),
            token_cache_file=_env_path("AZURE_TOKEN_CACHE_FILE"),
        )

    graph.set_auth_instance(auth)
    return graph


# ---------------------------------------------------------------------------
# Human-readable summary helpers


def _print_briefing(report: dict[str, Any]) -> None:
    print(f"Morning Briefing — {report.get('account', '?')}")
    print(f"Generated: {report.get('generated_at', '?')}")
    items = report.get("priority_items", [])
    if not items:
        print(_c("  No priority items.", "dim"))
        return
    print(f"\n{len(items)} priority item(s):")
    for item in items:
        score = item.get("score", 0)
        title = item.get("title", "?")
        desc = item.get("description", "")
        hint = item.get("action_hint", "")
        color = "red" if score >= 80 else "yellow" if score >= 50 else "dim"
        print(_c(f"  [{score:5.1f}] {title}", color))
        if desc:
            print(f"         {desc}")
        if hint:
            print(_c(f"         → {hint}", "cyan"))
    email_summary = report.get("email_summary", {})
    cal_summary = report.get("calendar_summary", {})
    print(
        f"\nEmail: {email_summary.get('unread_total', 0)} unread, "
        f"{len(email_summary.get('needs_response', []))} need response"
    )
    print(
        f"Calendar: {cal_summary.get('total_events_today', 0)} events today, "
        f"{cal_summary.get('meeting_hours_today', 0):.1f}h in meetings"
    )


def _print_signals(report: dict[str, Any]) -> None:
    print(f"Priority Signals — {report.get('account', '?')}")
    print(f"Generated: {report.get('generated_at', '?')}")
    total = report.get("total_signals", 0)
    print(f"Total signals: {total}")
    for bucket, color in (
        ("critical", "red"),
        ("important", "yellow"),
        ("informational", "dim"),
    ):
        items = report.get(bucket, [])
        if not items:
            continue
        print(f"\n{bucket.upper()} ({len(items)}):")
        for item in items:
            score = item.get("score", 0)
            title = item.get("title", "?")
            print(_c(f"  [{score:5.1f}] {title}", color))


def _print_contact(report: dict[str, Any]) -> None:
    print(f"Contact Intelligence — {report.get('target_name', '?')}")
    print(f"  Email: {report.get('target_email', '?')}")
    if report.get("company"):
        print(f"  Company: {report['company']}")
    if report.get("job_title"):
        print(f"  Title: {report['job_title']}")
    rel = report.get("relationship", {})
    print("\nRelationship:")
    print(f"  Engagement score: {rel.get('engagement_score', 0):.1f}")
    print(f"  Trend: {rel.get('trend', '?')}")
    print(f"  Days since contact: {rel.get('days_since_contact', '?')}")
    print(
        f"  Emails sent/received: {rel.get('sent_to', 0)}/{rel.get('received_from', 0)}"
    )
    threads = report.get("recent_threads", [])
    if threads:
        print(f"\nRecent threads ({len(threads)}):")
        for t in threads:
            print(f"  - {t.get('subject', '?')} ({t.get('direction', '?')})")
    pending = report.get("pending_items", [])
    if pending:
        print(_c(f"\nPending response ({len(pending)}):", "yellow"))
        for p in pending:
            print(f"  - {p.get('subject', '?')} ({p.get('age_hours', 0):.0f}h ago)")


def _print_recap(report: dict[str, Any]) -> None:
    print(f"End-of-Day Recap — {report.get('account', '?')}")
    print(f"Generated: {report.get('generated_at', '?')}")
    print("\nToday's activity:")
    print(f"  Emails received: {report.get('emails_received_today', 0)}")
    print(f"  Emails sent:     {report.get('emails_sent_today', 0)}")
    print(f"  Still unread:    {report.get('emails_still_unread', 0)}")
    print(f"  Meetings attended: {report.get('meetings_attended', 0)}")
    pending = report.get("threads_still_pending", [])
    if pending:
        print(_c(f"\nThreads still pending ({len(pending)}):", "yellow"))
        for t in pending:
            print(f"  - {t.get('subject', '?')}")
    tomorrow = report.get("tomorrow_preview", [])
    if tomorrow:
        print(f"\nTomorrow ({len(tomorrow)} events):")
        for ev in tomorrow:
            print(f"  - {ev.get('subject', '?')} @ {ev.get('start', '?')[:16]}")


# ---------------------------------------------------------------------------
# Subcommand handlers


def _cmd_briefing(args: argparse.Namespace) -> int:
    from datetime import datetime
    from zoneinfo import ZoneInfo

    from microsoft_mcp.intel import engine

    graph = _bootstrap_graph()
    account = _resolve_account()
    now = datetime.now(ZoneInfo(args.timezone))
    report = engine.generate_briefing(
        graph.request,
        account=account,
        timezone=args.timezone,
        now=now,
    )
    result: dict[str, Any] = dict(report)
    result["priority_items"] = result["priority_items"][: args.limit]
    if args.json:
        print(json.dumps(result, indent=2, default=str))
    else:
        _print_briefing(result)
    return 0


def _cmd_signals(args: argparse.Namespace) -> int:
    from datetime import datetime
    from zoneinfo import ZoneInfo

    from microsoft_mcp.intel import engine

    graph = _bootstrap_graph()
    account = _resolve_account()
    now = datetime.now(ZoneInfo(args.timezone))
    report = engine.generate_signals(
        graph.request,
        account=account,
        timezone=args.timezone,
        now=now,
    )
    result: dict[str, Any] = dict(report)
    if args.level != "all":
        for bucket in ("critical", "important", "informational"):
            if bucket != args.level:
                result[bucket] = []
    if args.json:
        print(json.dumps(result, indent=2, default=str))
    else:
        _print_signals(result)
    return 0


def _cmd_contact(args: argparse.Namespace) -> int:
    from datetime import datetime
    from zoneinfo import ZoneInfo

    from microsoft_mcp.intel import engine

    graph = _bootstrap_graph()
    account = _resolve_account()
    now = datetime.now(ZoneInfo("UTC"))
    report = engine.generate_contact_report(
        graph.request,
        account=account,
        target_email=args.email,
        now=now,
        lookback_days=args.days,
    )
    result = dict(report)
    if args.json:
        print(json.dumps(result, indent=2, default=str))
    else:
        _print_contact(result)
    return 0


def _cmd_recap(args: argparse.Namespace) -> int:
    from datetime import datetime
    from zoneinfo import ZoneInfo

    from microsoft_mcp.intel import engine

    graph = _bootstrap_graph()
    account = _resolve_account()
    now = datetime.now(ZoneInfo(args.timezone))
    report = engine.generate_recap(
        graph.request,
        account=account,
        timezone=args.timezone,
        now=now,
    )
    result = dict(report)
    if args.json:
        print(json.dumps(result, indent=2, default=str))
    else:
        _print_recap(result)
    return 0


# ---------------------------------------------------------------------------
# Argparse wiring


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="microsoft-mcp-intel",
        description="Generate intelligence reports from Microsoft 365 data.",
    )
    sub = parser.add_subparsers(dest="cmd", required=True, metavar="COMMAND")

    p_briefing = sub.add_parser("briefing", help="generate morning briefing")
    p_briefing.add_argument(
        "--timezone",
        default="UTC",
        metavar="TZ",
        help="IANA timezone (default: UTC)",
    )
    p_briefing.add_argument(
        "--limit",
        type=int,
        default=10,
        metavar="N",
        help="max priority items to return (default: 10)",
    )
    p_briefing.add_argument("--json", action="store_true", help="emit JSON output")
    p_briefing.set_defaults(func=_cmd_briefing)

    p_signals = sub.add_parser("signals", help="get priority signals by urgency level")
    p_signals.add_argument(
        "--timezone",
        default="UTC",
        metavar="TZ",
        help="IANA timezone (default: UTC)",
    )
    p_signals.add_argument(
        "--level",
        choices=["all", "critical", "important", "informational"],
        default="all",
        help="urgency bucket to return (default: all)",
    )
    p_signals.add_argument("--json", action="store_true", help="emit JSON output")
    p_signals.set_defaults(func=_cmd_signals)

    p_contact = sub.add_parser(
        "contact", help="generate intelligence report for a contact"
    )
    p_contact.add_argument("email", help="target contact email address")
    p_contact.add_argument(
        "--days",
        type=int,
        default=30,
        metavar="N",
        help="look-back window in days (default: 30)",
    )
    p_contact.add_argument("--json", action="store_true", help="emit JSON output")
    p_contact.set_defaults(func=_cmd_contact)

    p_recap = sub.add_parser("recap", help="generate end-of-day recap")
    p_recap.add_argument(
        "--timezone",
        default="UTC",
        metavar="TZ",
        help="IANA timezone (default: UTC)",
    )
    p_recap.add_argument("--json", action="store_true", help="emit JSON output")
    p_recap.set_defaults(func=_cmd_recap)

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
