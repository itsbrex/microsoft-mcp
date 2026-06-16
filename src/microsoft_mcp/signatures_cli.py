"""CLI for managing local plain-text Outlook signatures.

Exposed two ways:
- ``microsoft-mcp-signatures <cmd> ...`` (standalone console script)
- ``microsoft-mcp signatures <cmd> ...`` (subcommand of the main entry point)

Both routes share the ``main()`` function defined here.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path

from . import signatures
from .signatures import (
    InvalidSignatureNameError,
    NoAccountError,
    SignatureInfo,
    SignatureNotFoundError,
)


# ---------------------------------------------------------------------------
# Helpers


def _err(msg: str) -> None:
    print(f"error: {msg}", file=sys.stderr)


def _info_dict(info: SignatureInfo) -> dict:
    return {
        "account": info.account,
        "name": info.name,
        "path": str(info.path),
        "has_html": info.has_html,
        "size": info.size,
        "modified": dt.datetime.fromtimestamp(info.modified).isoformat(
            timespec="seconds"
        ),
    }


def _print_table(rows: list[SignatureInfo]) -> None:
    if not rows:
        print("(no signatures found)")
        return
    widths = {
        "account": max(len("ACCOUNT"), max(len(r.account) for r in rows)),
        "name": max(len("NAME"), max(len(r.name) for r in rows)),
        "size": max(len("BYTES"), max(len(str(r.size)) for r in rows)),
    }
    header = (
        f"{'ACCOUNT'.ljust(widths['account'])}  "
        f"{'NAME'.ljust(widths['name'])}  "
        f"{'BYTES'.rjust(widths['size'])}  HTML  MODIFIED"
    )
    print(header)
    print("-" * len(header))
    for r in rows:
        modified = dt.datetime.fromtimestamp(r.modified).strftime("%Y-%m-%d %H:%M")
        html_flag = "yes" if r.has_html else "no"
        print(
            f"{r.account.ljust(widths['account'])}  "
            f"{r.name.ljust(widths['name'])}  "
            f"{str(r.size).rjust(widths['size'])}  {html_flag:>4}  {modified}"
        )


def _read_content(args: argparse.Namespace) -> str:
    """Resolve signature content from --from-file / --stdin / --editor."""
    sources = [bool(args.from_file), bool(args.stdin), bool(args.editor)]
    if sum(sources) > 1:
        raise SystemExit("error: pass at most one of --from-file, --stdin, --editor")

    if args.from_file:
        return Path(args.from_file).expanduser().read_text(encoding="utf-8")

    if args.stdin:
        return sys.stdin.read()

    if args.editor:
        return _open_editor(initial="")

    if not sys.stdin.isatty():
        return sys.stdin.read()

    return _open_editor(initial="")


def _open_editor(initial: str) -> str:
    editor = os.environ.get("VISUAL") or os.environ.get("EDITOR") or "vi"
    with tempfile.NamedTemporaryFile(
        mode="w+",
        suffix=".txt",
        delete=False,
        encoding="utf-8",
    ) as tmp:
        tmp.write(initial)
        tmp_path = Path(tmp.name)
    try:
        rc = subprocess.call([editor, str(tmp_path)])
        if rc != 0:
            raise SystemExit(f"editor exited with status {rc}")
        return tmp_path.read_text(encoding="utf-8")
    finally:
        try:
            tmp_path.unlink()
        except OSError:
            pass


def _confirm(prompt: str) -> bool:
    if not sys.stdin.isatty():
        return False
    answer = input(f"{prompt} [y/N] ").strip().lower()
    return answer in {"y", "yes"}


# ---------------------------------------------------------------------------
# Subcommand handlers


def _cmd_dir(args: argparse.Namespace) -> int:
    del args
    print(str(signatures.resolve_dir()))
    return 0


def _cmd_path(args: argparse.Namespace) -> int:
    if not args.name:
        print(str(signatures.resolve_dir()))
        return 0
    path = signatures.signature_path(args.name, account=args.account, html=args.html)
    print(str(path))
    return 0


def _cmd_list(args: argparse.Namespace) -> int:
    rows = signatures.list_signatures(account=args.account)
    if args.json:
        print(json.dumps([_info_dict(r) for r in rows], indent=2))
    else:
        _print_table(rows)
    return 0


def _cmd_show(args: argparse.Namespace) -> int:
    content = signatures.read_signature(args.name, account=args.account, html=args.html)
    if content is None:
        suffix = ".html" if args.html else ".txt"
        slug = signatures.account_slug(override=args.account)
        _err(f"signature not found: {slug}/{args.name}{suffix}")
        return 1
    # No trailing newline forced; print as-is.
    sys.stdout.write(content)
    if not content.endswith("\n"):
        sys.stdout.write("\n")
    return 0


def _cmd_set(args: argparse.Namespace) -> int:
    content = _read_content(args)
    path = signatures.write_signature(
        args.name, content, account=args.account, html=args.html
    )
    print(f"wrote {path} ({len(content)} bytes)")
    return 0


def _cmd_edit(args: argparse.Namespace) -> int:
    existing = signatures.read_signature(
        args.name, account=args.account, html=args.html
    )
    updated = _open_editor(initial=existing or "")
    path = signatures.write_signature(
        args.name, updated, account=args.account, html=args.html
    )
    print(f"wrote {path} ({len(updated)} bytes)")
    return 0


def _cmd_rm(args: argparse.Namespace) -> int:
    path = signatures.signature_path(args.name, account=args.account, html=args.html)
    if not path.exists():
        _err(f"signature not found: {path}")
        return 1
    if not args.yes and not _confirm(f"delete {path}?"):
        print("aborted.")
        return 1
    signatures.delete_signature(args.name, account=args.account, html=args.html)
    print(f"deleted {path}")
    return 0


# ---------------------------------------------------------------------------
# Argparse wiring


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="microsoft-mcp-signatures",
        description=(
            "Manage local plain-text Outlook signatures used by the "
            "microsoft-mcp draft tools."
        ),
    )
    sub = parser.add_subparsers(dest="cmd", required=True, metavar="COMMAND")

    p_dir = sub.add_parser("dir", help="print the signatures directory")
    p_dir.set_defaults(func=_cmd_dir)

    p_path = sub.add_parser("path", help="print the resolved file path for a signature")
    p_path.add_argument("name", nargs="?", default=None)
    p_path.add_argument("--account")
    p_path.add_argument("--html", action="store_true")
    p_path.set_defaults(func=_cmd_path)

    p_list = sub.add_parser("list", help="list signatures")
    p_list.add_argument(
        "--account",
        help="account slug; use '*' or 'all' to list across all accounts",
    )
    p_list.add_argument("--json", action="store_true")
    p_list.set_defaults(func=_cmd_list)

    p_show = sub.add_parser("show", help="print a signature to stdout")
    p_show.add_argument("name")
    p_show.add_argument("--account")
    p_show.add_argument("--html", action="store_true")
    p_show.set_defaults(func=_cmd_show)

    p_set = sub.add_parser("set", help="create or replace a signature")
    p_set.add_argument("name")
    p_set.add_argument("--account")
    p_set.add_argument("--html", action="store_true")
    p_set.add_argument("--from-file")
    p_set.add_argument("--stdin", action="store_true", help="read content from stdin")
    p_set.add_argument("--editor", action="store_true", help="open $EDITOR for content")
    p_set.set_defaults(func=_cmd_set)

    p_edit = sub.add_parser("edit", help="open existing signature in $EDITOR")
    p_edit.add_argument("name")
    p_edit.add_argument("--account")
    p_edit.add_argument("--html", action="store_true")
    p_edit.set_defaults(func=_cmd_edit)

    p_rm = sub.add_parser("rm", help="delete a signature file")
    p_rm.add_argument("name")
    p_rm.add_argument("--account")
    p_rm.add_argument("--html", action="store_true")
    p_rm.add_argument("--yes", "-y", action="store_true", help="skip confirmation")
    p_rm.set_defaults(func=_cmd_rm)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except (
        InvalidSignatureNameError,
        NoAccountError,
        SignatureNotFoundError,
    ) as exc:
        _err(str(exc))
        return 1
    except SystemExit:
        raise
    except FileNotFoundError as exc:
        _err(str(exc))
        return 1


def cli_main() -> None:
    """Console-script entry point."""
    sys.exit(main(sys.argv[1:]))


if __name__ == "__main__":
    cli_main()
