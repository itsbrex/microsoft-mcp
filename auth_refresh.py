#!/usr/bin/env python3
"""Backward-compat shim. Delegates to microsoft_mcp.auth_cli.

Historical interface (still supported):
    auth_refresh.py                   -> auth refresh
    auth_refresh.py <email>           -> auth refresh <email>
    auth_refresh.py --force <email>   -> auth refresh <email> --force
    auth_refresh.py --verify          -> auth verify
    auth_refresh.py --verify --live   -> auth verify --live
    auth_refresh.py ... --json        -> ... --json

Prefer `microsoft-mcp auth <cmd>` going forward.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from dotenv import load_dotenv  # noqa: E402

load_dotenv()


def _translate(argv: list[str]) -> list[str]:
    """Map the legacy flag interface onto auth_cli subcommands."""
    args = list(argv)
    want_json = "--json" in args
    args = [a for a in args if a != "--json"]

    if "--verify" in args:
        out = ["verify"]
        if "--live" in args:
            out.append("--live")
    else:
        email = next((a for a in args if not a.startswith("-")), None)
        out = ["refresh"]
        if email:
            out.append(email)
        if "--force" in args:
            out.append("--force")
    if want_json:
        out.append("--json")
    return out


def main() -> None:
    from microsoft_mcp import auth_cli

    sys.exit(auth_cli.main(_translate(sys.argv[1:])))


if __name__ == "__main__":
    main()
