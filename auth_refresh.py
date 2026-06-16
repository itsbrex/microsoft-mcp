#!/usr/bin/env python3
"""Refresh, re-authenticate, or verify MSAL tokens for Microsoft MCP.

Usage:
    auth_refresh.py                   # refresh all saved accounts
    auth_refresh.py <email>           # refresh a single account
    auth_refresh.py --force <email>   # clear saved tokens + re-run device-code flow
    auth_refresh.py --verify          # check token-to-filename integrity
    auth_refresh.py --verify --live   # also call Graph /me for each token

Only supports MSAL auth. The Azure SDK path manages its own token cache.
"""

import argparse
import json
import os
import sys
from pathlib import Path

# Add src to path so we can import our modules
sys.path.insert(0, str(Path(__file__).parent / "src"))

from dotenv import load_dotenv

load_dotenv()


def _require_msal() -> None:
    method = os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower()
    if method != "msal":
        sys.stderr.write(
            f"auth_refresh: MICROSOFT_MCP_AUTH_METHOD is '{method}', must be 'msal'.\n"
            "Set MICROSOFT_MCP_AUTH_METHOD=msal to use this script.\n"
        )
        sys.exit(2)


def _resolve_env_args() -> dict:
    return {
        "tokens_dir": (
            Path(os.environ["MICROSOFT_MCP_TOKENS_DIR"])
            if os.getenv("MICROSOFT_MCP_TOKENS_DIR")
            else None
        ),
        "client_id": os.getenv("MICROSOFT_MCP_CLIENT_ID"),
        "tenant_id": os.getenv("MICROSOFT_MCP_TENANT_ID"),
    }


def _print_json(payload) -> None:
    print(json.dumps(payload, indent=2, default=str))


def _print_results_summary(results: list[dict]) -> None:
    if not results:
        print("No saved accounts found.")
        return
    print()
    for r in results:
        status = r.get("status", "?")
        ident = r.get("identifier", "?")
        exp = r.get("expires_at") or "-"
        line = f"  [{status:>11s}] {ident:<35s} expires_at={exp}"
        if r.get("error"):
            line += f"  error={r['error']}"
        print(line)


def _print_verify_summary(results: list[dict]) -> None:
    if not results:
        print("No saved accounts found.")
        return
    print()
    mismatches = 0
    for r in results:
        ident = r.get("identifier", "?")
        jwt_upn = r.get("jwt_upn")
        match = r.get("match", False)
        flag = "OK      " if match else "MISMATCH"
        if not match:
            mismatches += 1
        extra = ""
        if r.get("graph_userPrincipalName") is not None or r.get("graph_error"):
            extra = (
                f"  graph={r.get('graph_userPrincipalName') or r.get('graph_error')}"
            )
        if r.get("jwt_decode_error"):
            extra += f"  decode_error={r['jwt_decode_error']}"
        print(f"  [{flag}] {ident:<35s} jwt_upn={jwt_upn}{extra}")
    print()
    print(f"Summary: {len(results)} account(s), {mismatches} mismatch(es).")
    if mismatches:
        sys.exit(1)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Refresh, re-authenticate, or verify MSAL tokens.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "email",
        nargs="?",
        help="Account email to refresh. Omit to refresh all accounts.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Clear saved tokens for EMAIL and re-run the MSAL device-code flow. "
        "Requires an email argument.",
    )
    parser.add_argument(
        "--verify",
        action="store_true",
        help="Verify that saved tokens match their filenames (JWT-based check). "
        "Mutually exclusive with refresh/--force.",
    )
    parser.add_argument(
        "--live",
        action="store_true",
        help="With --verify: also call Graph /me for live tenant confirmation.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        help="Emit raw JSON output (default: human-readable summary).",
    )
    args = parser.parse_args()

    if args.verify:
        if args.force or args.email:
            parser.error("--verify is mutually exclusive with EMAIL and --force")
        _require_msal()
        from microsoft_mcp.auth_msal import verify_account_tokens

        env = _resolve_env_args()
        results = verify_account_tokens(tokens_dir=env["tokens_dir"], live=args.live)
        if args.json:
            _print_json(results)
            sys.exit(1 if any(not r.get("match") for r in results) else 0)
        _print_verify_summary(results)
        return

    if args.force:
        if not args.email:
            parser.error("--force requires an EMAIL argument")
        _require_msal()
        from microsoft_mcp.auth_msal import force_reauthenticate

        env = _resolve_env_args()
        result = force_reauthenticate(identifier=args.email, **env)
        if args.json:
            _print_json(result)
        else:
            print()
            print(f"  Re-authenticated: {result['identifier']}")
            print(f"  Signed in as:    {result.get('signed_in_as')}")
            print(f"  Expires at:      {result.get('expires_at')}")
            if (
                result.get("signed_in_as")
                and result["signed_in_as"].lower() != args.email.lower()
            ):
                print()
                print(
                    f"  WARNING: signed_in_as ({result['signed_in_as']}) does not "
                    f"match requested email ({args.email}). Saved tokens are "
                    f"labeled as '{args.email}' but belong to "
                    f"'{result['signed_in_as']}'."
                )
                sys.exit(1)
        return

    _require_msal()

    if args.email:
        from microsoft_mcp.auth_msal import refresh_account

        env = _resolve_env_args()
        result = refresh_account(identifier=args.email, **env)
        if args.json:
            _print_json(result)
        else:
            _print_results_summary([result])
        if result.get("status") == "failed":
            sys.exit(1)
        if result.get("status") == "missing":
            sys.exit(2)
        return

    from microsoft_mcp.auth_msal import refresh_all_accounts

    env = _resolve_env_args()
    results = refresh_all_accounts(**env)
    if args.json:
        _print_json(results)
    else:
        _print_results_summary(results)
    if any(r.get("status") == "failed" for r in results):
        sys.exit(1)


if __name__ == "__main__":
    main()
