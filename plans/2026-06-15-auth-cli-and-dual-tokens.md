# Auth CLI + Dual Graph/Outlook Tokens — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking. Dispatch ONE task per fresh subagent; review between tasks.

**Goal:** Give `microsoft-mcp` a polished first-class auth CLI — `microsoft-mcp auth refresh|verify|status|list|test|doctor` — that mirrors `outlook-creds`' `outlook auth refresh` UX (per-account `✓ Graph: Valid, expires 2026-06-15 21:25:06 UTC`), plus optional dual Graph + Outlook (`outlook.office365.com`) access tokens sharing one refresh token.

**Architecture:** The token-lifecycle engine already exists in `auth_msal.py` (`refresh_all_accounts`, `refresh_account`, `force_reauthenticate`, `verify_account_tokens`). We add a thin presentation/CLI layer (`auth_cli.py`) modeled exactly on the existing `signatures_cli.py` pattern — argparse subcommands, dual exposure via `server.main()` dispatch + a `microsoft-mcp-auth` console script. Zero-dependency ANSI color (TTY-gated, honors `NO_COLOR`). Dual-token support threads an `api_type` ("graph"|"outlook") through `MSALRefreshTokenAuth` so a single shared refresh token mints both access tokens. The untracked `auth_refresh.py` becomes a thin backward-compat shim.

**Tech Stack:** Python 3.12+, `argparse` (stdlib, matches `signatures_cli.py`), `msal`, `httpx` (already deps), `urllib` for the token-endpoint POST (already used by `auth_msal.py`). No new dependencies. Tests: `pytest` with the repo's single-account fixture policy (`TEST_EMAIL = "broach@cresa.com"`).

---

## Decisions Locked (from planning Q&A)

| Decision | Choice |
| --- | --- |
| Scope | **Core auth CLI** (`refresh`, `verify`, `status`, `list`, `test`, `doctor`) — no backup/cron |
| Dual tokens | **Include** separate `outlook.office365.com` token, shared refresh token, `--api=graph\|outlook\|both` |
| Color | **Zero-dep ANSI**, TTY-gated, honors `NO_COLOR` |
| CLI shape | **Both** — `microsoft-mcp auth <cmd>` subcommand AND `microsoft-mcp-auth` console script; logic in `auth_cli.py`; `auth_refresh.py` → thin shim |

---

## File Structure

| File | Disposition | Responsibility |
| --- | --- | --- |
| `src/microsoft_mcp/auth_cli.py` | **Create** | New argparse CLI: color/format helpers + `refresh`/`verify`/`status`/`list`/`test`/`doctor` handlers + `main(argv)` + `cli_main()`. Mirrors `signatures_cli.py`. |
| `src/microsoft_mcp/server.py` | Modify `main()` (`:73-80`) | Add `argv[0] == "auth"` dispatch before the heavy imports, next to the existing `signatures` dispatch. |
| `pyproject.toml` | Modify `[project.scripts]` (`:22-25`) | Add `microsoft-mcp-auth = "microsoft_mcp.auth_cli:cli_main"`. |
| `src/microsoft_mcp/auth_msal.py` | Modify | Thread `api_type` through `__init__`, path methods, `_save_tokens`, `_refresh_access_token`; add `GRAPH_SCOPE`/`OUTLOOK_SCOPE`; extend `refresh_all_accounts`/`refresh_account` with `api_type`. |
| `auth_refresh.py` | Replace body | Thin shim that calls `auth_cli.main(...)` for backward compat. |
| `tests/test_auth_cli.py` | **Create** | Unit tests for color/format helpers + every CLI subcommand (human + `--json` + exit codes). |
| `tests/test_auth_msal_dual_token.py` | **Create** | Unit tests for `api_type` path/scope/save behavior + dual-token refresh. |
| `tests/test_server_entry.py` | Modify (if exists) / else add case | Regression: `microsoft-mcp auth ...` dispatch routes to `auth_cli.main`. |
| `.claude/commands/auth-refresh.md`, `auth-verify.md` | Modify | Point at the new subcommand; add `/auth-status`. |
| `CLAUDE.md`, `README.md` | Modify | Document the new CLI surface + dual-token env/flags. |

**Token files on disk** (identifier = raw email, e.g. `broach@cresa.com`):
```
{id}_access_token.json          # graph access token + metadata  (api_type="graph")
{id}_access_only.txt            # graph raw access token
{id}_outlook_access_token.json  # outlook access token + metadata (api_type="outlook")  [NEW]
{id}_outlook_access_only.txt    # outlook raw access token                              [NEW]
{id}_refresh_only.txt           # SHARED refresh token (both api types)
```

---

## Waves & Dependency Graph

```
WAVE 0 — Pure helpers (no deps)
  T1 color+format helpers ─┐
                           │
WAVE 1 — CLI core          ▼
  T2 auth_cli skeleton + refresh/list/verify/status  (dep: T1)
  T3 server.main dispatch + console script           (dep: T2)
  T4 auth_refresh.py → shim                           (dep: T2)

WAVE 2 — Health commands (dep: T2; parallel to each other & to Wave 3)
  T5 auth status (read-only)
  T6 auth doctor
  T7 auth test (live /me)

WAVE 3 — Dual Graph/Outlook tokens (engine; dep: none for T8, T9 dep T8, T10 dep T8+T2)
  T8  api_type in MSALRefreshTokenAuth
  T9  api_type in refresh_all_accounts/refresh_account
  T10 --api flag in auth_cli refresh                  (dep: T8, T9, T2)

WAVE 4 — Docs, slash commands, regression (dep: all prior)
  T11 slash commands + docs
  T12 entry-point + tool-surface regression tests
```

**Dependency edges:** T2→{T3,T4,T5,T6,T7,T10}; T1→T2; T8→T9; {T8,T9,T2}→T10; everything→{T11,T12}.

**Parallelizable within a wave:** Wave 2 (T5/T6/T7) are independent of each other. Wave 3 T8 can start as soon as Wave 0 lands (it doesn't need the CLI). T3 and T4 are independent once T2 lands.

---

## WAVE 0 — Pure Helpers

### Task 1: ANSI color + UTC expiry formatting helpers

**Files:**
- Create: `src/microsoft_mcp/auth_cli.py` (helpers only — subcommands added in T2)
- Test: `tests/test_auth_cli.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_auth_cli.py
import os
import datetime as dt
import pytest
from microsoft_mcp import auth_cli

TEST_EMAIL = "broach@cresa.com"


def test_format_expiry_converts_stored_zulu_to_utc_human():
    assert (
        auth_cli._format_expiry("2026-06-15T21:25:06Z")
        == "2026-06-15 21:25:06 UTC"
    )


def test_format_expiry_handles_none_and_garbage():
    assert auth_cli._format_expiry(None) == "unknown"
    assert auth_cli._format_expiry("not-a-date") == "not-a-date"


def test_color_disabled_when_no_color_env(monkeypatch):
    monkeypatch.setenv("NO_COLOR", "1")
    assert auth_cli._color_enabled() is False
    assert auth_cli._c("x", "green") == "x"  # no escape codes


def test_color_disabled_for_non_tty(monkeypatch):
    monkeypatch.delenv("NO_COLOR", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_FORCE_COLOR", raising=False)

    class _NotATty:
        def isatty(self):
            return False

    assert auth_cli._color_enabled(_NotATty()) is False


def test_color_forced_on(monkeypatch):
    monkeypatch.delenv("NO_COLOR", raising=False)
    monkeypatch.setenv("MICROSOFT_MCP_FORCE_COLOR", "1")
    assert auth_cli._color_enabled() is True
    assert auth_cli._c("ok", "green") == "\033[32mok\033[0m"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_auth_cli.py -v`
Expected: FAIL — `ModuleNotFoundError: No module named 'microsoft_mcp.auth_cli'`

- [ ] **Step 3: Write minimal implementation**

```python
# src/microsoft_mcp/auth_cli.py
"""CLI for refreshing, verifying, and inspecting MSAL tokens.

Exposed two ways (mirrors signatures_cli.py):
- ``microsoft-mcp-auth <cmd> ...``   (standalone console script)
- ``microsoft-mcp auth <cmd> ...``   (subcommand of the main entry point)

Both routes share ``main()``. Only the MSAL auth method is supported; the
Azure SDK path manages its own token cache.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
import sys
from pathlib import Path
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
# Formatting

def _format_expiry(expires_at: str | None) -> str:
    """Convert stored '%Y-%m-%dT%H:%M:%SZ' to '%Y-%m-%d %H:%M:%S UTC'."""
    if not expires_at:
        return "unknown"
    raw = expires_at.replace("Z", "+00:00")
    try:
        parsed = dt.datetime.fromisoformat(raw)
    except ValueError:
        return expires_at
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=dt.timezone.utc)
    return parsed.astimezone(dt.timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_auth_cli.py -v`
Expected: PASS (5 passed)

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/auth_cli.py tests/test_auth_cli.py
git commit -m "feat(auth-cli): add zero-dep ANSI color + UTC expiry formatting helpers"
```

---

## WAVE 1 — CLI Core

### Task 2: `auth_cli` argparse skeleton + `refresh` / `list` / `verify` / `status`

**Files:**
- Modify: `src/microsoft_mcp/auth_cli.py`
- Test: `tests/test_auth_cli.py`

> Reuses `auth_msal.refresh_all_accounts`, `refresh_account`, `verify_account_tokens` (already present). Env resolution mirrors `auth_refresh.py:_resolve_env_args` / `_require_msal`.

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_auth_cli.py
import json as _json


def _write_graph_token(tmp_path, identifier, expires_delta_seconds=3600):
    now = dt.datetime.now(dt.timezone.utc)
    expires = now + dt.timedelta(seconds=expires_delta_seconds)
    payload = {
        "email": identifier,
        "access_token": "header.eyJ1cG4iOiJ4In0.sig",  # decodable junk ok here
        "token_type": "Bearer",
        "expires_in": expires_delta_seconds,
        "expires_at": expires.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "refreshed_at": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "scopes": "https://graph.microsoft.com/.default offline_access",
        "api_type": "graph",
    }
    (tmp_path / f"{identifier}_access_token.json").write_text(_json.dumps(payload))
    (tmp_path / f"{identifier}_refresh_only.txt").write_text("fake-refresh-token")
    return payload


def _msal_env(monkeypatch, tmp_path):
    monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", "msal")
    monkeypatch.setenv("MICROSOFT_MCP_TOKENS_DIR", str(tmp_path))
    monkeypatch.setenv("MICROSOFT_MCP_CLIENT_ID", "d3590ed6-52b3-4102-aeff-aad2292ab01c")
    monkeypatch.delenv("MICROSOFT_MCP_TENANT_ID", raising=False)


def test_refresh_all_prints_per_account_status(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL)
    rc = auth_cli.main(["refresh"])
    out = capsys.readouterr().out
    assert rc == 0
    assert TEST_EMAIL in out
    assert "Graph: Valid, expires" in out
    assert "UTC" in out


def test_refresh_requires_msal(monkeypatch, tmp_path):
    monkeypatch.setenv("MICROSOFT_MCP_AUTH_METHOD", "azure")
    with pytest.raises(SystemExit) as exc:
        auth_cli.main(["refresh"])
    assert exc.value.code == 2


def test_refresh_json_output(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL)
    rc = auth_cli.main(["refresh", "--json"])
    out = capsys.readouterr().out
    assert rc == 0
    parsed = _json.loads(out)
    assert parsed[0]["identifier"] == TEST_EMAIL
    assert parsed[0]["status"] == "valid"


def test_list_accounts_human(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL)
    rc = auth_cli.main(["list"])
    out = capsys.readouterr().out
    assert rc == 0
    assert TEST_EMAIL in out
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_auth_cli.py -k "refresh or list" -v`
Expected: FAIL — `AttributeError: module 'microsoft_mcp.auth_cli' has no attribute 'main'`

- [ ] **Step 3: Write minimal implementation**

Append to `src/microsoft_mcp/auth_cli.py`:

```python
# ---------------------------------------------------------------------------
# Environment / guards

def _require_msal() -> None:
    method = os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower()
    if method != "msal":
        sys.stderr.write(
            f"auth: MICROSOFT_MCP_AUTH_METHOD is '{method}', must be 'msal'.\n"
            "Set MICROSOFT_MCP_AUTH_METHOD=msal to use this command.\n"
        )
        raise SystemExit(2)


def _env_kwargs() -> dict[str, Any]:
    return {
        "tokens_dir": (
            Path(os.environ["MICROSOFT_MCP_TOKENS_DIR"])
            if os.getenv("MICROSOFT_MCP_TOKENS_DIR")
            else None
        ),
        "client_id": os.getenv("MICROSOFT_MCP_CLIENT_ID"),
        "tenant_id": os.getenv("MICROSOFT_MCP_TENANT_ID"),
    }


def _print_json(payload: Any) -> None:
    print(json.dumps(payload, indent=2, default=str))


# ---------------------------------------------------------------------------
# Result rendering

def _print_refresh_results(results: list[dict[str, Any]], api_label: str = "Graph") -> None:
    if not results:
        print("No saved accounts found.")
        return
    for r in results:
        ident = r.get("identifier", "?")
        status = r.get("status")
        print(ident)
        if status in ("valid", "refreshed", "reauthenticated"):
            word = {"valid": "Valid", "refreshed": "Refreshed",
                    "reauthenticated": "Re-authenticated"}[status]
            exp = _format_expiry(r.get("expires_at"))
            print(_c(f"  ✓ {api_label}: {word}, expires {exp}", "green"))
        elif status == "missing":
            print(_c(f"  ✗ {api_label}: no saved token", "red"))
        else:  # failed / unknown
            err = r.get("error") or "unknown error"
            print(_c(f"  ✗ {api_label} refresh failed: {err}", "red"))


# ---------------------------------------------------------------------------
# Subcommand handlers

def _cmd_refresh(args: argparse.Namespace) -> int:
    _require_msal()
    env = _env_kwargs()
    if args.force:
        if not args.email:
            raise SystemExit("error: --force requires an EMAIL argument")
        from .auth_msal import force_reauthenticate

        result = force_reauthenticate(identifier=args.email, **env)
        if args.json:
            _print_json(result)
        else:
            _print_refresh_results([result])
            if (
                result.get("signed_in_as")
                and result["signed_in_as"].lower() != args.email.lower()
            ):
                print(
                    _c(
                        f"  ⚠ WARNING: signed_in_as ({result['signed_in_as']}) "
                        f"does not match requested email ({args.email}).",
                        "yellow",
                    )
                )
                return 1
        return 0

    if args.email:
        from .auth_msal import refresh_account

        result = refresh_account(identifier=args.email, **env)
        if args.json:
            _print_json(result)
        else:
            _print_refresh_results([result])
        if result.get("status") == "failed":
            return 1
        if result.get("status") == "missing":
            return 2
        return 0

    from .auth_msal import refresh_all_accounts

    results = refresh_all_accounts(**env)
    if args.json:
        _print_json(results)
    else:
        _print_refresh_results(results)
    return 1 if any(r.get("status") == "failed" for r in results) else 0


def _cmd_verify(args: argparse.Namespace) -> int:
    _require_msal()
    from .auth_msal import verify_account_tokens

    env = _env_kwargs()
    results = verify_account_tokens(tokens_dir=env["tokens_dir"], live=args.live)
    if args.json:
        _print_json(results)
        return 1 if any(not r.get("match") for r in results) else 0
    if not results:
        print("No saved accounts found.")
        return 0
    mismatches = 0
    for r in results:
        ident = r.get("identifier", "?")
        match = r.get("match", False)
        if match:
            print(_c(f"  ✓ {ident}  jwt_upn={r.get('jwt_upn')}", "green"))
        else:
            mismatches += 1
            extra = ""
            if r.get("graph_error"):
                extra = f"  graph_error={r['graph_error']}"
            if r.get("jwt_decode_error"):
                extra += f"  decode_error={r['jwt_decode_error']}"
            print(_c(f"  ✗ {ident}  jwt_upn={r.get('jwt_upn')}{extra}", "red"))
    print()
    print(f"Summary: {len(results)} account(s), {mismatches} mismatch(es).")
    return 1 if mismatches else 0


def _cmd_status(args: argparse.Namespace) -> int:
    # Implemented in Task 5.
    raise NotImplementedError


def _cmd_list(args: argparse.Namespace) -> int:
    _require_msal()
    env = _env_kwargs()
    tokens_dir = env["tokens_dir"] or _default_tokens_dir()
    rows = _enumerate_accounts(tokens_dir)
    if args.json:
        _print_json(rows)
        return 0
    if not rows:
        print("No saved accounts found.")
        return 0
    for row in rows:
        print(f"{row['identifier']}  (expires {_format_expiry(row['expires_at'])})")
    return 0


# ---------------------------------------------------------------------------
# Shared disk helpers

def _default_tokens_dir() -> Path:
    default_dir = Path.home() / ".config" / "microsoft-mcp" / "tokens"
    return Path(os.getenv("MICROSOFT_MCP_TOKENS_DIR", str(default_dir)))


def _enumerate_accounts(tokens_dir: Path) -> list[dict[str, Any]]:
    """List graph accounts (skips the *_outlook_access_token.json siblings)."""
    if not tokens_dir.exists():
        return []
    rows: list[dict[str, Any]] = []
    for f in sorted(tokens_dir.glob("*_access_token.json")):
        if f.stem.endswith("_outlook_access_token"):
            continue
        identifier = f.stem[: -len("_access_token")]
        try:
            data = json.loads(f.read_text())
        except Exception:
            data = {}
        rows.append(
            {
                "identifier": identifier,
                "email": data.get("email", identifier),
                "expires_at": data.get("expires_at"),
            }
        )
    return rows


# ---------------------------------------------------------------------------
# Argparse wiring

def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="microsoft-mcp-auth",
        description="Refresh, verify, and inspect MSAL tokens for microsoft-mcp.",
    )
    sub = parser.add_subparsers(dest="cmd", required=True, metavar="COMMAND")

    p_refresh = sub.add_parser("refresh", help="refresh tokens (all / one / force re-auth)")
    p_refresh.add_argument("email", nargs="?", default=None,
                           help="account email; omit to refresh all")
    p_refresh.add_argument("--force", action="store_true",
                           help="clear EMAIL's tokens and re-run device-code flow")
    p_refresh.add_argument("--json", action="store_true")
    p_refresh.set_defaults(func=_cmd_refresh)

    p_verify = sub.add_parser("verify", help="verify tokens match their filenames (JWT)")
    p_verify.add_argument("--live", action="store_true", help="also call Graph /me")
    p_verify.add_argument("--json", action="store_true")
    p_verify.set_defaults(func=_cmd_verify)

    p_status = sub.add_parser("status", help="read-only token health (no network)")
    p_status.add_argument("--json", action="store_true")
    p_status.set_defaults(func=_cmd_status)

    p_list = sub.add_parser("list", help="list saved accounts")
    p_list.add_argument("--json", action="store_true")
    p_list.set_defaults(func=_cmd_list)

    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    return args.func(args)


def cli_main() -> None:
    """Console-script entry point."""
    sys.exit(main(sys.argv[1:]))


if __name__ == "__main__":
    cli_main()
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_auth_cli.py -v`
Expected: PASS (all)

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/auth_cli.py tests/test_auth_cli.py
git commit -m "feat(auth-cli): add refresh/verify/list subcommands with colored per-account output"
```

---

### Task 3: Wire `microsoft-mcp auth ...` dispatch + `microsoft-mcp-auth` console script

**Files:**
- Modify: `src/microsoft_mcp/server.py:73-80`
- Modify: `pyproject.toml:22-25`
- Test: `tests/test_server_entry.py` (add case; create file if absent)

- [ ] **Step 1: Write the failing test**

```python
# tests/test_server_entry.py  (add this test; keep existing ones)
import sys
from unittest import mock


def test_auth_subcommand_dispatches_to_auth_cli(monkeypatch):
    from microsoft_mcp import server

    called = {}

    def fake_main(argv):
        called["argv"] = argv
        return 0

    monkeypatch.setattr("microsoft_mcp.auth_cli.main", fake_main)
    monkeypatch.setattr(sys, "argv", ["microsoft-mcp", "auth", "refresh", "--json"])
    with mock.patch.object(sys, "exit") as fake_exit:
        server.main()
    assert called["argv"] == ["refresh", "--json"]
    fake_exit.assert_called_once_with(0)
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_server_entry.py::test_auth_subcommand_dispatches_to_auth_cli -v`
Expected: FAIL — `server.main()` falls through to the normal startup path (no dispatch), so `auth_cli.main` is never called.

- [ ] **Step 3: Write minimal implementation**

In `src/microsoft_mcp/server.py`, extend the dispatch block (currently `:76-80`):

```python
    argv = sys.argv[1:]
    if argv and argv[0] == "signatures":
        from microsoft_mcp import signatures_cli

        sys.exit(signatures_cli.main(argv[1:]))

    if argv and argv[0] == "auth":
        from microsoft_mcp import auth_cli

        sys.exit(auth_cli.main(argv[1:]))
```

In `pyproject.toml`, add to `[project.scripts]`:

```toml
[project.scripts]
microsoft-mcp = "microsoft_mcp.server:main"
microsoft-mcp-utcp-config = "microsoft_mcp.utcp_bridge_config:main"
microsoft-mcp-signatures = "microsoft_mcp.signatures_cli:cli_main"
microsoft-mcp-auth = "microsoft_mcp.auth_cli:cli_main"
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_server_entry.py -v`
Expected: PASS

Manual smoke (with MSAL env set):
Run: `uv run microsoft-mcp auth list`
Expected: lists saved accounts (or "No saved accounts found.")

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/server.py pyproject.toml tests/test_server_entry.py
git commit -m "feat(auth-cli): expose via microsoft-mcp auth subcommand + microsoft-mcp-auth script"
```

---

### Task 4: Convert `auth_refresh.py` to a thin backward-compat shim

**Files:**
- Modify: `auth_refresh.py` (repo root, untracked)
- Test: `tests/test_auth_refresh_shim.py` (create)

> The standalone `auth_refresh.py` and the `/auth-refresh` `/auth-verify` slash commands invoke `uv run auth_refresh.py $ARGUMENTS`. Preserve that interface but delegate to `auth_cli` so logic lives in one place. The old positional/flag interface (`auth_refresh.py <email>`, `--force`, `--verify [--live]`) maps onto the new subcommands.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_auth_refresh_shim.py
import sys
import runpy
import pytest


def test_shim_maps_verify_to_auth_cli(monkeypatch):
    import auth_refresh

    captured = {}

    def fake_main(argv):
        captured["argv"] = argv
        return 0

    monkeypatch.setattr("microsoft_mcp.auth_cli.main", fake_main)
    monkeypatch.setattr(sys, "argv", ["auth_refresh.py", "--verify", "--live"])
    with pytest.raises(SystemExit) as exc:
        auth_refresh.main()
    assert captured["argv"] == ["verify", "--live"]
    assert exc.value.code == 0


def test_shim_maps_force_email(monkeypatch):
    import auth_refresh

    captured = {}
    monkeypatch.setattr("microsoft_mcp.auth_cli.main", lambda argv: captured.setdefault("argv", argv) or 0)
    monkeypatch.setattr(sys, "argv", ["auth_refresh.py", "--force", "broach@cresa.com"])
    with pytest.raises(SystemExit):
        auth_refresh.main()
    assert captured["argv"] == ["refresh", "broach@cresa.com", "--force"]


def test_shim_maps_bare_refresh_all(monkeypatch):
    import auth_refresh

    captured = {}
    monkeypatch.setattr("microsoft_mcp.auth_cli.main", lambda argv: captured.setdefault("argv", argv) or 0)
    monkeypatch.setattr(sys, "argv", ["auth_refresh.py"])
    with pytest.raises(SystemExit):
        auth_refresh.main()
    assert captured["argv"] == ["refresh"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_auth_refresh_shim.py -v`
Expected: FAIL — the current `auth_refresh.main()` calls `auth_msal` functions directly, not `auth_cli.main`.

- [ ] **Step 3: Write minimal implementation**

Replace the body of `auth_refresh.py` with:

```python
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

from dotenv import load_dotenv

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
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_auth_refresh_shim.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add auth_refresh.py tests/test_auth_refresh_shim.py
git commit -m "refactor(auth): make auth_refresh.py a thin shim over auth_cli"
```

---

## WAVE 2 — Health Commands (depend on T2; independent of each other)

### Task 5: `auth status` — read-only token health (no network)

**Files:**
- Modify: `src/microsoft_mcp/auth_cli.py` (replace `_cmd_status` stub)
- Test: `tests/test_auth_cli.py`

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_auth_cli.py
def test_status_reports_valid_without_network(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL, expires_delta_seconds=3600)
    rc = auth_cli.main(["status"])
    out = capsys.readouterr().out
    assert rc == 0
    assert TEST_EMAIL in out
    assert "Graph: Valid" in out
    assert "UTC" in out


def test_status_reports_expired(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL, expires_delta_seconds=-10)
    rc = auth_cli.main(["status"])
    out = capsys.readouterr().out
    assert rc == 1  # at least one account is not valid
    assert "Expired" in out or "expired" in out


def test_status_json(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL)
    rc = auth_cli.main(["status", "--json"])
    parsed = _json.loads(capsys.readouterr().out)
    assert parsed[0]["identifier"] == TEST_EMAIL
    assert parsed[0]["valid"] is True
    assert parsed[0]["has_refresh_token"] is True
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_auth_cli.py -k status -v`
Expected: FAIL — `_cmd_status` raises `NotImplementedError`.

- [ ] **Step 3: Write minimal implementation**

Replace the `_cmd_status` stub and add helpers in `auth_cli.py`:

```python
def _account_health(tokens_dir: Path) -> list[dict[str, Any]]:
    """Disk-only health (no network). Buffer matches auth_msal (60s)."""
    BUFFER_SECONDS = 60
    rows: list[dict[str, Any]] = []
    for row in _enumerate_accounts(tokens_dir):
        identifier = row["identifier"]
        expires_at = row["expires_at"]
        valid = False
        remaining = None
        if expires_at:
            raw = expires_at.replace("Z", "+00:00")
            try:
                exp = dt.datetime.fromisoformat(raw)
                if exp.tzinfo is None:
                    exp = exp.replace(tzinfo=dt.timezone.utc)
                remaining = (exp - dt.datetime.now(dt.timezone.utc)).total_seconds()
                valid = remaining > BUFFER_SECONDS
            except ValueError:
                pass
        has_refresh = (tokens_dir / f"{identifier}_refresh_only.txt").exists()
        rows.append(
            {
                "identifier": identifier,
                "expires_at": expires_at,
                "valid": valid,
                "remaining_seconds": int(remaining) if remaining is not None else None,
                "has_refresh_token": has_refresh,
            }
        )
    return rows


def _cmd_status(args: argparse.Namespace) -> int:
    _require_msal()
    tokens_dir = (_env_kwargs()["tokens_dir"]) or _default_tokens_dir()
    rows = _account_health(tokens_dir)
    if args.json:
        _print_json(rows)
        return 0 if rows and all(r["valid"] for r in rows) else (0 if not rows else 1)
    if not rows:
        print("No saved accounts found.")
        return 0
    all_valid = True
    for r in rows:
        print(r["identifier"])
        exp = _format_expiry(r["expires_at"])
        if r["valid"]:
            print(_c(f"  ✓ Graph: Valid, expires {exp}", "green"))
        else:
            all_valid = False
            refresh_note = "" if r["has_refresh_token"] else " (no refresh token!)"
            print(_c(f"  ✗ Graph: Expired, expired {exp}{refresh_note}", "red"))
    return 0 if all_valid else 1
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_auth_cli.py -k status -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/auth_cli.py tests/test_auth_cli.py
git commit -m "feat(auth-cli): add read-only 'auth status' health command"
```

---

### Task 6: `auth doctor` — diagnose duplicates / perms / mismatches / missing refresh tokens

**Files:**
- Modify: `src/microsoft_mcp/auth_cli.py`
- Test: `tests/test_auth_cli.py`

> Reuses `verify_account_tokens(live=False)` for the JWT/filename match, then layers on: file-permission check (expect `0o600`), missing `_refresh_only.txt`, and duplicate-identity detection (two filenames whose JWT `upn` is the same).

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_auth_cli.py
import stat as _stat


def test_doctor_flags_missing_refresh_token(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL)
    (tmp_path / f"{TEST_EMAIL}_refresh_only.txt").unlink()  # remove refresh token
    rc = auth_cli.main(["doctor"])
    out = capsys.readouterr().out
    assert rc == 1
    assert "refresh token" in out.lower()


def test_doctor_flags_loose_permissions(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL)
    f = tmp_path / f"{TEST_EMAIL}_access_token.json"
    f.chmod(0o644)  # too open
    rc = auth_cli.main(["doctor"])
    out = capsys.readouterr().out
    assert rc == 1
    assert "permission" in out.lower()


def test_doctor_clean_passes(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL)
    (tmp_path / f"{TEST_EMAIL}_access_token.json").chmod(0o600)
    (tmp_path / f"{TEST_EMAIL}_refresh_only.txt").chmod(0o600)
    # JWT upn in the fixture token decodes to "x" (won't match TEST_EMAIL);
    # use a token whose upn matches so verify passes:
    monkeypatch.setattr(
        "microsoft_mcp.auth_msal.verify_account_tokens",
        lambda **kw: [{"identifier": TEST_EMAIL, "jwt_upn": TEST_EMAIL, "match": True}],
    )
    rc = auth_cli.main(["doctor"])
    out = capsys.readouterr().out
    assert rc == 0
    assert "OK" in out or "no issues" in out.lower()
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_auth_cli.py -k doctor -v`
Expected: FAIL — no `doctor` subcommand registered (`SystemExit` from argparse "invalid choice").

- [ ] **Step 3: Write minimal implementation**

Add to `auth_cli.py` (handler + parser registration):

```python
def _cmd_doctor(args: argparse.Namespace) -> int:
    _require_msal()
    from .auth_msal import verify_account_tokens

    tokens_dir = (_env_kwargs()["tokens_dir"]) or _default_tokens_dir()
    verify = {v["identifier"]: v for v in verify_account_tokens(tokens_dir=tokens_dir)}
    health = _account_health(tokens_dir)

    issues: list[str] = []
    upn_to_ids: dict[str, list[str]] = {}

    for r in health:
        identifier = r["identifier"]
        # permissions
        f = tokens_dir / f"{identifier}_access_token.json"
        mode = _stat_mode(f)
        if mode is not None and (mode & 0o077):
            issues.append(f"{identifier}: loose permissions on token file ({oct(mode)})")
        # refresh token presence
        if not r["has_refresh_token"]:
            issues.append(f"{identifier}: missing refresh token (re-auth required)")
        # expiry
        if not r["valid"]:
            issues.append(f"{identifier}: access token expired")
        # JWT / filename match + duplicate identity
        v = verify.get(identifier, {})
        if v and not v.get("match", True):
            issues.append(
                f"{identifier}: token JWT upn ({v.get('jwt_upn')}) does not match filename"
            )
        upn = (v.get("jwt_upn") or "").lower()
        if upn:
            upn_to_ids.setdefault(upn, []).append(identifier)

    for upn, ids in upn_to_ids.items():
        if len(ids) > 1:
            issues.append(f"duplicate identity {upn} across files: {', '.join(ids)}")

    if args.json:
        _print_json({"issues": issues, "accounts": health})
        return 1 if issues else 0

    if not health:
        print("No saved accounts found.")
        return 0
    if not issues:
        print(_c(f"  ✓ OK — {len(health)} account(s), no issues found.", "green"))
        return 0
    print(_c(f"  ✗ {len(issues)} issue(s) found:", "red"))
    for msg in issues:
        print(_c(f"    - {msg}", "yellow"))
    return 1


def _stat_mode(path: Path) -> int | None:
    try:
        import stat as _s

        return _s.S_IMODE(path.stat().st_mode)
    except OSError:
        return None
```

Register in `_build_parser()` (before `return parser`):

```python
    p_doctor = sub.add_parser("doctor", help="diagnose token health (perms, dups, mismatches)")
    p_doctor.add_argument("--json", action="store_true")
    p_doctor.set_defaults(func=_cmd_doctor)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_auth_cli.py -k doctor -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/auth_cli.py tests/test_auth_cli.py
git commit -m "feat(auth-cli): add 'auth doctor' health diagnostics"
```

---

### Task 7: `auth test` — live Graph `/me` per account

**Files:**
- Modify: `src/microsoft_mcp/auth_cli.py`
- Test: `tests/test_auth_cli.py`

> Reuses `verify_account_tokens(live=True)` — it already performs the `GET /me` call and records `graph_userPrincipalName` / `graph_error`. `auth test` is a friendly wrapper that renders pass/fail per account.

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_auth_cli.py
def test_auth_test_reports_live_identity(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL)
    monkeypatch.setattr(
        "microsoft_mcp.auth_msal.verify_account_tokens",
        lambda **kw: [
            {
                "identifier": TEST_EMAIL,
                "graph_userPrincipalName": TEST_EMAIL,
                "graph_error": None,
                "match": True,
            }
        ],
    )
    rc = auth_cli.main(["test"])
    out = capsys.readouterr().out
    assert rc == 0
    assert TEST_EMAIL in out
    assert "/me" in out or "OK" in out


def test_auth_test_reports_failure(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    _write_graph_token(tmp_path, TEST_EMAIL)
    monkeypatch.setattr(
        "microsoft_mcp.auth_msal.verify_account_tokens",
        lambda **kw: [
            {
                "identifier": TEST_EMAIL,
                "graph_userPrincipalName": None,
                "graph_error": "HTTP 401: token expired",
                "match": False,
            }
        ],
    )
    rc = auth_cli.main(["test"])
    out = capsys.readouterr().out
    assert rc == 1
    assert "401" in out
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_auth_cli.py -k auth_test -v`
Expected: FAIL — no `test` subcommand registered.

- [ ] **Step 3: Write minimal implementation**

Add handler + parser registration in `auth_cli.py`:

```python
def _cmd_test(args: argparse.Namespace) -> int:
    _require_msal()
    from .auth_msal import verify_account_tokens

    tokens_dir = (_env_kwargs()["tokens_dir"]) or _default_tokens_dir()
    results = verify_account_tokens(tokens_dir=tokens_dir, live=True)
    if args.json:
        _print_json(results)
        return 1 if any(r.get("graph_error") for r in results) else 0
    if not results:
        print("No saved accounts found.")
        return 0
    failures = 0
    for r in results:
        ident = r.get("identifier", "?")
        if r.get("graph_error"):
            failures += 1
            print(_c(f"  ✗ {ident}: /me failed — {r['graph_error']}", "red"))
        else:
            upn = r.get("graph_userPrincipalName") or "?"
            print(_c(f"  ✓ {ident}: /me OK — {upn}", "green"))
    return 1 if failures else 0
```

Register in `_build_parser()`:

```python
    p_test = sub.add_parser("test", help="live test each token against Graph /me")
    p_test.add_argument("--json", action="store_true")
    p_test.set_defaults(func=_cmd_test)
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_auth_cli.py -k auth_test -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/auth_cli.py tests/test_auth_cli.py
git commit -m "feat(auth-cli): add 'auth test' live Graph /me check"
```

---

## WAVE 3 — Dual Graph / Outlook Tokens

### Task 8: Thread `api_type` through `MSALRefreshTokenAuth`

**Files:**
- Modify: `src/microsoft_mcp/auth_msal.py` (constants `:54`; `__init__` `:130-189`; path methods `:210-220`; `_save_tokens` `:282-324`; `_refresh_access_token` scope block `:372-398`; `clear_cache` `:711-728`)
- Test: `tests/test_auth_msal_dual_token.py`

> The shared refresh token (`{id}_refresh_only.txt`) is NOT keyed by api_type. Only the access-token JSON, raw access file, and `api_type`/scope vary.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_auth_msal_dual_token.py
import json
from pathlib import Path
from microsoft_mcp import auth_msal

TEST_EMAIL = "broach@cresa.com"


def _auth(tmp_path, api_type):
    return auth_msal.MSALRefreshTokenAuth(
        tokens_dir=tmp_path,
        client_id="d3590ed6-52b3-4102-aeff-aad2292ab01c",
        tenant_id="common",
        account_identifier=TEST_EMAIL,
        api_type=api_type,
    )


def test_graph_paths_unchanged(tmp_path):
    a = _auth(tmp_path, "graph")
    assert a._access_token_json_path().name == f"{TEST_EMAIL}_access_token.json"
    assert a._access_token_raw_path().name == f"{TEST_EMAIL}_access_only.txt"
    assert a._refresh_token_path().name == f"{TEST_EMAIL}_refresh_only.txt"


def test_outlook_paths_use_outlook_suffix(tmp_path):
    a = _auth(tmp_path, "outlook")
    assert a._access_token_json_path().name == f"{TEST_EMAIL}_outlook_access_token.json"
    assert a._access_token_raw_path().name == f"{TEST_EMAIL}_outlook_access_only.txt"
    # refresh token is SHARED, not suffixed
    assert a._refresh_token_path().name == f"{TEST_EMAIL}_refresh_only.txt"


def test_save_records_api_type(tmp_path):
    a = _auth(tmp_path, "outlook")
    a._save_tokens("tok", "refresh", 3600, auth_msal.OUTLOOK_SCOPE, email=TEST_EMAIL)
    data = json.loads(a._access_token_json_path().read_text())
    assert data["api_type"] == "outlook"


def test_default_scope_selection(tmp_path):
    g = _auth(tmp_path, "graph")
    o = _auth(tmp_path, "outlook")
    assert g._default_scope() == auth_msal.GRAPH_SCOPE
    assert o._default_scope() == auth_msal.OUTLOOK_SCOPE


def test_api_type_defaults_to_graph(tmp_path):
    a = auth_msal.MSALRefreshTokenAuth(
        tokens_dir=tmp_path, account_identifier=TEST_EMAIL
    )
    assert a.api_type == "graph"
    assert a._access_token_json_path().name == f"{TEST_EMAIL}_access_token.json"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_auth_msal_dual_token.py -v`
Expected: FAIL — `__init__` has no `api_type` kwarg (`TypeError`).

- [ ] **Step 3: Write minimal implementation**

In `auth_msal.py`:

Add constants near `:54`:

```python
GRAPH_SCOPE = "https://graph.microsoft.com/.default offline_access"
OUTLOOK_SCOPE = "https://outlook.office365.com/.default offline_access"
```

Add `api_type` to `__init__` signature (`:130-136`) and store it (after `self.account_identifier = ...` at `:167`):

```python
    def __init__(
        self,
        tokens_dir: Optional[Path] = None,
        client_id: Optional[str] = None,
        tenant_id: Optional[str] = None,
        account_identifier: Optional[str] = None,
        api_type: str = "graph",
    ):
```

```python
        # API resource this instance targets ("graph" or "outlook"). Only the
        # access-token files and refresh scope vary; the refresh token is shared.
        self.api_type = api_type if api_type in ("graph", "outlook") else "graph"
```

Add a scope helper (anywhere in the class, e.g. after `_get_msal_app`):

```python
    def _default_scope(self) -> str:
        return OUTLOOK_SCOPE if self.api_type == "outlook" else GRAPH_SCOPE
```

Update the two access-token path methods (`:210-220`) — leave `_refresh_token_path` unchanged:

```python
    def _access_token_json_path(self) -> Path:
        """Path to structured access token JSON file."""
        suffix = "_outlook_access_token" if self.api_type == "outlook" else "_access_token"
        return self.tokens_dir / f"{self.account_identifier}{suffix}.json"

    def _access_token_raw_path(self) -> Path:
        """Path to raw access token file."""
        suffix = "_outlook_access_only" if self.api_type == "outlook" else "_access_only"
        return self.tokens_dir / f"{self.account_identifier}{suffix}.txt"
```

In `_save_tokens` (`:303-312`), make `api_type` dynamic:

```python
        access_token_data = {
            "email": email or self.account_identifier,
            "access_token": access_token,
            "token_type": "Bearer",
            "expires_in": expires_in,
            "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "refreshed_at": now.strftime("%Y-%m-%dT%H:%M:%SZ"),
            "scopes": scopes,
            "api_type": self.api_type,
        }
```

In `_refresh_access_token` (`:393-398`), when there are no saved scopes, fall back to the api-type default:

```python
        if parts:
            if "offline_access" not in parts:
                parts.append("offline_access")
            scopes = " ".join(parts)
        else:
            scopes = self._default_scope()
```

In `clear_cache` (`:715-719`), include both access files for this api_type (the path methods already resolve to the right files; no change needed if `clear_cache` keeps using `self._access_token_json_path()` / `self._access_token_raw_path()`). Verify it still lists `_access_token_json_path()`, `_refresh_token_path()`, `_access_token_raw_path()` — it does. No edit required.

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_auth_msal_dual_token.py -v`
Expected: PASS

Regression — existing auth tests still pass:
Run: `uv run pytest tests/test_auth_msal.py tests/test_refresh_all_accounts.py tests/test_graph_401_retry.py -v`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/auth_msal.py tests/test_auth_msal_dual_token.py
git commit -m "feat(auth): add api_type (graph/outlook) to MSALRefreshTokenAuth with shared refresh token"
```

---

### Task 9: `api_type` in `refresh_all_accounts` / `refresh_account`

**Files:**
- Modify: `src/microsoft_mcp/auth_msal.py` (`refresh_all_accounts` `:735-848`; `refresh_account` `:1019-1107`)
- Test: `tests/test_auth_msal_dual_token.py`

> For `outlook`, the token file enumerated is still the graph `*_access_token.json` (to discover identifiers), but the probe is built with `api_type="outlook"`. Minting an outlook token uses the SHARED refresh token via `_do_refresh_locked()` (which loads `_refresh_only.txt`, POSTs with the outlook scope, and saves `_outlook_access_token.json`).

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_auth_msal_dual_token.py
import datetime as dt


def _seed_graph_account(tmp_path, identifier, valid=True):
    now = dt.datetime.now(dt.timezone.utc)
    delta = 3600 if valid else -10
    exp = now + dt.timedelta(seconds=delta)
    (tmp_path / f"{identifier}_access_token.json").write_text(json.dumps({
        "email": identifier, "access_token": "a.b.c", "token_type": "Bearer",
        "expires_in": delta, "expires_at": exp.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "scopes": auth_msal.GRAPH_SCOPE, "api_type": "graph",
    }))
    (tmp_path / f"{identifier}_refresh_only.txt").write_text("shared-refresh")


def test_refresh_account_outlook_uses_outlook_probe(tmp_path, monkeypatch):
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)

    captured = {}

    def fake_refresh(self, refresh_token):
        captured["api_type"] = self.api_type
        captured["scope_default"] = self._default_scope()
        return {"access_token": "new", "refresh_token": "shared-refresh",
                "expires_in": 3600, "scope": auth_msal.OUTLOOK_SCOPE}

    monkeypatch.setattr(auth_msal.MSALRefreshTokenAuth, "_refresh_access_token", fake_refresh)
    result = auth_msal.refresh_account(TEST_EMAIL, tokens_dir=tmp_path, api_type="outlook")
    assert result["status"] == "refreshed"
    assert captured["api_type"] == "outlook"
    # outlook token file was written
    assert (tmp_path / f"{TEST_EMAIL}_outlook_access_token.json").exists()


def test_refresh_all_accounts_both_returns_two_entries(tmp_path, monkeypatch):
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)
    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth, "_refresh_access_token",
        lambda self, rt: {"access_token": "n", "refresh_token": "shared-refresh",
                          "expires_in": 3600, "scope": self._default_scope()},
    )
    results = auth_msal.refresh_all_accounts(tokens_dir=tmp_path, api_type="both")
    api_types = sorted(r.get("api_type") for r in results)
    assert api_types == ["graph", "outlook"]


def test_refresh_all_accounts_default_is_graph_only(tmp_path):
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=True)
    results = auth_msal.refresh_all_accounts(tokens_dir=tmp_path)
    assert len(results) == 1
    assert results[0].get("api_type", "graph") == "graph"
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_auth_msal_dual_token.py -k "outlook or both or graph_only" -v`
Expected: FAIL — `refresh_account`/`refresh_all_accounts` have no `api_type` kwarg.

- [ ] **Step 3: Write minimal implementation**

In `auth_msal.py`, extend `refresh_all_accounts` signature (`:735-739`) and probe construction. Add `api_type` param and an inner per-api worker; tag each result with `api_type`:

```python
def refresh_all_accounts(
    tokens_dir: Optional[Path] = None,
    client_id: Optional[str] = None,
    tenant_id: Optional[str] = None,
    api_type: str = "graph",
) -> list[dict[str, Any]]:
```

Inside the loop (replace the body that builds `probe` and appends results) iterate the requested api types:

```python
    api_types = ["graph", "outlook"] if api_type == "both" else [api_type]

    results: list[dict[str, Any]] = []

    for token_file in token_files:
        identifier = token_file.stem[: -len("_access_token")]
        if identifier.endswith("_outlook"):
            continue  # skip outlook sibling files during enumeration
        for current_api in api_types:
            entry = _refresh_one(
                identifier=identifier,
                tokens_dir=resolved_dir,
                client_id=client_id,
                tenant_id=tenant_id,
                api_type=current_api,
            )
            entry["api_type"] = current_api
            results.append(entry)

    return results
```

Extract the existing valid/refresh/failed logic into a module-level helper `_refresh_one` (reuse the current probe code from `:793-846`):

```python
def _refresh_one(
    identifier: str,
    tokens_dir: Path,
    client_id: Optional[str],
    tenant_id: Optional[str],
    api_type: str = "graph",
) -> dict[str, Any]:
    probe = MSALRefreshTokenAuth(
        tokens_dir=tokens_dir,
        client_id=client_id,
        tenant_id=tenant_id,
        account_identifier=identifier,
        api_type=api_type,
    )
    if probe._is_token_valid():
        token_data = probe._load_access_token_data() or {}
        return {"identifier": identifier, "status": "valid",
                "expires_at": token_data.get("expires_at"), "error": None}
    try:
        token_data = probe._do_refresh_locked()
        return {"identifier": identifier, "status": "refreshed",
                "expires_at": token_data.get("expires_at"), "error": None}
    except Exception as e:
        stale = probe._load_access_token_data() or {}
        logger.warning(f"_refresh_one: '{identifier}' ({api_type}) refresh failed: {e}")
        return {"identifier": identifier, "status": "failed",
                "expires_at": stale.get("expires_at"), "error": str(e)}
```

Update `refresh_account` (`:1019-1107`) to take `api_type` and delegate to `_refresh_one` after the missing-file check:

```python
def refresh_account(
    identifier: str,
    tokens_dir: Optional[Path] = None,
    client_id: Optional[str] = None,
    tenant_id: Optional[str] = None,
    api_type: str = "graph",
) -> dict[str, Any]:
    if not identifier or not identifier.strip():
        raise ValueError("identifier must be a non-empty string")
    resolved_dir = _resolve_tokens_dir(tokens_dir)
    token_file = resolved_dir / f"{identifier}_access_token.json"
    if not token_file.exists():
        return {"identifier": identifier, "status": "missing",
                "expires_at": None, "error": f"no token file at {token_file}"}
    result = _refresh_one(identifier, resolved_dir, client_id, tenant_id, api_type=api_type)
    result["api_type"] = api_type
    return result
```

> Note: `refresh_all_accounts` already tags entries with `api_type` in the loop; for the single-api default path the existing `tests/test_refresh_all_accounts.py` assertions on `status`/`expires_at`/`error` keys remain valid (new `api_type` key is additive).

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_auth_msal_dual_token.py -v`
Expected: PASS

Regression:
Run: `uv run pytest tests/test_refresh_all_accounts.py tests/test_auth_verify_and_refresh_account.py -v`
Expected: PASS (additive `api_type` key does not break existing key-based assertions)

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/auth_msal.py tests/test_auth_msal_dual_token.py
git commit -m "feat(auth): support api_type graph/outlook/both in refresh_all_accounts and refresh_account"
```

---

### Task 10: `--api graph|outlook|both` flag on `auth refresh`

**Files:**
- Modify: `src/microsoft_mcp/auth_cli.py` (`_cmd_refresh`, parser, `_print_refresh_results` label)
- Test: `tests/test_auth_cli.py`

- [ ] **Step 1: Write the failing test**

```python
# append to tests/test_auth_cli.py
def test_refresh_api_outlook_labels_output(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    monkeypatch.setattr(
        "microsoft_mcp.auth_msal.refresh_all_accounts",
        lambda **kw: [{"identifier": TEST_EMAIL, "status": "refreshed",
                       "expires_at": "2026-06-15T22:00:00Z", "error": None,
                       "api_type": kw.get("api_type", "graph")}],
    )
    rc = auth_cli.main(["refresh", "--api", "outlook"])
    out = capsys.readouterr().out
    assert rc == 0
    assert "Outlook: Refreshed" in out


def test_refresh_api_both_groups_by_label(monkeypatch, tmp_path, capsys):
    _msal_env(monkeypatch, tmp_path)
    monkeypatch.setattr(
        "microsoft_mcp.auth_msal.refresh_all_accounts",
        lambda **kw: [
            {"identifier": TEST_EMAIL, "status": "valid",
             "expires_at": "2026-06-15T22:00:00Z", "error": None, "api_type": "graph"},
            {"identifier": TEST_EMAIL, "status": "refreshed",
             "expires_at": "2026-06-15T22:30:00Z", "error": None, "api_type": "outlook"},
        ],
    )
    rc = auth_cli.main(["refresh", "--api", "both"])
    out = capsys.readouterr().out
    assert rc == 0
    assert "Graph: Valid" in out
    assert "Outlook: Refreshed" in out


def test_refresh_api_invalid_rejected(monkeypatch, tmp_path):
    _msal_env(monkeypatch, tmp_path)
    with pytest.raises(SystemExit):
        auth_cli.main(["refresh", "--api", "bogus"])
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_auth_cli.py -k api -v`
Expected: FAIL — `refresh` has no `--api` argument.

- [ ] **Step 3: Write minimal implementation**

Update the `refresh` parser in `_build_parser()`:

```python
    p_refresh.add_argument(
        "--api", choices=["graph", "outlook", "both"], default="graph",
        help="which access token(s) to refresh (default: graph)",
    )
```

Update `_cmd_refresh` to thread `api_type` and render per-api labels. Replace the `refresh_all_accounts`/`refresh_account` branches so they pass `api_type=args.api` and group output by `api_type`:

```python
def _api_label(api_type: str) -> str:
    return "Outlook" if api_type == "outlook" else "Graph"
```

In the single-email branch:

```python
        result = refresh_account(identifier=args.email, api_type=args.api, **env)
```

In the all-accounts branch:

```python
    results = refresh_all_accounts(api_type=args.api, **env)
    if args.json:
        _print_json(results)
    else:
        for r in results:
            _print_refresh_results([r], api_label=_api_label(r.get("api_type", "graph")))
    return 1 if any(r.get("status") == "failed" for r in results) else 0
```

> `--force` ignores `--api` (device-code flow always re-auths the graph identity, and the shared refresh token then mints outlook on demand). Keep the existing force branch unchanged.

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_auth_cli.py -v`
Expected: PASS (all)

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/auth_cli.py tests/test_auth_cli.py
git commit -m "feat(auth-cli): add --api graph|outlook|both to 'auth refresh'"
```

---

## WAVE 4 — Docs, Slash Commands, Regression

### Task 11: Update slash commands + docs

**Files:**
- Modify: `.claude/commands/auth-refresh.md`, `.claude/commands/auth-verify.md`
- Create: `.claude/commands/auth-status.md`
- Modify: `CLAUDE.md`, `README.md`

- [ ] **Step 1: Update `/auth-refresh` and `/auth-verify` to call the subcommand**

In `.claude/commands/auth-refresh.md`, change the dispatch line from `uv run auth_refresh.py $ARGUMENTS` to:

```
uv run microsoft-mcp auth refresh $ARGUMENTS
```

In `.claude/commands/auth-verify.md`, change to:

```
uv run microsoft-mcp auth verify $ARGUMENTS
```

- [ ] **Step 2: Create `.claude/commands/auth-status.md`**

```markdown
---
description: Read-only MSAL token health (no network) — per-account validity + expiry
allowed-tools: Bash(uv run:*)
argument-hint: "[--json]"
---

Show read-only token health for all saved MSAL accounts (does not refresh):

uv run microsoft-mcp auth status $ARGUMENTS
```

- [ ] **Step 3: Update `CLAUDE.md`**

Under the `tools.py` Account Management description and the MSAL "Known gotchas"/CLI section, add:

```markdown
- **`microsoft-mcp auth <cmd>` CLI** (MSAL only). Mirrors `outlook auth refresh`:
  - `auth refresh [email] [--api graph|outlook|both] [--force] [--json]`
  - `auth verify [--live] [--json]`, `auth status [--json]`, `auth list [--json]`,
    `auth test [--json]`, `auth doctor [--json]`
  - Dual exposure: `microsoft-mcp auth ...` (subcommand) and `microsoft-mcp-auth ...`
    (console script), mirroring the signatures CLI. Logic lives in `auth_cli.py`;
    `auth_refresh.py` is a thin backward-compat shim.
  - Color is zero-dependency ANSI, auto-disabled when stdout is not a TTY or
    `NO_COLOR` is set (`MICROSOFT_MCP_FORCE_COLOR=1` to force).

- **Dual Graph/Outlook tokens.** `MSALRefreshTokenAuth(api_type="outlook")` writes
  `{id}_outlook_access_token.json` using `outlook.office365.com/.default` and the
  SHARED `{id}_refresh_only.txt`. `auth refresh --api=both` mints both.
```

- [ ] **Step 4: Update `README.md`** with the same CLI usage block (user-facing examples showing the `✓ Graph: Valid, expires …` output).

- [ ] **Step 5: Commit**

```bash
git add .claude/commands/auth-refresh.md .claude/commands/auth-verify.md \
        .claude/commands/auth-status.md CLAUDE.md README.md
git commit -m "docs(auth): document microsoft-mcp auth CLI + dual-token support"
```

---

### Task 12: Entry-point + tool-surface regression tests

**Files:**
- Modify: `tests/test_server_entry.py`
- Modify: `tests/test_tool_surface_contract.py` (assert no MCP tool regressions; CLI adds no tools)
- Test: full suite

- [ ] **Step 1: Write the failing/guard test**

```python
# tests/test_server_entry.py — assert all three CLI dispatches coexist
import sys
from unittest import mock
import pytest


@pytest.mark.parametrize("head,module,fn", [
    ("signatures", "microsoft_mcp.signatures_cli", "main"),
    ("auth", "microsoft_mcp.auth_cli", "main"),
])
def test_subcommand_dispatch_routes(monkeypatch, head, module, fn):
    from microsoft_mcp import server

    called = {}
    monkeypatch.setattr(f"{module}.{fn}", lambda argv: called.setdefault("argv", argv) or 0)
    monkeypatch.setattr(sys, "argv", ["microsoft-mcp", head, "x", "--json"])
    with mock.patch.object(sys, "exit") as fake_exit:
        server.main()
    assert called["argv"] == ["x", "--json"]
    fake_exit.assert_called_once_with(0)
```

- [ ] **Step 2: Run it**

Run: `uv run pytest tests/test_server_entry.py -v`
Expected: PASS (both params)

- [ ] **Step 3: Run the entire suite + lint + types**

```bash
uv run pytest tests/ -v
uvx ruff format .
uvx ruff check --fix --unsafe-fixes .
uv run pyright src/microsoft_mcp/auth_cli.py src/microsoft_mcp/auth_msal.py
```
Expected: all green; ruff clean; pyright no new errors in the two modules.

- [ ] **Step 4: Manual end-to-end smoke (real tokens)**

```bash
MICROSOFT_MCP_AUTH_METHOD=msal uv run microsoft-mcp auth status
MICROSOFT_MCP_AUTH_METHOD=msal uv run microsoft-mcp auth refresh
MICROSOFT_MCP_AUTH_METHOD=msal uv run microsoft-mcp auth verify
```
Expected (matches the pasted target UX):
```
broach@cresa.com
  ✓ Graph: Valid, expires 2026-06-15 21:25:06 UTC
...
```

- [ ] **Step 5: Commit**

```bash
git add tests/test_server_entry.py tests/test_tool_surface_contract.py
git commit -m "test(auth): regression-guard CLI subcommand dispatch + tool surface"
```

---

## Self-Review

**Spec coverage:**
- Polished `✓ Graph: Valid, expires <UTC>` output → T1 (formatter) + T2 (`_print_refresh_results`). ✅
- `refresh` / `verify` / `status` / `list` / `test` / `doctor` → T2, T5, T6, T7. ✅
- Zero-dep ANSI, TTY-gated, `NO_COLOR` → T1. ✅
- Both exposures (`microsoft-mcp auth` + `microsoft-mcp-auth`) → T3. ✅
- `auth_refresh.py` → shim → T4. ✅
- Dual Graph/Outlook tokens, shared refresh, `--api` → T8, T9, T10. ✅
- Docs + slash commands + regression → T11, T12. ✅

**Placeholder scan:** No TBD/TODO; every code step shows full code. ✅

**Type/name consistency:** `_format_expiry`, `_c`, `_color_enabled`, `_env_kwargs`, `_print_refresh_results`, `_account_health`, `_enumerate_accounts`, `_refresh_one`, `_api_label`, `api_type`, `GRAPH_SCOPE`/`OUTLOOK_SCOPE`, `_default_scope` used consistently across tasks. ✅

**Single-account fixture policy:** every test uses `TEST_EMAIL = "broach@cresa.com"` only. ✅

---

## Rollback

Each task is one commit. To roll back the whole feature: `git revert` the range, or since all changes are additive/new-file except the `auth_msal.py` `api_type` thread:
- `auth_cli.py`, `tests/test_auth_cli.py`, `tests/test_auth_msal_dual_token.py`, `tests/test_auth_refresh_shim.py`, `.claude/commands/auth-status.md` — delete (new files).
- `server.py`, `pyproject.toml`, `auth_refresh.py`, `auth_msal.py`, docs — `git revert` their commits.
- `auth_msal.py` `api_type` defaults to `"graph"` everywhere, so reverting T10/T9 first then T8 leaves the engine in its original single-token behavior with no orphaned callers.

## Risks

1. **`auth_msal.py` regression (medium).** Threading `api_type` touches the hot refresh path. Mitigation: `api_type="graph"` default keeps every existing call-site byte-identical in behavior; T8/T9 run the existing `test_refresh_all_accounts.py` / `test_graph_401_retry.py` suites as regression gates.
2. **Outlook resource not consented (low/medium).** The Office public client is normally pre-consented for `outlook.office365.com`, but a refresh against the outlook scope can return AADSTS65001/70011 on some tenants. Mitigation: `--api=outlook` failures surface as `status: failed` per account (non-fatal), exactly like graph failures; graph refresh is unaffected.
3. **`auth_refresh.py` shim interface drift (low).** The legacy flag→subcommand translation (T4) is covered by `test_auth_refresh_shim.py`; the `/auth-refresh` slash command is repointed in T11.
