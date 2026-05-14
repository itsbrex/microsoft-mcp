# Local Plain-Text Signatures — Design

**Date:** 2026-05-13
**Status:** Approved (pending implementation plan)
**Author:** itsbrex / Claude

## Context and motivation

Microsoft Graph does not expose Outlook signature settings: there is no
`/me/signature` endpoint, signatures are not part of `mailboxSettings`, and
draft/send endpoints have no "apply default signature" flag. Microsoft Q&A and
the Graph mail API overview both confirm this — signatures are a client-side
concern stored per Outlook installation.

The supported workaround is to store signature content in our own system and
append it to message bodies when creating drafts. This spec defines a small,
local plain-text file store, a CLI for managing it, and an opt-in injection
path inside the existing draft tools.

## Goals

- Per-account, per-name plain-text signature files on disk.
- A first-class CLI (both a standalone console script and a subcommand on the
  existing `microsoft-mcp` entry point) to create, read, update, delete, and
  inspect signatures.
- Read-only MCP tools so the assistant can discover and quote signatures
  without being able to rewrite them.
- Opt-in signature injection on the existing `create_email_draft` and
  `update_email_draft` tools, with environment-variable defaults so a server
  can be configured to always sign new mail / replies.
- Optional HTML sibling files for fidelity in HTML drafts, with a plain-text
  fallback that auto-converts.

## Non-goals (v1)

- No WYSIWYG / HTML editor.
- No team-shared or remote signature stores; no templating (`{{name}}`).
- No write-capable MCP tools — assistants cannot mutate signature content.
- No "global" (account-less) fallback signatures. Signatures are always
  account-scoped.
- No automatic embedded-image signatures (cid: attachments).

## File store

- **Directory:** `~/.config/microsoft-mcp/signatures/` by default;
  overridable via `MICROSOFT_MCP_SIGNATURES_DIR`.
- **Naming:** `<account-slug>-<signature-name>.txt` (e.g.,
  `brian-work-default.txt`, `jp-work-replies.txt`).
- **HTML sibling (optional):** `<account-slug>-<signature-name>.html` — when
  present, used verbatim for HTML drafts; ignored for text drafts.
- **Account slug resolution:**
  1. `MICROSOFT_MCP_SIGNATURE_ACCOUNT` if set (used verbatim, lowercased).
  2. Else slugify `MICROSOFT_MCP_ACCOUNT_ID`: lowercase, replace `@` and `.`
     with `-`, strip any character outside `[a-z0-9-]`, collapse repeated
     `-`, trim leading/trailing `-`.
  3. If both are unset and no `account` override is passed to the API, raise
     `ValueError("no signature account; set MICROSOFT_MCP_SIGNATURE_ACCOUNT, MICROSOFT_MCP_ACCOUNT_ID, or pass account=")`.
- **Signature-name constraints:** must match `[A-Za-z0-9._-]+`. Names are
  lowercased at every API entry point (`write_signature`, `read_signature`,
  `apply_signature`, CLI args, MCP tool args), so the on-disk filename is
  always lowercase. This avoids the ambiguity of two files differing only in
  case on case-insensitive filesystems like APFS.
- **Sentinel:** the name `"none"` (case-insensitive) means "do not inject" and
  is never written to disk.

## Module: `src/microsoft_mcp/signatures.py`

Pure helpers — no MCP, no Graph. The only side effects are reads/writes inside
the resolved signatures directory.

Public surface:

```python
def resolve_dir() -> Path
def account_slug(account_id: str | None = None, override: str | None = None) -> str
def signature_path(name: str, *, account: str | None = None, html: bool = False) -> Path

@dataclass(frozen=True)
class SignatureInfo:
    account: str
    name: str
    path: Path
    has_html: bool
    size: int
    modified: float  # POSIX mtime

def list_signatures(account: str | None = None) -> list[SignatureInfo]
def read_signature(name: str, *, account: str | None = None, html: bool = False) -> str | None
def write_signature(name: str, content: str, *, account: str | None = None, html: bool = False) -> Path
def delete_signature(name: str, *, account: str | None = None, html: bool = False) -> bool

def apply_signature(
    body: str | None,
    body_content_type: str,
    name: str,
    *,
    account: str | None = None,
) -> str
```

`apply_signature` rules:

- `name == "none"` (case-insensitive) → returns `body` unchanged (or `""` if
  `body is None`). No file access.
- If `body_content_type == "html"`:
  - If `<account>-<name>.html` exists → append its raw contents.
  - Else read `<account>-<name>.txt`, escape HTML-significant characters,
    convert `\n` to `<br>\n`, wrap in `<div class="signature">…</div>`,
    append.
- If `body_content_type == "text"`:
  - Read `<account>-<name>.txt` and append as plain text.
  - `.html` siblings are ignored.
- **Separator:**
  - Default: `"\n\n"`.
  - If `MICROSOFT_MCP_SIGNATURE_RFC3676` is truthy (`1`, `true`, `yes`):
    `"\n\n-- \n"` for text bodies; `"<br><br>-- <br>\n"` for HTML bodies.
- If the requested signature file does not exist:
  - In `apply_signature`, raise `SignatureNotFoundError(name, account, html)`
    so the caller can decide policy (the draft tools convert this to a
    structured warning; the CLI surfaces it as a non-zero exit + stderr).
- If `body` is `None` or empty, the signature becomes the entire body (no
  leading separator).

## CLI: `src/microsoft_mcp/signatures_cli.py`

Argparse-based. Exposed two ways (see "Entry points" below); both invocations
share the same `main(argv: list[str] | None = None) -> int` function.

Subcommands:

```
microsoft-mcp-signatures list  [--account NAME] [--json]
microsoft-mcp-signatures show  NAME [--account NAME] [--html]
microsoft-mcp-signatures set   NAME [--account NAME] [--html]
                               [--from-file PATH | --stdin | --editor]
microsoft-mcp-signatures edit  NAME [--account NAME] [--html]
microsoft-mcp-signatures rm    NAME [--account NAME] [--html] [--yes]
microsoft-mcp-signatures path  [NAME] [--account NAME] [--html]
microsoft-mcp-signatures dir
```

Behavior:

- `list`: enumerates signatures for the resolved account (or all accounts if
  `--account` is passed as `*` / `all`). Human-readable table by default,
  JSON array when `--json` is passed.
- `show`: prints the file contents to stdout. Honors `--html` to request the
  sibling file.
- `set`: content source priority is `--from-file` > `--stdin` > `--editor`.
  If none given and stdin is a TTY, behaves like `--editor`; otherwise reads
  stdin. `--editor` uses `$VISUAL`, then `$EDITOR`, then `vi`. The directory
  is created on demand (`mkdir(parents=True, exist_ok=True)`).
- `edit`: shorthand for `set --editor` that pre-loads existing content.
- `rm`: prompts unless `--yes` is given. Returns non-zero if the file does
  not exist.
- `path`: prints the resolved file path (does not require the file to exist).
  With no `NAME`, prints the directory (alias for `dir`).
- `dir`: prints the resolved signatures directory.

Exit codes: `0` on success, `1` on user errors (bad arg, missing file when
required), `2` on argparse usage errors.

## Entry points and subcommand bridge

`pyproject.toml` adds:

```
[project.scripts]
microsoft-mcp = "microsoft_mcp.server:main"
microsoft-mcp-signatures = "microsoft_mcp.signatures_cli:cli_main"
```

`signatures_cli.cli_main()` is a thin wrapper around `main()` that calls
`sys.exit(main(sys.argv[1:]))`.

`server.main()` grows a small dispatch at its top:

```python
def main():
    argv = sys.argv[1:]
    if argv and argv[0] == "signatures":
        from microsoft_mcp import signatures_cli
        sys.exit(signatures_cli.main(argv[1:]))
    # …existing server start path…
```

Net effect: `microsoft-mcp signatures list` and
`microsoft-mcp-signatures list` are equivalent. All other server invocation
forms remain unchanged.

## MCP tools (read-only)

Added to `tools.py`. Registered unconditionally; signatures are local and
auth-method-agnostic.

```python
@mcp.tool
def list_signatures(account: str | None = None) -> list[dict[str, Any]]:
    """List available signatures for the current (or specified) account."""

@mcp.tool
def get_signature(
    name: str,
    account: str | None = None,
    html: bool = False,
) -> dict[str, Any]:
    """Return the resolved signature content (or a structured 'not_found')."""
```

`get_signature` returns:

```python
{"status": "ok", "account": "...", "name": "...", "html": bool,
 "path": "...", "content": "..."}
# or
{"status": "not_found", "account": "...", "name": "...", "html": bool,
 "path": "..."}
```

`get_signature` does not raise on missing files — the model gets a
recoverable signal. Other I/O errors (permission denied, etc.) still raise.

No write/delete tools are exposed. The CLI is the only mutation path.

## Draft-tool integration

Both `create_email_draft` and `update_email_draft` grow one new parameter:

```python
signature: str | None = None,  # signature name; "none" suppresses; None → env default
```

Resolution algorithm (executed inside each tool, before the Graph payload is
constructed):

1. If `signature` is a non-empty string and equals `"none"` (case-insensitive)
   → no injection.
2. Else if `signature` is a non-empty string → use that name verbatim.
3. Else (`signature is None`):
   - `create_email_draft(draft_type="new")` → `MICROSOFT_MCP_DEFAULT_SIGNATURE`.
   - `create_email_draft(draft_type in {"reply", "reply_all"})` →
     `MICROSOFT_MCP_REPLY_SIGNATURE`, falling back to
     `MICROSOFT_MCP_DEFAULT_SIGNATURE` when the reply var is unset.
   - `update_email_draft` → `MICROSOFT_MCP_DEFAULT_SIGNATURE`, **only when
     `body` is being set**. A pure metadata update (e.g., changing
     recipients) never re-applies a signature.
4. If no name is resolved → no injection.
5. With a resolved name, call `signatures.apply_signature(body, content_type,
   name)`:
   - On success, the returned string replaces `body` in the payload.
   - On `SignatureNotFoundError`, the tool result includes
     `"signature_warning": "signature not found: <account>/<name>"` and the
     draft is still created/updated without the signature. Never hard-fails.
6. The applied body is the `body` field that goes into the Graph payload.
   For `reply`/`reply_all` drafts, Graph generates the quoted history when
   the draft is created; our signature lives in the user-supplied `body`
   that we patch on top, so the rendered draft has user text + signature
   above the quoted history (standard top-posting).

The successful tool result includes one new optional field:

```python
"signature_applied": {"account": "...", "name": "...", "html": bool}
```

omitted when no signature was injected.

## Environment variables (new)

| Name                                 | Purpose                                                       |
| ------------------------------------ | ------------------------------------------------------------- |
| `MICROSOFT_MCP_SIGNATURES_DIR`       | Override signatures directory (default `~/.config/microsoft-mcp/signatures/`). |
| `MICROSOFT_MCP_SIGNATURE_ACCOUNT`    | Override the account slug used for filename lookup.           |
| `MICROSOFT_MCP_DEFAULT_SIGNATURE`    | Implicit signature name for new drafts.                       |
| `MICROSOFT_MCP_REPLY_SIGNATURE`      | Implicit signature name for reply / reply_all drafts.         |
| `MICROSOFT_MCP_SIGNATURE_RFC3676`    | `1`/`true` to use the RFC 3676 `-- ` delimiter (default `\n\n`). |

## Testing

- `tests/test_signatures.py`
  - `account_slug` for `MICROSOFT_MCP_SIGNATURE_ACCOUNT`, slug-from-email,
    explicit override, and the no-account error.
  - `signature_path` shape for `.txt` and `.html`.
  - `write` → `read` → `list` → `delete` round-trips.
  - `apply_signature` for text bodies, HTML bodies with and without `.html`
    sibling, with and without `MICROSOFT_MCP_SIGNATURE_RFC3676`.
  - `apply_signature` with `name == "none"` returns body unchanged and does
    no file I/O.
  - `SignatureNotFoundError` raised for missing files.
- `tests/test_signatures_cli.py`
  - Each subcommand via `signatures_cli.main([...])` with `tmp_path` and
    monkeypatched env. Editor path stubbed via `EDITOR=true`.
- `tests/test_draft_signatures.py`
  - `signature="name"` overrides env default.
  - Env default applied when `signature` is `None`.
  - `signature="none"` suppresses injection (env default ignored).
  - Missing signature surfaces `signature_warning` and still returns
    `status="draft_created"`.
  - `update_email_draft` does not apply a signature when `body` is not
    supplied.
  - HTML body uses `.html` sibling when present, else converts from `.txt`.

All file I/O is routed through a `tmp_path`-backed
`MICROSOFT_MCP_SIGNATURES_DIR` fixture.

## Docs

- `CLAUDE.md` — add a "Signatures" subsection under Architecture covering the
  file-store layout, env vars, injection rules, and the two read-only MCP
  tools.
- `README.md` — short "Email signatures" section with the CLI cheat sheet.
- `env.example` — add the five new env vars, commented out, with one-line
  comments.

## Risk and rollback

- **Scope is additive.** Every new field has a safe default and existing
  callers continue to work unchanged.
- **No new dependencies.** Argparse and pathlib are stdlib; the rest is
  built on the existing `graph.request` / FastMCP plumbing.
- **Rollback** is a single revert of the feature branch; no data migration
  is involved.

## Open considerations (deliberately deferred)

- Embedded-image signatures (vCard logos, etc.) — requires multipart
  attachments and `cid:` references; defer until a real ask comes in.
- Per-recipient or per-conversation signature selection — out of scope until
  we see usage data.
- Sharing signatures across machines — once the file format stabilizes,
  syncing via the user's existing dotfiles or `git`-managed config dir is
  trivial.
