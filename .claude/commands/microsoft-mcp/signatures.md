---
description: Manage the local plain-text signature store (list/show/set/rm/path)
allowed-tools: ["Bash(uv run:*)"]
argument-hint: "<list|show|set|edit|rm|path|dir> [name]"
---

Manage local email signatures stored at `~/.config/microsoft-mcp/signatures/` (override via `MICROSOFT_MCP_SIGNATURES_DIR`). Subcommands: `list`, `show`, `set` (`--from-file`, `--stdin`, `--editor`), `edit`, `rm`, `path`, `dir`.

Signatures are appended to drafts by `create_email_draft`/`update_email_draft`; account slug comes from `MICROSOFT_MCP_SIGNATURE_ACCOUNT` or `MICROSOFT_MCP_ACCOUNT_ID`.

Arguments passed: `$ARGUMENTS`

!`uv run microsoft-mcp signatures $ARGUMENTS`
