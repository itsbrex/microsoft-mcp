---
description: Generate Outlook intelligence reports (briefing/signals/contact/recap)
allowed-tools: ["Bash(uv run:*)"]
---

Run the intel CLI. Subcommands: `briefing [--timezone TZ] [--limit N]`,
`signals [--timezone TZ] [--level all|critical|important|informational]`,
`contact <email> [--days N]`, `recap [--timezone TZ]`. Add `--json` for
machine-readable output. Makes live Graph calls — requires a valid auth session (see `/auth`).

Usage: `/intel briefing --json`. Pass through args after the command.

!`uv run microsoft-mcp intel $ARGUMENTS`
