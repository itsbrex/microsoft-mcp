---
description: Scan a mail folder for bounce/NDR messages and export results
allowed-tools: ["Bash(uv run:*)"]
---

Run the bounces CLI. Subcommands: `scan [--folder inbox] [--limit 200] [--output bounces.csv]`
and `patterns` (print the bounce-detection pattern catalogs, no Graph call). Add `--json`
for machine-readable output. `scan` makes live Graph calls — requires a valid auth session (see `/auth`).

Usage: `/bounces scan --folder inbox --limit 200 --output bounces.csv`. Pass through args after the command.

!`uv run microsoft-mcp bounces $ARGUMENTS`
