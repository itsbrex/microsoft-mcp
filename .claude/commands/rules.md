---
description: Manage Outlook inbox message rules (list/get/create/delete/toggle/export/import)
allowed-tools: ["Bash(uv run:*)"]
---

Run the inbox-rules CLI. Subcommands: `list`, `get <id>`, `create`, `delete <id>`,
`toggle <id>`, `export [--output rules.yaml]`, `import <rules.yaml>`. Add `--json`
for machine-readable output. Requires a valid auth session (see `/auth`).

Usage: `/rules list` → runs `microsoft-mcp rules list`. Pass through args after the command.

!`uv run microsoft-mcp rules $ARGUMENTS`
