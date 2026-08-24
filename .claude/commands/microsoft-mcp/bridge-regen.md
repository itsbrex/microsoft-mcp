---
description: Regenerate the UTCP code-mode bridge config from the live Claude Desktop config
argument-hint: "[--include-server name ...] [--exclude-server name ...]"
allowed-tools: ["Bash(PYTHONPATH=src python -m microsoft_mcp.utcp_bridge_config:*)"]
---

Wrap the live Claude Desktop `mcpServers` config into a UTCP bridge configuration. Non-destructive — only reads the source, writes to `./tmp/claude-desktop-utcp-review-$(date +%Y%m%d)/`.

Existing `code-mode-mcp` entries are skipped by default (self-recursion guard). Pass `--include-server code-mode-mcp` if you want to wrap a nested bridge.

!`PYTHONPATH=src python -m microsoft_mcp.utcp_bridge_config "$HOME/Library/Application Support/Claude/claude_desktop_config.json" --output-dir "./tmp/claude-desktop-utcp-review-$(date +%Y%m%d)" $ARGUMENTS`

Outputs:
- `.utcp_config.json` — manuals catalog (one per wrapped server)
- `claude_desktop_config.utcp.json` — replacement Claude config with a single `code-mode-mcp` entry
- `manual_map.json` — source → sanitized-manual-name mapping

Review the three files, then manually swap the active Claude Desktop config if you want to activate. Never overwrites the live file.
