---
description: List and inspect YAML email/calendar templates (built-in + user dir)
argument-hint: "[category/name]"
---

# Email templates

Inspect the YAML template system (`templates_engine.py` + bundled `templates_data/`, user overrides in `~/.config/microsoft-mcp/templates/` or `$MICROSOFT_MCP_TEMPLATES_DIR`).

Argument: `$ARGUMENTS` (optional `category/name` to inspect one template).

## Steps

1. List all templates:
   `uv run python -c "from microsoft_mcp.templates_engine import list_templates; [print(t['category'] + '/' + t['name']) for t in list_templates()]"`
2. If an argument was given, load that template and show its placeholders (`required`, `default`) and `conditional_sections`:
   `uv run python -c "from microsoft_mcp.templates_engine import load_template; import json; print(json.dumps(load_template('<category>', '<name>'), indent=2, default=str))"`
3. Remind: user templates shadow built-ins by `(category, name)` key; all placeholder values are HTML-escaped on render.
