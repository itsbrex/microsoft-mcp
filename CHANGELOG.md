# Changelog

All notable changes to this project are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed

- Migrated the server and test compatibility layer to FastMCP 3 while
  preserving the existing 109-tool hybrid surface and five-tool code-only
  surface. The advertised server identity is now consistently `microsoft-mcp`.
- MSAL refresh handling now stores Microsoft's latest replacement refresh
  token for both Graph and Outlook refreshes.

### Fixed

- Noninteractive refresh failures preserve cached credentials instead of
  clearing them before returning an actionable error.
- Forced reauthentication rejects and removes credentials when Azure reports a
  different account, and reports partial failure when a requested Outlook token
  cannot be minted.
- Graph 401 recovery surfaces the authentication refresh failure instead of
  hiding it behind the original HTTP 401 response.
- Test startup isolates all token/cache paths and blocks authentication or live
  verification for `cresa.com`; auth fixtures use `cresa.email` instead.
- Token files are replaced atomically, and refresh-token availability is
  checked under the refresh lock to prevent concurrent callers from observing
  truncated credentials and entering interactive authentication.
- `AADSTS65002` guidance now identifies first-party client/resource
  preauthorization failure instead of incorrectly treating refresh tokens as
  resource-scoped.

## [0.2.0] - 2026-06-18

Large "outlook-creds mail port" — ~40 new MCP tools, 3 new CLIs, and a
signal-intelligence engine. Microsoft Graph REST only (no EWS/SOAP, no `lxml`).

### Added

- **Inbox rules** — `list_inbox_rules`, `get_inbox_rule`, `create_inbox_rule`,
  `update_inbox_rule`, `delete_inbox_rule`, `toggle_inbox_rule`,
  `reorder_inbox_rules`, plus YAML `export_inbox_rules` / `import_inbox_rules`.
  New `rules.py` (Graph-free payload/template builders + validation) and a
  `microsoft-mcp rules` CLI (`rules_cli.py`).
- **Focused Inbox overrides** — `list_focused_overrides`,
  `create_focused_override`, `update_focused_override`, `delete_focused_override`.
- **Reply / forward as drafts** — `reply_email_draft`, `reply_all_email_draft`,
  `forward_email_draft`. Draft-first: `send_email_draft` is the only tool that
  sends to the wire.
- **MailTips** — `get_mailtips` (out-of-office, mailbox-full, external-recipient
  warnings before sending).
- **Attachments** — `list_attachments`, `download_attachments` (path-traversal
  hardened: basename-only, skips `""`/`.`/`..`).
- **Microsoft To-Do** — list/task CRUD (`list_todo_lists`, `create_todo_list`,
  `list_tasks`, `create_task`, `update_task`, `complete_task`, `delete_task`),
  checklist items, and `create_task_from_email`. New `todo.py` helpers
  (due-date parsing + payload builders, injected clock).
- **Email templates** — YAML template engine (`templates_engine.py` +
  `templates_data/`) with HTML-escaped (XSS-safe) variable substitution.
  Tools: `list_email_templates`, `render_email_template`,
  `find_template_variables`, `get_template_placeholders`,
  `substitute_template_variables`. New `MICROSOFT_MCP_TEMPLATES_DIR`.
- **Signature parsing** — `parse_email_signature` (signature/OOO contact +
  job-change extraction) and `normalize_phone_number` (E.164). New
  `signature_parser.py`.
- **Intel reporting engine** — `intel/` package (collectors → analyzers →
  engine) producing `generate_morning_briefing`, `get_priority_signals`,
  `get_contact_intelligence`, `get_end_of_day_recap`, plus a `microsoft-mcp
  intel` CLI (`briefing`/`signals`/`contact`/`recap`).
- **Bounce detection** — `bounces.py` (NDR classifier, DSN parser, pattern
  catalogs), `scan_bounces` tool (folder scan → `{count, reasons, rows}` + CSV),
  and a `microsoft-mcp bounces` CLI (`scan`/`patterns`).
- **New console scripts** — `microsoft-mcp-rules`, `microsoft-mcp-intel`,
  `microsoft-mcp-bounces`, each also dispatched as `microsoft-mcp <name> …`
  (routed in `server.py` before the Graph stack is imported).
- **Dependency** — `pyyaml>=6.0,<7`.
- **Tests** — `tests/test_tool_surface.py` guards that all mail-port tools stay
  registered. Full suite is 960 tests.

### Security

- Capped signature-parse body length (`_MAX_BODY_LEN = 20000`) to prevent a
  ReDoS / CPU-DoS reachable via `parse_email_signature`.
- Path-traversal hardening in `download_attachments`.

### Fixed

- Intel calendar collector now converts Graph UTC datetimes to the target
  timezone (the `Prefer: outlook.timezone` header could not be sent because
  `graph.request` has no `headers` kwarg).
- Intel collector fetches and `bounces.iter_folder_messages` follow
  `@odata.nextLink` instead of silently truncating at one page.
- `intel` sent-count metric paginates instead of capping at `$top`.

### Design notes

- Pure modules (`intel/`, `bounces.py`, `todo.py`) take a dependency-injected
  `request` callable (real call site passes `graph.request`) and an injected
  `now`/`today` — they never `import graph` or call `datetime.now()`, keeping
  them unit-testable.

## [0.1.0]

Initial Microsoft 365 MCP server: Outlook mail, Calendar, OneDrive, Contacts,
Teams, and unified search over Microsoft Graph; dual Azure-SDK / MSAL auth with
multi-account support; local plain-text signature store; `microsoft-mcp auth`
and `microsoft-mcp signatures` CLIs.
