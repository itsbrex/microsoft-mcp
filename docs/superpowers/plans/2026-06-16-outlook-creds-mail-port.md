# Outlook-Creds Mail Management Port — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Port the mail-rule management, draft/template, To-Do, signature-parsing, intel, and bounce-detection capabilities from `/Users/hack/github/outlook-creds` into `microsoft-mcp` as native `@mcp.tool` functions and CLI sub-apps, using Microsoft Graph REST only.

**Architecture:** Each capability becomes pure, unit-testable helper modules under `src/microsoft_mcp/` plus thin `@mcp.tool` wrappers in `tools.py` that call `graph.request`/`graph.request_paginated` and pass results through `response_shaping`. Three new CLI sub-apps (`rules`, `intel`, `bounces`) mirror the existing `auth_cli`/`signatures_cli` pattern (standalone console script + `microsoft-mcp <name>` subcommand dispatched in `server.main()`). The source repo's EWS SOAP path is deliberately **not** ported — Graph `messageRules` is strictly richer (30+ predicates vs ~10) and avoids the `lxml` dependency and the EWS-vs-Graph folder-ID mismatch.

**Tech Stack:** Python 3.13, FastMCP, httpx (via `graph.request`), MSAL/Azure auth, `pyyaml` (new dependency — for rule + draft templates), pytest + `unittest.mock`, ruff, pyright.

## Global Constraints

- **Graph REST only.** No EWS SOAP, no `lxml`. Every API call goes through `microsoft_mcp.graph.request` / `graph.request_paginated` so retry, pagination, 401-recovery, and auth headers apply. Copied verbatim from CLAUDE.md: "all Graph calls must go through microsoft_mcp.graph.request".
- **Draft-first.** The server never auto-sends mail. `reply`/`reply_all`/`forward` create **drafts** (`/createReply`, `/createReplyAll`, `/createForward`). Actual sending is a single explicit `send_email_draft(draft_id)` tool the user must call deliberately.
- **Response shaping.** Every list/detail tool accepts `response_profile: str = "auto"`, resolves it via the existing `get_response_profile()` helper (tools.py), and returns `cleanup_graph_payload(...)`-cleaned output for `legacy` and hand-shaped compact dicts for `assistant`.
- **Logging.** Every tool logs entry with key params (`logger.info`) and errors with `logger.error(..., exc_info=True)` before re-raising, matching existing tools.
- **Tests.** Mock `src.microsoft_mcp.tools.graph.request` (and `graph.request_paginated`). Token-file tests use a single account canonicalized to `broach@cresa.com` (declared `TEST_EMAIL` per module) per CLAUDE.md single-account fixture policy.
- **Auto-format.** A PostToolUse hook runs `ruff format` on `Write|Edit`. Still run `uvx ruff check --fix --unsafe-fixes .` and `uv run pyright` before finishing each wave.
- **Dependency floor.** New runtime deps use bounded versions: `pyyaml>=6.0,<7`.
- **Tool-surface regression.** `tests/test_server_entry.py` (or equivalent tool-count test) MUST be updated whenever the tool count changes, and CLI dispatch additions get regression coverage like the existing `auth`/`signatures` dispatch tests.

## Convention: "pattern shown once"

Several tool families are near-identical CRUD wrappers (rules, focused overrides, To-Do, checklist). To stay DRY this plan shows **complete, copyable code for the first tool of each family** and then specifies each sibling as an exact contract: endpoint, HTTP method, request payload, return shape, and the one test that proves it. A sibling's implementation is the canonical code with only those deltas applied. This is intentional and overrides the usual "repeat the code" rule for mechanically identical siblings only — anything with non-trivial logic gets full code.

---

## File Structure

**New modules:**
- `src/microsoft_mcp/rules.py` — pure helpers: rule TypedDicts, condition/action summarizers, Graph-payload builders, YAML-template ⇄ Graph converters, template validation.
- `src/microsoft_mcp/rules_cli.py` — `microsoft-mcp rules` sub-app (list/get/export/import/create/delete/toggle).
- `src/microsoft_mcp/todo.py` — pure helpers: due-date parsing, task payload builder, linked-resource builder.
- `src/microsoft_mcp/templates_engine.py` — YAML template loader/renderer (placeholders + conditional sections), `{{var}}` substitution, CSV recipient parsing.
- `src/microsoft_mcp/templates_data/calendar/*.yaml`, `templates_data/email/*.yaml`, `templates_data/*/_schema.yaml` — bundled built-in templates.
- `src/microsoft_mcp/signature_parser.py` — signature/OOO contact extraction, phone E.164 normalization, job-change detection.
- `src/microsoft_mcp/intel/__init__.py`, `intel/types.py`, `intel/_utils.py`, `intel/collectors/{email,calendar,threads,contacts}.py`, `intel/analyzers/{priority,schedule,relationships}.py`, `intel/engine.py` — collector→analyzer→engine pipeline.
- `src/microsoft_mcp/intel_cli.py` — `microsoft-mcp intel` sub-app (briefing/signals/contact/recap).
- `src/microsoft_mcp/bounces.py` — NDR pattern catalog, classifier, DSN parser, folder scan, CSV writer.
- `src/microsoft_mcp/bounces_cli.py` — `microsoft-mcp bounces` sub-app (scan/export/patterns).

**Modified:**
- `src/microsoft_mcp/tools.py` — add ~30 `@mcp.tool` functions (rules, focused, reply/forward/send, mailtips, attachments, To-Do, templates, signature, intel) + gating for any scope-restricted tools.
- `src/microsoft_mcp/server.py:73-87` — add `rules` / `intel` / `bounces` argv dispatch.
- `pyproject.toml:10-26` — add `pyyaml` dep + 3 console-script entries + package-data for `templates_data`.
- `CLAUDE.md`, `IMPLEMENTATION.md`, `README.md` — document new tools, CLIs, env vars.
- `tests/test_server_entry.py` — updated tool-count + CLI-dispatch regression.

**Tests:** one `tests/test_<module>.py` per new module, plus `tests/test_tools_<family>.py` for the tool wrappers.

---

## Wave dependency graph

```
WAVE 1  Foundation (pyyaml dep + rules.py helpers)        ─┐
WAVE 2  Inbox-rule MCP tools          (needs W1)           │
WAVE 3  Rule templates + rules CLI    (needs W1, W2)       │
WAVE 4  Focused Inbox overrides       (needs base only) ───┤ parallel after W1
WAVE 5  Reply/forward/send + mailtips + attachments (base) │ parallel
WAVE 6  Microsoft To-Do               (base)               │ parallel
WAVE 7  Template engine               (needs W1 pyyaml)    │ parallel after W1
WAVE 8  Signature parser + Intel      (base)               │ parallel
WAVE 9  Bounce detection + CLI        (base)               ─┘ parallel
WAVE 10 Integration: docs + regression tests  (needs ALL)
```
"base" = the existing codebase only. Waves 4–9 can be built concurrently once Wave 1 lands (Wave 7 needs the `pyyaml` dep from Wave 1 Task 1.1; everything else in 4/5/6/8/9 needs nothing from Wave 1).

---

# WAVE 1 — Foundation & rule helpers

### Task 1.1: Add `pyyaml` dependency + package data

**Depends:** none
**Files:**
- Modify: `pyproject.toml:10-20` (dependencies), `pyproject.toml` (add `[tool.hatch.build]` package-data for templates if hatchling is the backend; otherwise setuptools `package-data`).

- [ ] **Step 1:** Add `"pyyaml>=6.0,<7",` to the `dependencies` list in `pyproject.toml` (after the existing entries, before the closing `]` at line ~20).

- [ ] **Step 2:** Ensure bundled YAML ships in the wheel. Under the build-backend's package config, add `templates_data/**/*.yaml` to included package data (mirror however the repo currently ships non-`.py` files; if none exists, add `[tool.hatch.build.targets.wheel] include = ["src/microsoft_mcp/templates_data/**/*.yaml"]` or the setuptools equivalent).

- [ ] **Step 3:** Sync deps.

Run: `uv sync`
Expected: resolves and installs `pyyaml`.

- [ ] **Step 4:** Verify import works.

Run: `uv run python -c "import yaml; print(yaml.__version__)"`
Expected: prints a 6.x version.

- [ ] **Step 5: Commit**

```bash
git add pyproject.toml uv.lock
git commit -m "build: add pyyaml dependency for rule + draft templates"
```

---

### Task 1.2: Rule condition/action summarizers (pure helpers)

**Depends:** none (can start immediately)
**Files:**
- Create: `src/microsoft_mcp/rules.py`
- Test: `tests/test_rules_helpers.py`

**Interfaces:**
- Produces:
  - `RULE_LIST_FIELDS: str` = `"id,displayName,sequence,isEnabled,conditions,actions"`
  - `RULE_DETAIL_FIELDS: str` = `"id,displayName,sequence,isEnabled,hasError,isReadOnly,conditions,actions,exceptions"`
  - `summarize_conditions(conditions: dict[str, Any] | None) -> str`
  - `summarize_actions(actions: dict[str, Any] | None) -> str`
  - `shape_rule_summary(rule: dict[str, Any]) -> dict[str, Any]` → `{id, display_name, sequence, is_enabled, conditions_summary, actions_summary}`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_rules_helpers.py
from microsoft_mcp import rules


def test_summarize_conditions_human_readable():
    conds = {"senderContains": ["acme.com"], "subjectContains": ["invoice"], "hasAttachments": True}
    out = rules.summarize_conditions(conds)
    assert "acme.com" in out and "invoice" in out and "attachment" in out.lower()


def test_summarize_actions_human_readable():
    acts = {"moveToFolder": "AAMk123", "markAsRead": True, "stopProcessingRules": True}
    out = rules.summarize_actions(acts)
    assert "move" in out.lower() and "read" in out.lower() and "stop" in out.lower()


def test_shape_rule_summary_keys():
    rule = {
        "id": "r1", "displayName": "Newsletters", "sequence": 3, "isEnabled": True,
        "conditions": {"senderContains": ["news"]}, "actions": {"markAsRead": True},
    }
    s = rules.shape_rule_summary(rule)
    assert s == {
        "id": "r1", "display_name": "Newsletters", "sequence": 3, "is_enabled": True,
        "conditions_summary": s["conditions_summary"], "actions_summary": s["actions_summary"],
    }
    assert "news" in s["conditions_summary"]
```

- [ ] **Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_rules_helpers.py -v`
Expected: FAIL — `ModuleNotFoundError: microsoft_mcp.rules` / attribute errors.

- [ ] **Step 3: Write minimal implementation**

```python
# src/microsoft_mcp/rules.py
"""Pure, Graph-free helpers for Outlook inbox message rules.

All network I/O lives in tools.py; this module only builds payloads,
summarizes rules for display, and converts to/from YAML templates.
"""
from __future__ import annotations

from typing import Any

RULE_LIST_FIELDS = "id,displayName,sequence,isEnabled,conditions,actions"
RULE_DETAIL_FIELDS = (
    "id,displayName,sequence,isEnabled,hasError,isReadOnly,"
    "conditions,actions,exceptions"
)

# Human-readable labels for the boolean/scalar predicate keys.
_CONDITION_LABELS: dict[str, str] = {
    "hasAttachments": "has attachments",
    "isApprovalRequest": "is approval request",
    "isAutomaticForward": "is auto-forward",
    "isAutomaticReply": "is auto-reply",
    "isEncrypted": "is encrypted",
    "isMeetingRequest": "is meeting request",
    "isMeetingResponse": "is meeting response",
    "isNonDeliveryReport": "is NDR/bounce",
    "isPermissionControlled": "is permission-controlled",
    "isReadReceipt": "is read receipt",
    "isSigned": "is signed",
    "isVoicemail": "is voicemail",
    "sentToMe": "sent to me",
    "sentCcMe": "cc's me",
    "sentOnlyToMe": "sent only to me",
    "sentToOrCcMe": "sent to or cc's me",
    "notSentToMe": "not sent to me",
}
_LIST_CONDITION_LABELS: dict[str, str] = {
    "bodyContains": "body contains",
    "bodyOrSubjectContains": "body/subject contains",
    "headerContains": "header contains",
    "subjectContains": "subject contains",
    "senderContains": "sender contains",
    "recipientContains": "recipient contains",
    "categories": "categorized",
}


def _addrs(recips: list[dict[str, Any]]) -> list[str]:
    out = []
    for r in recips or []:
        ea = r.get("emailAddress", {}) if isinstance(r, dict) else {}
        out.append(ea.get("address") or ea.get("name") or "")
    return [a for a in out if a]


def summarize_conditions(conditions: dict[str, Any] | None) -> str:
    if not conditions:
        return "(any message)"
    parts: list[str] = []
    for key, label in _LIST_CONDITION_LABELS.items():
        vals = conditions.get(key)
        if vals:
            parts.append(f"{label} {', '.join(vals)}")
    if conditions.get("fromAddresses"):
        parts.append(f"from {', '.join(_addrs(conditions['fromAddresses']))}")
    if conditions.get("sentToAddresses"):
        parts.append(f"sent to {', '.join(_addrs(conditions['sentToAddresses']))}")
    if conditions.get("importance"):
        parts.append(f"importance={conditions['importance']}")
    for key, label in _CONDITION_LABELS.items():
        if conditions.get(key):
            parts.append(label)
    return "; ".join(parts) if parts else "(any message)"


def summarize_actions(actions: dict[str, Any] | None) -> str:
    if not actions:
        return "(no actions)"
    parts: list[str] = []
    if actions.get("moveToFolder"):
        parts.append(f"move to {actions['moveToFolder']}")
    if actions.get("copyToFolder"):
        parts.append(f"copy to {actions['copyToFolder']}")
    if actions.get("assignCategories"):
        parts.append(f"categorize {', '.join(actions['assignCategories'])}")
    if actions.get("markImportance"):
        parts.append(f"mark importance {actions['markImportance']}")
    if actions.get("markAsRead"):
        parts.append("mark as read")
    if actions.get("forwardTo"):
        parts.append(f"forward to {', '.join(_addrs(actions['forwardTo']))}")
    if actions.get("redirectTo"):
        parts.append(f"redirect to {', '.join(_addrs(actions['redirectTo']))}")
    if actions.get("delete"):
        parts.append("delete")
    if actions.get("permanentDelete"):
        parts.append("permanently delete")
    if actions.get("stopProcessingRules"):
        parts.append("stop processing further rules")
    return "; ".join(parts) if parts else "(no actions)"


def shape_rule_summary(rule: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": rule.get("id"),
        "display_name": rule.get("displayName"),
        "sequence": rule.get("sequence"),
        "is_enabled": rule.get("isEnabled"),
        "conditions_summary": summarize_conditions(rule.get("conditions")),
        "actions_summary": summarize_actions(rule.get("actions")),
    }
```

- [ ] **Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_rules_helpers.py -v`
Expected: 3 passed.

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/rules.py tests/test_rules_helpers.py
git commit -m "feat(rules): add rule condition/action summarizers"
```

---

### Task 1.3: Rule payload builder (predicates + actions from kwargs)

**Depends:** 1.2
**Files:**
- Modify: `src/microsoft_mcp/rules.py`
- Test: `tests/test_rules_helpers.py`

**Interfaces:**
- Produces `build_rule_payload(...) -> dict[str, Any]` building a Graph `messageRule` body. Signature:
  ```python
  def build_rule_payload(
      *, display_name: str, sequence: int = 1, is_enabled: bool = True,
      sender_contains: list[str] | None = None,
      subject_contains: list[str] | None = None,
      body_contains: list[str] | None = None,
      from_addresses: list[str] | None = None,
      has_attachments: bool | None = None,
      importance: str | None = None,
      move_to_folder: str | None = None,
      copy_to_folder: str | None = None,
      assign_categories: list[str] | None = None,
      mark_as_read: bool | None = None,
      mark_importance: str | None = None,
      forward_to: list[str] | None = None,
      delete: bool | None = None,
      stop_processing_rules: bool | None = None,
  ) -> dict[str, Any]
  ```
  Emits `{"displayName","sequence","isEnabled","conditions":{...},"actions":{...}}`, omitting empty condition/action sub-keys. Email strings become `{"emailAddress": {"address": x}}` recipient objects.

- [ ] **Step 1: Write the failing test**

```python
def test_build_rule_payload_minimal_move():
    p = rules.build_rule_payload(
        display_name="News", sender_contains=["news.com"], move_to_folder="AAMkFolder",
        mark_as_read=True,
    )
    assert p["displayName"] == "News"
    assert p["sequence"] == 1 and p["isEnabled"] is True
    assert p["conditions"] == {"senderContains": ["news.com"]}
    assert p["actions"] == {"moveToFolder": "AAMkFolder", "markAsRead": True}


def test_build_rule_payload_addresses_and_forward():
    p = rules.build_rule_payload(
        display_name="Fwd", from_addresses=["a@x.com"], forward_to=["b@y.com"],
    )
    assert p["conditions"]["fromAddresses"] == [{"emailAddress": {"address": "a@x.com"}}]
    assert p["actions"]["forwardTo"] == [{"emailAddress": {"address": "b@y.com"}}]
```

- [ ] **Step 2: Run** `uv run pytest tests/test_rules_helpers.py -k build_rule_payload -v` → FAIL (no attribute).

- [ ] **Step 3: Implement** (append to `rules.py`):

```python
def _recipients(emails: list[str] | None) -> list[dict[str, Any]] | None:
    if not emails:
        return None
    return [{"emailAddress": {"address": e}} for e in emails]


def build_rule_payload(
    *,
    display_name: str,
    sequence: int = 1,
    is_enabled: bool = True,
    sender_contains: list[str] | None = None,
    subject_contains: list[str] | None = None,
    body_contains: list[str] | None = None,
    from_addresses: list[str] | None = None,
    has_attachments: bool | None = None,
    importance: str | None = None,
    move_to_folder: str | None = None,
    copy_to_folder: str | None = None,
    assign_categories: list[str] | None = None,
    mark_as_read: bool | None = None,
    mark_importance: str | None = None,
    forward_to: list[str] | None = None,
    delete: bool | None = None,
    stop_processing_rules: bool | None = None,
) -> dict[str, Any]:
    conditions: dict[str, Any] = {}
    if sender_contains:
        conditions["senderContains"] = sender_contains
    if subject_contains:
        conditions["subjectContains"] = subject_contains
    if body_contains:
        conditions["bodyContains"] = body_contains
    if from_addresses:
        conditions["fromAddresses"] = _recipients(from_addresses)
    if has_attachments is not None:
        conditions["hasAttachments"] = has_attachments
    if importance:
        conditions["importance"] = importance

    actions: dict[str, Any] = {}
    if move_to_folder:
        actions["moveToFolder"] = move_to_folder
    if copy_to_folder:
        actions["copyToFolder"] = copy_to_folder
    if assign_categories:
        actions["assignCategories"] = assign_categories
    if mark_as_read is not None:
        actions["markAsRead"] = mark_as_read
    if mark_importance:
        actions["markImportance"] = mark_importance
    if forward_to:
        actions["forwardTo"] = _recipients(forward_to)
    if delete is not None:
        actions["delete"] = delete
    if stop_processing_rules is not None:
        actions["stopProcessingRules"] = stop_processing_rules

    payload: dict[str, Any] = {
        "displayName": display_name,
        "sequence": sequence,
        "isEnabled": is_enabled,
    }
    if conditions:
        payload["conditions"] = conditions
    if actions:
        payload["actions"] = actions
    return payload
```

- [ ] **Step 4: Run** `uv run pytest tests/test_rules_helpers.py -v` → all pass.

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/rules.py tests/test_rules_helpers.py
git commit -m "feat(rules): add Graph messageRule payload builder"
```

---

# WAVE 2 — Inbox-rule MCP tools

All tools live in `tools.py`, call Graph at `/me/mailFolders/inbox/messageRules`, and reuse `rules.py`. Folder names in `move_to_folder`/`copy_to_folder` are resolved to IDs via the existing `_resolve_mail_folder()` (tools.py:2251) before building the payload.

### Task 2.1: `list_inbox_rules`

**Depends:** 1.2
**Files:**
- Modify: `src/microsoft_mcp/tools.py` (new tool near the categories/folders tools)
- Test: `tests/test_tools_rules.py`

**Interfaces:**
- Consumes: `graph.request_paginated`, `rules.shape_rule_summary`, `rules.RULE_LIST_FIELDS`, `get_response_profile`, `cleanup_graph_payload`.
- Produces: `list_inbox_rules(response_profile: str = "auto") -> list[dict[str, Any]]`.

- [ ] **Step 1: Write the failing test**

```python
# tests/test_tools_rules.py
from unittest.mock import patch
from src.microsoft_mcp import tools


@patch("src.microsoft_mcp.tools.graph.request_paginated")
def test_list_inbox_rules_assistant_profile(mock_paged):
    mock_paged.return_value = iter([
        {"id": "r1", "displayName": "News", "sequence": 1, "isEnabled": True,
         "conditions": {"senderContains": ["news"]}, "actions": {"markAsRead": True}},
    ])
    out = tools.list_inbox_rules(response_profile="assistant")
    assert out[0]["id"] == "r1"
    assert out[0]["display_name"] == "News"
    assert "news" in out[0]["conditions_summary"]
    called_path = mock_paged.call_args[0][0]
    assert called_path == "/me/mailFolders/inbox/messageRules"
```

- [ ] **Step 2: Run** `uv run pytest tests/test_tools_rules.py::test_list_inbox_rules_assistant_profile -v` → FAIL (no attribute `list_inbox_rules`).

- [ ] **Step 3: Implement** (add to tools.py):

```python
from . import rules as _rules  # add near other intra-package imports at top of tools.py


@mcp.tool
def list_inbox_rules(response_profile: str = "auto") -> list[dict[str, Any]]:
    """List Outlook inbox rules (server-side message rules).

    Returns each rule with a human-readable summary of its conditions and
    actions. Rules run top-to-bottom by `sequence` (lower = earlier).

    Args:
        response_profile: "auto" | "legacy" | "assistant".

    Returns:
        List of rules. assistant: {id, display_name, sequence, is_enabled,
        conditions_summary, actions_summary}. legacy: cleaned raw Graph rules.
    """
    logger.info("list_inbox_rules called")
    profile = get_response_profile(response_profile)
    try:
        raw = list(
            graph.request_paginated(
                "/me/mailFolders/inbox/messageRules",
                params={"$select": _rules.RULE_LIST_FIELDS},
                limit=None,
            )
        )
        if profile == "assistant":
            return [_rules.shape_rule_summary(r) for r in raw]
        return [cleanup_graph_payload(r) for r in raw]
    except Exception as e:
        logger.error(f"list_inbox_rules failed: {e}", exc_info=True)
        raise
```

> Note: if `graph.request_paginated` does not yield from `messageRules` (the endpoint returns a `value` array without `@odata.nextLink`), it still yields each element of `value`; confirm by reading `graph.py:173-201`. If it requires a `value` wrapper that paginated handling already unwraps, no change is needed. If messageRules is non-paginated, fall back to `graph.request("GET", path)["value"]`.

- [ ] **Step 4: Run** `uv run pytest tests/test_tools_rules.py -v` → pass.

- [ ] **Step 5: Commit**

```bash
git add src/microsoft_mcp/tools.py tests/test_tools_rules.py
git commit -m "feat(tools): add list_inbox_rules"
```

---

### Task 2.2: `get_inbox_rule` / `create_inbox_rule` / `update_inbox_rule` / `delete_inbox_rule` / `toggle_inbox_rule`

**Depends:** 2.1, 1.3
**Files:** Modify `tools.py`; Test `tests/test_tools_rules.py`.

These follow the canonical pattern from 2.1. Implement each with the contract below; write one test per tool asserting the Graph method+path+payload and the return shape.

- [ ] **`get_inbox_rule(rule_id: str, response_profile="auto") -> dict`**
  - Graph: `GET /me/mailFolders/inbox/messageRules/{rule_id}` with `$select=RULE_DETAIL_FIELDS`.
  - Return: assistant → `shape_rule_summary` + `exceptions_summary` (call `summarize_conditions(rule.get("exceptions"))`); legacy → `cleanup_graph_payload`.
  - Test: mock `graph.request`, assert path includes `rule_id`, assert summary fields.

- [ ] **`create_inbox_rule(...)`** — params mirror `rules.build_rule_payload` kwargs plus `move_to_folder`/`copy_to_folder` accept folder **names or IDs**.
  - Resolve folder names: `if move_to_folder: move_to_folder = _resolve_mail_folder(move_to_folder)` (same for copy). `_resolve_mail_folder` returns an ID and accepts an ID unchanged.
  - Build body via `_rules.build_rule_payload(**kwargs)`.
  - Graph: `POST /me/mailFolders/inbox/messageRules` json=payload.
  - Return: `cleanup_graph_payload(created)` (legacy) / `shape_rule_summary(created)` (assistant).
  - Test: assert `graph.request` called `("POST", "/me/mailFolders/inbox/messageRules", json=...)` with resolved folder id.

- [ ] **`update_inbox_rule(rule_id, ...)`** — same optional kwargs; build a **partial** payload (only provided fields). For `display_name`/`sequence`/`is_enabled` set top-level keys; for conditions/actions, merge only provided sub-keys (do **not** send empty `conditions`/`actions` objects).
  - Graph: `PATCH /me/mailFolders/inbox/messageRules/{rule_id}` json=partial.
  - Return shaped updated rule.
  - Test: passing only `is_enabled=False` sends `{"isEnabled": False}` and nothing else.

- [ ] **`delete_inbox_rule(rule_id: str) -> dict`**
  - Graph: `DELETE /me/mailFolders/inbox/messageRules/{rule_id}`.
  - Return: `{"status": "deleted", "rule_id": rule_id}`.
  - Test: assert DELETE + path.

- [ ] **`toggle_inbox_rule(rule_id: str) -> dict`**
  - GET the rule's `isEnabled`, then PATCH the inverse.
  - Return: `{"rule_id": rule_id, "is_enabled": <new value>}`.
  - Test: GET returns `isEnabled=True` → PATCH sends `{"isEnabled": False}`.

Each tool: `logger.info` entry, try/except with `logger.error(..., exc_info=True)`.

- [ ] **Final step: Commit**

```bash
git add src/microsoft_mcp/tools.py tests/test_tools_rules.py
git commit -m "feat(tools): add get/create/update/delete/toggle inbox rule"
```

---

### Task 2.3: `reorder_inbox_rules`

**Depends:** 2.2
**Files:** Modify `tools.py`; Test `tests/test_tools_rules.py`.

**Interfaces:** `reorder_inbox_rules(rule_ids_in_order: list[str]) -> list[dict]` — assigns `sequence = index+1` to each rule id by PATCHing each. Returns the new ordering as `[{rule_id, sequence}]`.

- [ ] **Step 1: Test** — given `["rB","rA"]`, expect two PATCH calls: `rB`→`{"sequence":1}`, `rA`→`{"sequence":2}`.
- [ ] **Step 2:** Run → FAIL.
- [ ] **Step 3: Implement** — loop with `enumerate(rule_ids_in_order, start=1)`, `graph.request("PATCH", f"/me/mailFolders/inbox/messageRules/{rid}", json={"sequence": seq})`.
- [ ] **Step 4:** Run → pass.
- [ ] **Step 5: Commit** `feat(tools): add reorder_inbox_rules`.

---

# WAVE 3 — Rule templates + `rules` CLI

### Task 3.1: YAML template ⇄ Graph converters + validation

**Depends:** 1.2, 1.3
**Files:** Modify `src/microsoft_mcp/rules.py`; Test `tests/test_rules_templates.py`.

**Template YAML schema** (snake_case, folder names not IDs):
```yaml
name: Move newsletters
enabled: true
sequence: 5
conditions:
  sender_contains: [news.com, mailchimp.com]
  subject_contains: ["[Newsletter]"]
actions:
  move_to: "Newsletters"      # folder NAME or path; resolved at import
  mark_as_read: true
  assign_categories: [Newsletter]
  stop_processing: true
```

**Interfaces:**
- `template_to_rule_payload(tpl: dict, folder_resolver: Callable[[str], str] | None = None) -> dict` — maps snake_case template → Graph payload via `build_rule_payload`. If `folder_resolver` given, `move_to`/`copy_to` names are resolved to IDs; if `None`, names pass through unchanged (for dry tests).
- `rule_to_template(rule: dict, folder_namer: Callable[[str], str] | None = None) -> dict` — inverse; folder IDs → names if `folder_namer` given.
- `validate_template(tpl: dict) -> list[str]` — returns error strings; empty list = valid. Rules: must have `name`; ≥1 condition; ≥1 action OR `stop_processing: true`; `enabled` bool; `sequence` int; `mark_importance`/`importance` ∈ {low,normal,high} (case-insensitive).

- [ ] **Step 1: Test** round-trip + validation:
```python
from microsoft_mcp import rules

def test_template_to_rule_payload_resolves_folder():
    tpl = {"name": "N", "sequence": 5, "conditions": {"sender_contains": ["x.com"]},
           "actions": {"move_to": "News", "mark_as_read": True, "stop_processing": True}}
    p = rules.template_to_rule_payload(tpl, folder_resolver=lambda n: f"ID::{n}")
    assert p["displayName"] == "N" and p["sequence"] == 5
    assert p["conditions"]["senderContains"] == ["x.com"]
    assert p["actions"]["moveToFolder"] == "ID::News"
    assert p["actions"]["markAsRead"] is True
    assert p["actions"]["stopProcessingRules"] is True

def test_validate_template_requires_condition_and_action():
    assert rules.validate_template({"name": "x"}) != []
    assert rules.validate_template({"name": "x", "conditions": {"sender_contains": ["a"]},
                                    "actions": {"mark_as_read": True}}) == []
```
- [ ] **Step 2:** Run → FAIL.
- [ ] **Step 3: Implement** converters (map snake→camel keys, call `build_rule_payload`; for `rule_to_template`, read `displayName/sequence/isEnabled/conditions/actions` and emit snake_case) and `validate_template`.
- [ ] **Step 4:** Run → pass.
- [ ] **Step 5: Commit** `feat(rules): add YAML template converters + validation`.

---

### Task 3.2: `export_inbox_rules` / `import_inbox_rules` MCP tools

**Depends:** 3.1, 2.2
**Files:** Modify `tools.py`; Test `tests/test_tools_rules.py`.

- [ ] **`export_inbox_rules(path: str | None = None) -> dict`** — fetch all rules (reuse list logic via `graph.request_paginated`), convert each via `rule_to_template` (with a folder-namer that maps folder IDs→display names using `list_mail_folders` data), `yaml.safe_dump({"rules": [...]})`. If `path`, write to disk and return `{"path": path, "count": n}`; else return `{"yaml": <string>, "count": n}`.
  - Test: mock rules + folders, assert YAML contains `name:` and folder names.
- [ ] **`import_inbox_rules(yaml_text: str | None = None, path: str | None = None, mode: str = "create", dry_run: bool = False) -> dict`** — load YAML (`path` or `yaml_text`), for each template `validate_template`; resolve `move_to`/`copy_to` via `_resolve_mail_folder`; `mode="create"` skips names that already exist, `mode="sync"` PATCHes existing by name else POSTs. `dry_run` returns the planned actions without calling Graph mutating endpoints.
  - Return: `{"created": [...], "updated": [...], "skipped": [...], "errors": [...]}`.
  - Test: dry_run returns plan with no POST/PATCH calls (assert `graph.request` not called with POST/PATCH).
- [ ] **Commit** `feat(tools): add export/import inbox rules (YAML)`.

---

### Task 3.3: `rules` CLI sub-app + dispatch + entry point

**Depends:** 3.2
**Files:**
- Create: `src/microsoft_mcp/rules_cli.py`
- Modify: `src/microsoft_mcp/server.py:73-87` (add dispatch), `pyproject.toml` (`microsoft-mcp-rules` entry)
- Test: `tests/test_rules_cli.py`, update `tests/test_server_entry.py`

**Interfaces:** mirror `auth_cli.py` exactly — `main(argv) -> int`, `cli_main()`, zero-dep ANSI color (`_color_enabled`/`_c` copied from `auth_cli.py`), `--json` on every subcommand. Subcommands: `list`, `get <id>`, `export [--output FILE]`, `import <FILE> [--mode create|sync] [--dry-run]`, `create --name ... [--from-contains ...] [--move-to ...] [--mark-read] [--stop]`, `delete <id> [--confirm]`, `toggle <id>`. Each handler imports `microsoft_mcp.tools` lazily and calls the matching `@mcp.tool` function, printing a table (non-JSON) or `json.dumps` (JSON).

- [ ] **Step 1: Test** dispatch + a subcommand:
```python
# tests/test_rules_cli.py
from microsoft_mcp import rules_cli
from unittest.mock import patch

def test_rules_cli_list_json(capsys):
    with patch("microsoft_mcp.tools.list_inbox_rules", return_value=[{"id": "r1", "display_name": "N"}]):
        rc = rules_cli.main(["list", "--json"])
    assert rc == 0
    assert "r1" in capsys.readouterr().out
```
And in `tests/test_server_entry.py` add (mirroring existing auth/signatures dispatch test):
```python
def test_server_dispatches_rules_subcommand():
    import sys
    from microsoft_mcp import server
    with patch("microsoft_mcp.rules_cli.main", return_value=0) as m, \
         patch.object(sys, "argv", ["microsoft-mcp", "rules", "list"]):
        with pytest.raises(SystemExit) as ex:
            server.main()
    assert ex.value.code == 0
    m.assert_called_once_with(["list"])
```
- [ ] **Step 2:** Run both → FAIL.
- [ ] **Step 3: Implement** `rules_cli.py` (copy `auth_cli.py` skeleton: parser, color, handlers) and add to `server.py` after the `auth` block (lines ~84-87):
```python
    if argv and argv[0] == "rules":
        from microsoft_mcp import rules_cli
        sys.exit(rules_cli.main(argv[1:]))
```
Add to `pyproject.toml` `[project.scripts]`: `microsoft-mcp-rules = "microsoft_mcp.rules_cli:cli_main"`.
- [ ] **Step 4:** Run → pass. Also `uv run microsoft-mcp rules --help` shows subcommands.
- [ ] **Step 5: Commit** `feat(cli): add 'microsoft-mcp rules' sub-app`.

---

# WAVE 4 — Focused Inbox overrides

Endpoint family: `/me/inferenceClassification/overrides`. Independent of rules; needs only base. Canonical CRUD pattern.

### Task 4.1: `list_focused_overrides` (+ siblings)

**Depends:** none (base)
**Files:** Modify `tools.py`; Test `tests/test_tools_focused.py`.

- [ ] **`list_focused_overrides(response_profile="auto") -> list[dict]`** — `GET /me/inferenceClassification/overrides` `$select=id,classifyAs,senderEmailAddress`. assistant → `{id, classify_as, email, name}` (flatten `senderEmailAddress.{address,name}`); legacy → cleaned. Full TDD cycle (test asserts path + flattened shape).
- [ ] **`create_focused_override(sender_email: str, classify_as: str = "focused", name: str = "") -> dict`** — `POST` body `{"classifyAs": classify_as, "senderEmailAddress": {"address": sender_email, "name": name or sender_email}}`. Validate `classify_as ∈ {"focused","other"}` (raise ValueError otherwise). Test asserts POST + body + validation error path.
- [ ] **`update_focused_override(override_id: str, classify_as: str) -> dict`** — `PATCH /.../overrides/{id}` `{"classifyAs": classify_as}`.
- [ ] **`delete_focused_override(override_id: str) -> dict`** — `DELETE`; return `{"status":"deleted","override_id":id}`.
- [ ] **Commit** `feat(tools): add Focused Inbox override CRUD`.

---

# WAVE 5 — Reply / forward (draft-first) + send + MailTips + attachments

### Task 5.1: `reply_email_draft` / `reply_all_email_draft` / `forward_email_draft`

**Depends:** none (base)
**Files:** Modify `tools.py`; Test `tests/test_tools_reply.py`.

Draft-first: each uses Graph's draft-creating actions, then optionally PATCHes body/recipients, returning the new **draft** (never sends). Reuse the signature machinery already wired into `create_email_draft` where applicable (pass `signature` through to `apply_signature` on the comment body).

- [ ] **`reply_email_draft(email_id: str, body: str = "", signature: str | None = None, response_profile="auto") -> dict`**
  - `POST /me/messages/{email_id}/createReply` → returns a draft message dict with an `id`.
  - If `body`: apply signature (mirror `create_email_draft`'s signature handling), then `PATCH /me/messages/{draft_id}` `{"body": {"contentType": "HTML", "content": <body+sig>}}`.
  - Return shaped draft (`shape_email_detail` legacy / compact assistant `{id, web_link, is_draft: True}`).
  - Test: assert createReply POST, assert PATCH body when `body` given, assert NO `/send` call ever.
- [ ] **`reply_all_email_draft(...)`** — identical but `POST /me/messages/{id}/createReplyAll`.
- [ ] **`forward_email_draft(email_id, to: list[str], comment: str = "", signature=None, ...)`** — `POST /me/messages/{id}/createForward`, then PATCH `toRecipients` = `[{"emailAddress":{"address": a}} for a in to]` and body=comment(+sig). Validate `to` non-empty.
- [ ] **Commit** `feat(tools): add reply/reply-all/forward as drafts`.

---

### Task 5.2: `send_email_draft` (explicit send)

**Depends:** 5.1
**Files:** Modify `tools.py`; Test `tests/test_tools_reply.py`.

**Interfaces:** `send_email_draft(draft_id: str) -> dict` — `POST /me/messages/{draft_id}/send` (202). Return `{"status": "sent", "draft_id": draft_id}`. This is the ONLY tool that puts mail on the wire; docstring states so explicitly.

- [ ] TDD cycle: test asserts POST to `/send`; return shape. Commit `feat(tools): add explicit send_email_draft`.

---

### Task 5.3: `get_mailtips`

**Depends:** none (base)
**Files:** Modify `tools.py`; Test `tests/test_tools_mailtips.py`.

**Interfaces:** `get_mailtips(emails: list[str], options: list[str] | None = None) -> list[dict]`.
- `POST /me/getMailTips` body `{"EmailAddresses": emails, "MailTipsOptions": ",".join(options or DEFAULT)}` where `DEFAULT = ["automaticReplies","mailboxFullStatus","maxMessageSize","recipientScope","deliveryRestriction"]`.
- Shape each tip: `{email, auto_reply (HTML-stripped message or None), mailbox_full (bool), max_message_size_bytes, recipient_scope, delivery_restricted}`. Reuse `response_shaping._html_to_text` for the auto-reply message.

- [ ] TDD: mock POST returning a `value` array of tips, assert flattened shape + HTML stripped. Commit `feat(tools): add get_mailtips`.

---

### Task 5.4: `list_attachments` + `download_attachments`

**Depends:** none (base)
**Files:** Modify `tools.py`; Test `tests/test_tools_attachments.py`.

- [ ] **`list_attachments(email_id: str) -> list[dict]`** — `GET /me/messages/{email_id}/attachments` `$select=id,name,contentType,size,isInline,@odata.type`. Return `[{id, name, content_type, size, is_inline, kind}]` where kind = "file"/"item" from `@odata.type`.
- [ ] **`download_attachments(email_id: str, save_dir: str, names: list[str] | None = None) -> dict`** — list attachments; for each file attachment (optionally filtered by `names`), if `contentBytes` absent re-fetch single attachment, base64-decode, write to `save_dir/<name>`. Return `{"saved": [paths], "skipped": [names]}`. Reuse the os.path expanduser + write pattern from `get_attachment` (tools.py:3384).
  - Test: mock list + single-attachment fetch with base64 content; assert file written to tmp_path; non-file attachments skipped.
- [ ] **Commit** `feat(tools): add list_attachments + bulk download_attachments`.

---

# WAVE 6 — Microsoft To-Do

Endpoints under `/me/todo`. Pure helpers in `todo.py`; tools in `tools.py`.

### Task 6.1: `todo.py` helpers (due-date parsing + payloads)

**Depends:** none (base)
**Files:** Create `src/microsoft_mcp/todo.py`; Test `tests/test_todo_helpers.py`.

**Interfaces:**
- `parse_due_date(text: str, *, today: date) -> dict[str, str]` → `{"dateTime": "YYYY-MM-DDT23:59:00", "timeZone": "UTC"}`. Accepts `today`, `tomorrow`, `+Nd`, `YYYY-MM-DD`. Raises `ValueError` on unparseable. **`today` is injected** (never call `date.today()` inside — keeps tests deterministic and matches the workflow no-`Date.now()` discipline).
- `build_task_payload(*, title, importance="normal", body=None, due=None) -> dict` → Graph `todoTask` body.
- `build_linked_resource(web_url: str, display_name: str) -> dict` → `{"applicationName":"Outlook","webUrl":web_url,"displayName":display_name}`.

- [ ] **Step 1: Test**:
```python
from datetime import date
from microsoft_mcp import todo

def test_parse_due_relative():
    assert todo.parse_due_date("+3d", today=date(2026, 6, 16))["dateTime"].startswith("2026-06-19")

def test_parse_due_absolute():
    assert todo.parse_due_date("2026-07-01", today=date(2026, 6, 16))["dateTime"].startswith("2026-07-01")

def test_parse_due_invalid():
    import pytest
    with pytest.raises(ValueError):
        todo.parse_due_date("someday", today=date(2026, 6, 16))
```
- [ ] **Step 2-4:** implement + green.
- [ ] **Step 5: Commit** `feat(todo): add due-date parser + task payload helpers`.

---

### Task 6.2: To-Do list + task tools

**Depends:** 6.1
**Files:** Modify `tools.py`; Test `tests/test_tools_todo.py`. Add a `_resolve_todo_list(name_or_id) -> str` helper (lists `/me/todo/lists`, matches by `displayName` or id, optionally auto-creates).

Implement (canonical-pattern CRUD; one test each):
- [ ] `list_todo_lists(response_profile="auto")` — `GET /me/todo/lists`.
- [ ] `create_todo_list(name: str)` — `POST /me/todo/lists` `{"displayName": name}`.
- [ ] `list_tasks(list_name: str, status: str | None = None, response_profile="auto")` — `GET /me/todo/lists/{id}/tasks` with optional `$filter=status eq '...'`.
- [ ] `create_task(list_name, title, importance="normal", body="", due="")` — resolve list (auto-create if missing), `build_task_payload` (parse `due` via `todo.parse_due_date(due, today=date.today())` — `date.today()` is called in the **tool**, passed into the pure helper), `POST .../tasks`.
- [ ] `update_task(list_name, task_id, **fields)` — partial `PATCH`.
- [ ] `complete_task(list_name, task_id)` — `PATCH` `{"status":"completed"}`.
- [ ] `delete_task(list_name, task_id)` — `DELETE`.
- [ ] **Commit** `feat(tools): add Microsoft To-Do list + task CRUD`.

---

### Task 6.3: Checklist items + `create_task_from_email`

**Depends:** 6.2
**Files:** Modify `tools.py`; Test `tests/test_tools_todo.py`.

- [ ] `list_checklist_items(list_name, task_id)` — `GET .../tasks/{id}/checklistItems`.
- [ ] `add_checklist_item(list_name, task_id, text, is_checked=False)` — `POST` `{"displayName":text,"isChecked":is_checked}`.
- [ ] `update_checklist_item(list_name, task_id, item_id, text=None, is_checked=None)` — partial PATCH.
- [ ] `delete_checklist_item(list_name, task_id, item_id)` — DELETE.
- [ ] `create_task_from_email(email_id, list_name, title=None, importance="normal")` — `GET /me/messages/{email_id}?$select=subject,webLink`; build task with `title or "Follow up: "+subject` and `linkedResources=[todo.build_linked_resource(webLink, "View Email")]`; POST.
  - Test: mock message fetch + task POST; assert `linkedResources[0].webUrl == webLink`.
- [ ] **Commit** `feat(tools): add To-Do checklist items + create_task_from_email`.

---

# WAVE 7 — Template engine

### Task 7.1: `templates_engine.py` — loader + renderer

**Depends:** 1.1 (pyyaml)
**Files:** Create `src/microsoft_mcp/templates_engine.py`, bundled YAML under `src/microsoft_mcp/templates_data/`; Test `tests/test_templates_engine.py`.

**Interfaces:**
- `template_dirs() -> list[Path]` — search path: user dir `~/.config/microsoft-mcp/templates/` (override `MICROSOFT_MCP_TEMPLATES_DIR`) first, then bundled `templates_data/`. User shadows bundled by name.
- `list_templates(category: str | None = None) -> list[dict]` — `{name, description, version, category, source, placeholders}`.
- `load_template(category: str, name: str) -> dict` — parse YAML, require `name` + `html_template`.
- `validate_template_data(tpl: dict, data: dict) -> list[str]` — missing required placeholders.
- `render_template(category: str, name: str, data: dict) -> str` — substitute `{placeholder}` via regex `\{(\w+)\}`, evaluate `conditional_sections` (`field`, `a|b`, `a&b`), HTML-escape plain values except pre-rendered list keys, strip empty lines.

- [ ] **Step 1: Test** with a tiny inline template written to `tmp_path` (set `MICROSOFT_MCP_TEMPLATES_DIR`):
```python
def test_render_substitutes_and_conditionals(tmp_path, monkeypatch):
    cat = tmp_path / "email"; cat.mkdir()
    (cat / "hi.yaml").write_text(
        "name: hi\nhtml_template: '<p>Hi {first}{closing}</p>'\n"
        "placeholders:\n  - {name: first, required: true}\n"
        "conditional_sections:\n  closing: {condition: sign, template: ', {sign}'}\n")
    monkeypatch.setenv("MICROSOFT_MCP_TEMPLATES_DIR", str(tmp_path))
    from microsoft_mcp import templates_engine as te
    out = te.render_template("email", "hi", {"first": "Sam", "sign": "JP"})
    assert "Hi Sam, JP" in out
```
- [ ] **Step 2-4:** implement loader/renderer (port `loader.py`/`renderer.py` logic; drop the Rich/CLI bits) + green.
- [ ] **Step 5: Commit** `feat(templates): add YAML template loader + renderer`.

---

### Task 7.2: Draft variable substitution + CSV merge

**Depends:** 7.1
**Files:** Modify `templates_engine.py`; Test `tests/test_templates_engine.py`.

**Interfaces:**
- `find_template_variables(content: str, decode_html: bool = True) -> list[str]` — unique `{{var}}` (and HTML-encoded `&#123;&#123;`) in order.
- `substitute_variables(content: str, values: dict, strict: bool = False) -> str` — replace; `strict` raises `VariableSubstitutionError` on missing.
- `parse_recipients_csv(path: str) -> list[dict]` — UTF-8/UTF-8-sig/Latin-1 fallback; headers→keys.

- [ ] TDD: variable detection, substitution (strict + non-strict), CSV parse from `tmp_path`. Commit `feat(templates): add {{var}} substitution + CSV recipients`.

---

### Task 7.3: Template MCP tools + draft integration

**Depends:** 7.2
**Files:** Modify `tools.py`; Test `tests/test_tools_templates.py`. Bundle 4 starter templates (`calendar/meeting.yaml`, `calendar/interview.yaml`, `email/_schema.yaml`, `calendar/_schema.yaml`) copied/trimmed from outlook-creds `templates/`.

- [ ] `list_email_templates(category="")`, `get_template_placeholders(category, name)`, `render_template(category, name, data)` — thin wrappers over `templates_engine`.
- [ ] `find_template_variables(content)` + `substitute_template_variables(content, values, strict=False)` tools.
- [ ] Integrate into `create_email_draft` (tools.py:1937): add optional `template: str | None = None` (format `"category/name"`) and `template_data: dict | None = None`; when set, body = `render_template(...)` before signature application. Add a regression test that `create_email_draft(template="email/hi", template_data={...})` produces the rendered body in the PATCH/POST payload.
- [ ] **Commit** `feat(tools): expose template engine + wire into create_email_draft`.

---

# WAVE 8 — Signature parser + Intel engine

### Task 8.1: `signature_parser.py`

**Depends:** none (base)
**Files:** Create `src/microsoft_mcp/signature_parser.py`; Test `tests/test_signature_parser.py`.

**Interfaces:**
- `normalize_phone_e164(phone: str, default_region: str = "US") -> str` — 10 digits→`+1…`, 11 w/ leading 1→`+1…`, else best-effort; `""` if junk.
- `parse_signature_block(text: str) -> dict` — `{first_name,last_name,full_name,job_title,company,work_email,mobile_phone,business_phone,website,linkedin,twitter,confidence_score}`.
- `parse_email_body(body: str, *, html: bool = False, extract_alternatives: bool = True) -> dict` — `{contacts: [...], job_changes: {...}}`; strips HTML when `html`.

- [ ] TDD: phone normalization (`"(949) 462-4106"`→`"+19494624106"`), name+title+email extraction from a sample block, LinkedIn handle, confidence in `[0,1]`. Port regex/heuristics from outlook-creds `api/signature_parser.py` (no `lxml`; use `response_shaping._html_to_text` for HTML).
- [ ] **Commit** `feat(signatures): add signature/OOO contact parser`.

---

### Task 8.2: Signature MCP tools

**Depends:** 8.1
**Files:** Modify `tools.py`; Test `tests/test_tools_signature.py`.

- [ ] `parse_email_signature(email_body: str, is_html: bool = False, extract_alternatives: bool = True) -> dict` — wraps `parse_email_body`.
- [ ] `normalize_phone_number(phone: str, region: str = "US") -> str` — wraps `normalize_phone_e164`.
- [ ] **Commit** `feat(tools): add signature parsing tools`.

---

### Task 8.3: Intel types + collectors

**Depends:** none (base; can parallel 8.1)
**Files:** Create `src/microsoft_mcp/intel/{__init__,types,_utils}.py`, `intel/collectors/{email,calendar,threads,contacts}.py`; Test `tests/test_intel_collectors.py`.

Each collector takes a `request` callable (dependency-injected = `graph.request`) + a timezone + look-back window and returns the typed signal dict from the outlook-creds report (`EmailSignals`, `CalendarSignals`, `ThreadSignals`, `ContactSignals`). Inject the Graph caller so tests pass a fake — do **not** import `graph` at module top inside `intel/` (keeps the package pure + testable).

- [ ] TDD per collector with a fake `request` returning canned `value` arrays; assert the derived counts (unread totals, today's events, awaiting-reply threads via `conversationId` grouping, VIP threshold ≥3). Commit per collector or one `feat(intel): add signal collectors`.

---

### Task 8.4: Intel analyzers + engine

**Depends:** 8.3
**Files:** Create `intel/analyzers/{priority,schedule,relationships}.py`, `intel/engine.py`; Test `tests/test_intel_engine.py`.

- [ ] `score_priorities(emails, calendar, threads) -> list[PriorityItem]` with the documented base scores + modifiers (needs_response 60, vip 50, conflict 90, prep 40, awaiting_my_reply 55, stale 30).
- [ ] `analyze_schedule(calendar) -> ScheduleAnalysis`; `analyze_relationships(contacts) -> list[RelationshipInsight]`.
- [ ] `engine.generate_briefing/generate_signals/generate_contact_report/generate_recap(request, account, tz, ...)` orchestrating collectors+analyzers.
- [ ] TDD: feed canned collector outputs, assert priority ordering + signal bucketing (critical ≥80, important 50–79, info <50). Commit `feat(intel): add analyzers + report engine`.

---

### Task 8.5: Intel MCP tools + `intel` CLI

**Depends:** 8.4
**Files:** Modify `tools.py`; Create `src/microsoft_mcp/intel_cli.py`; Modify `server.py`, `pyproject.toml`; Test `tests/test_tools_intel.py`, `tests/test_intel_cli.py`, update `tests/test_server_entry.py`.

- [ ] Tools: `generate_morning_briefing(timezone="UTC", limit=10)`, `get_priority_signals(timezone="UTC", level="all")`, `get_contact_intelligence(target_email, days=30)`, `get_end_of_day_recap(timezone="UTC")` — each calls `engine.*` passing `graph.request`.
- [ ] `intel_cli.py` (mirror `auth_cli`): `briefing|signals|contact <email>|recap`, each `--json`. Dispatch in `server.py` (`argv[0] == "intel"`) + `microsoft-mcp-intel` entry point + server-dispatch regression test.
- [ ] **Commit** `feat(intel): add MCP tools + 'microsoft-mcp intel' CLI`.

---

# WAVE 9 — Bounce detection + CLI

### Task 9.1: `bounces.py` — patterns + classifier + DSN parse

**Depends:** none (base)
**Files:** Create `src/microsoft_mcp/bounces.py`; Test `tests/test_bounces.py`.

**Interfaces (ported from outlook-creds `api/bounces.py`):**
- `is_bounce_message(subject, sender_email, body=None, *, use_body=True) -> bool`
- `determine_bounce_reason(subject, body) -> str`
- `classify_bounce_message(msg: dict) -> dict` — `{first_name,last_name,email,reason,date,iso_date,subject,sender,body,message_id,has_attachments}`
- `parse_dsn_content(text: str) -> dict` — `{final_recipient,action,status,diagnostic_code,display_name}`
- `extract_email_from_text`, `parse_name_from_email`
- Pattern catalogs as module constants (subject keywords, sender patterns, body patterns, reason regexes).

- [ ] TDD: a postmaster "Undeliverable" subject → `is_bounce_message` True; `determine_bounce_reason` maps `550 5.1.1`→"Invalid Recipient"; `parse_dsn_content` extracts `final_recipient`/`status` from a sample DSN. Commit `feat(bounces): add NDR classifier + DSN parser`.

---

### Task 9.2: Folder scan + CSV + `bounces` CLI + optional tool

**Depends:** 9.1
**Files:** Modify `bounces.py` (add `iter_folder_messages(request, folder_id, limit)`, `scan_folder(request, folder_id, limit) -> list[dict]`, `write_csv(rows, path)`); Create `src/microsoft_mcp/bounces_cli.py`; Modify `server.py`, `pyproject.toml`, `tools.py`; Test `tests/test_bounces.py`, `tests/test_bounces_cli.py`, update `tests/test_server_entry.py`.

- [ ] `scan_folder` uses injected `request` + `@odata.nextLink` pagination + `is_bounce_message`/`classify_bounce_message`.
- [ ] MCP tool `scan_bounces(folder="Inbox", limit=200, save_csv: str | None = None) -> dict` — resolve folder via `_resolve_mail_folder`, scan via `graph.request`, return `{count, reasons: {...}, rows: [...]}` and write CSV if `save_csv`.
- [ ] `bounces_cli.py`: `scan [--folder] [--limit] [--output CSV] [--json]`, `patterns [--json]`. Dispatch `argv[0]=="bounces"` + `microsoft-mcp-bounces` entry + dispatch regression test.
- [ ] **Commit** `feat(bounces): add folder scan, CSV export, 'microsoft-mcp bounces' CLI + scan_bounces tool`.

---

# WAVE 10 — Integration, docs, regression

### Task 10.1: Tool-surface + entry-point regression

**Depends:** ALL prior waves
**Files:** Modify `tests/test_server_entry.py` (and any tool-count test); run `/weekly-audit` mentally.

- [ ] Update the expected tool count to the new total (was 65; add the count of new `@mcp.tool`s — verify with `uv run python -c "from microsoft_mcp import tools; print(len([t for t in tools.mcp._tool_manager._tools]))"`).
- [ ] Confirm all three new CLI dispatch tests (`rules`, `intel`, `bounces`) pass alongside existing `auth`/`signatures` ones.
- [ ] Run the FULL suite: `uv run pytest tests/ -v` → all green.
- [ ] `uv run pyright` → clean. `uvx ruff check --fix --unsafe-fixes . && uvx ruff format .` → clean.
- [ ] **Commit** `test: update tool-surface + CLI-dispatch regression for mail port`.

---

### Task 10.2: Docs sync

**Depends:** 10.1
**Files:** `CLAUDE.md`, `IMPLEMENTATION.md`, `README.md`.

- [ ] CLAUDE.md: under "Core Modules", document `rules.py`/`rules_cli.py`, `todo.py`, `templates_engine.py` (+`templates_data/`), `signature_parser.py`, `intel/`, `bounces.py`/`bounces_cli.py`; under tools, list the new families; add the new env vars (`MICROSOFT_MCP_TEMPLATES_DIR`); document the 3 new CLI sub-apps + their `microsoft-mcp <name>` dispatch and standalone console scripts; note **Graph-only, no EWS** and **draft-first reply/forward** as design decisions.
- [ ] IMPLEMENTATION.md: architectural notes for the rule data model, template search path, intel collector→analyzer→engine pipeline, bounce pattern catalog.
- [ ] README.md: usage snippets for `microsoft-mcp rules import rules.yaml`, `microsoft-mcp intel briefing`, `microsoft-mcp bounces scan`.
- [ ] Run `doc-sync` agent if available.
- [ ] **Commit** `docs: document mail-rule/templates/todo/intel/bounces port`.

---

## Self-Review

**Spec coverage** (against the chosen "Everything + all add-ons" scope):
- A Inbox rules → Wave 2 (+reorder 2.3). ✅
- B Rule templates → Wave 3. ✅
- C Focused Inbox → Wave 4. ✅
- D Reply/forward (draft) + send → Wave 5.1–5.2. ✅
- E MailTips → 5.3. ✅
- F Attachments → 5.4. ✅
- G To-Do → Wave 6. ✅
- H Template engine → Wave 7. ✅
- I Signature parser → Wave 8.1–8.2. ✅
- J Intel → Wave 8.3–8.5. ✅
- K Bounces → Wave 9. ✅
- Docs + regression → Wave 10. ✅

**Type consistency:** `rules.build_rule_payload` kwargs ↔ `create_inbox_rule`/`template_to_rule_payload` callers align (snake_case in, camelCase Graph out). `todo.parse_due_date(text, *, today)` is called with `today` everywhere. Intel collectors take an injected `request` callable consistently.

**Placeholder scan:** Canonical code is complete for the first tool of each family; siblings are specified by exact endpoint/method/payload/return/test, per the stated "pattern shown once" convention — not vague TODOs.

**Open risk to verify during execution:** whether `graph.request_paginated` cleanly yields `messageRules`/`overrides`/`todo` collection items (these endpoints return a `value` array; confirm against `graph.py:173-201` and fall back to `graph.request("GET", path)["value"]` if a given endpoint isn't paginated). Flagged in Task 2.1.
