# Tool Surface Audit — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Fix the 16 confirmed bugs and 11 convention-drift issues identified in the 2026-04-23 microsoft-mcp tool-surface audit, gated by failing tests and separated into five sequential passes.

**Architecture:** Each pass is its own PR-sized unit; within a pass each task is one bug with its own failing-test-first commit. Pass 1 fixes user-visible bugs (sandbox, envelope bloat, ranker, example). Pass 2 closes convention drift (response_profile sweep, dead code removal, search_cache writers). Pass 3 hardens auth + HTTP robustness (retry, locking, scope preservation, param mutation). Pass 4 is hygiene (rename/encode/atomic). Pass 5 adds regression guardrails.

**Tech Stack:** Python 3.12+, uv, pytest, pytest-asyncio, ruff, pyright, RestrictedPython, httpx (sync client), MSAL, Azure SDK, FastMCP.

**Branch:** `feature/tool-surface-audit` (already created off `feature/alternative-auth-method` at `afcbe00`).

**Out of scope (deferred):** `search_query` cursor migration, MarkItDown async, httpx.Client connection-pool rework.

---

## Working rules for every task

- One failing test → run to confirm failure → minimal fix → run to confirm pass → commit. Do not skip the "run to confirm failure" step.
- Commit messages: `fix(<area>): <bug-id> <short>` / `refactor(<area>): ...` / `chore(<area>): ...`, matching the repo's conventional-commit style.
- Run `/lint` (or `uv run pyright && uvx ruff check .`) before every commit; fix anything red. The `PostToolUse` hook already runs `ruff format` + `ruff check --fix` automatically on Write/Edit/MultiEdit.
- Do **not** bundle fixes. One commit per bug. Reverts must be targeted.
- The FastMCP tools in `tools.py` are decorated with `@mcp.tool`; test them by calling `.fn(...)` directly (see `tests/test_tool_contracts.py` for the pattern).
- HTTP tests mock `microsoft_mcp.graph.request` (not `httpx`). See `tests/conftest.py` for `mock_auth` and `sample_*_data` fixtures.
- If a test file already exists for the module you're touching, add to it rather than creating a new one.

---

## Pass 1 — Correctness (user-visible bugs)

### Task 1: B1 — Install RestrictedPython iteration guards in the code-mode sandbox

Agent code currently cannot use comprehensions or `for` loops inside `call_tool_chain` — it raises `name '_getiter_' is not defined`. Reproduced live.

**Files:**
- Modify: `src/microsoft_mcp/code_mode.py:541-616`
- Test: `tests/test_code_mode_tools.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_code_mode_tools.py`:

```python
import pytest
import asyncio
from microsoft_mcp.code_mode import CodeModeRuntime


@pytest.mark.asyncio
async def test_call_tool_chain_supports_list_comprehensions(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain(
        "return [n * 2 for n in range(4)]"
    )
    assert result["result"] == [0, 2, 4, 6]


@pytest.mark.asyncio
async def test_call_tool_chain_supports_for_loops(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain(
        """
acc = []
for n in range(3):
    acc.append(n * n)
return acc
"""
    )
    assert result["result"] == [0, 1, 4]


@pytest.mark.asyncio
async def test_call_tool_chain_supports_augmented_assignment(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain(
        """
total = 0
for n in range(5):
    total += n
return total
"""
    )
    assert result["result"] == 10
```

If `mcp_with_runtime` fixture does not exist in `conftest.py`, add this fixture there:

```python
import pytest
from fastmcp import FastMCP
from microsoft_mcp.code_mode import CodeModeRuntime


@pytest.fixture
async def mcp_with_runtime():
    mcp = FastMCP("test")
    return await CodeModeRuntime.create(mcp)
```

- [ ] **Step 2: Run test to verify it fails**

```
uv run pytest tests/test_code_mode_tools.py::test_call_tool_chain_supports_list_comprehensions -v
```

Expected: FAIL with `name '_getiter_' is not defined` (or `_inplacevar_` / `_iter_unpack_sequence_` depending on which guard fires first).

- [ ] **Step 3: Implement guards in the sandbox**

Edit `src/microsoft_mcp/code_mode.py` — replace `_load_restricted_python_globals` (currently at lines 602–616) with:

```python
def _load_restricted_python_globals(self) -> dict[str, Any]:
    try:
        guards = importlib.import_module("RestrictedPython.Guards")
        eval_mod = importlib.import_module("RestrictedPython.Eval")
        print_collector = importlib.import_module("RestrictedPython.PrintCollector")
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "RestrictedPython is required for the code-mode sandbox."
        ) from exc

    safe_globals = dict(getattr(guards, "safe_globals", {}))

    # Iteration + subscripting + augmented-assignment guards.
    # Without these, comprehensions, `for` loops, and `+=` all fail.
    safe_globals["_getiter_"] = getattr(eval_mod, "default_guarded_getiter", iter)
    safe_globals["_getitem_"] = getattr(
        eval_mod, "default_guarded_getitem", lambda obj, key: obj[key]
    )
    safe_globals["_iter_unpack_sequence_"] = getattr(
        guards, "guarded_iter_unpack_sequence", lambda it, spec, _getiter_: tuple(it)
    )
    safe_globals["_unpack_sequence_"] = getattr(
        guards, "guarded_unpack_sequence", lambda it, spec, _getiter_: tuple(it)
    )
    safe_globals["_inplacevar_"] = _inplace_var

    shared_print_collector = print_collector.PrintCollector()
    safe_globals["_print_"] = lambda _getattr=None: shared_print_collector
    safe_globals["_print"] = shared_print_collector
    safe_globals["__shared_print_collector__"] = shared_print_collector
    return safe_globals
```

Add this helper at module scope (above `class CodeModeRuntime`):

```python
def _inplace_var(op: str, x: Any, y: Any) -> Any:
    """Support for `+=`, `-=`, etc. inside the RestrictedPython sandbox."""
    if op == "+=":
        return x + y
    if op == "-=":
        return x - y
    if op == "*=":
        return x * y
    if op == "/=":
        return x / y
    if op == "//=":
        return x // y
    if op == "%=":
        return x % y
    if op == "**=":
        return x ** y
    if op == "|=":
        return x | y
    if op == "&=":
        return x & y
    if op == "^=":
        return x ^ y
    if op == "<<=":
        return x << y
    if op == ">>=":
        return x >> y
    raise ValueError(f"Unsupported inplace operator: {op}")
```

- [ ] **Step 4: Run tests to verify they pass**

```
uv run pytest tests/test_code_mode_tools.py -v
```

Expected: all three new tests pass. Full suite must stay green — also run:

```
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/code_mode.py tests/test_code_mode_tools.py tests/conftest.py
git commit -m "fix(code-mode): B1 install RestrictedPython iteration guards

Sandbox was missing _getiter_, _iter_unpack_sequence_, and _inplacevar_,
which caused all list/dict/set comprehensions, for loops, and augmented
assignments to fail with 'name _getiter_ is not defined' inside
call_tool_chain."
```

---

### Task 2: B2 — Make the call_tool_chain interface catalog opt-in

Every successful `call_tool_chain` response currently includes the entire generated TypedDict catalog (`interfaces`), its JSON sibling (`interface_map_json`), and two `available_*` arrays, ballooning responses by ~50 KB regardless of user need.

**Files:**
- Modify: `src/microsoft_mcp/code_mode.py:292-379` (`call_tool_chain`)
- Modify: `src/microsoft_mcp/tools.py` (the `@mcp.tool` wrapper for `call_tool_chain`; grep for the definition)
- Test: `tests/test_code_mode_tools.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_code_mode_tools.py`:

```python
@pytest.mark.asyncio
async def test_call_tool_chain_default_excludes_interface_catalog(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain("return 42")
    assert result["result"] == 42
    assert "interfaces" not in result
    assert "interface_map_json" not in result
    assert "available_tools" not in result
    assert "available_access_patterns" not in result


@pytest.mark.asyncio
async def test_call_tool_chain_include_interfaces_flag(mcp_with_runtime):
    runtime = mcp_with_runtime
    result = await runtime.call_tool_chain("return 1", include_interfaces=True)
    assert "interfaces" in result
    assert "interface_map_json" in result
    assert "available_tools" in result
```

- [ ] **Step 2: Run test to verify it fails**

```
uv run pytest tests/test_code_mode_tools.py::test_call_tool_chain_default_excludes_interface_catalog -v
```

Expected: FAIL on `assert "interfaces" not in result` — it's present.

- [ ] **Step 3: Implement**

Edit `src/microsoft_mcp/code_mode.py`. Change the `call_tool_chain` signature and final return:

```python
async def call_tool_chain(
    self,
    code: str,
    timeout: float | None = None,
    include_interfaces: bool = False,
) -> dict[str, Any]:
    """Execute trusted code against the live tool registry in a sandbox.

    By default the response contains only the user-code result, logs, and
    trace. Pass ``include_interfaces=True`` to also embed the generated
    TypedDict catalog (useful for first-run discovery, expensive in tokens
    on every call).
    """
    # ...existing body until the return block...
    try:
        result = await asyncio.wait_for(asyncio.to_thread(execute), timeout=timeout)
        collector = sandbox.get("__shared_print_collector__")
        if collector is not None:
            output = collector()
            if output:
                for line in str(output).splitlines():
                    logs.append(line)

        response: dict[str, Any] = {
            "result": result,
            "logs": logs,
            "trace": trace,
        }
        if include_interfaces:
            response["interfaces"] = interfaces
            response["interface_map_json"] = interface_map_json
            response["available_tools"] = available_tools
            response["available_access_patterns"] = available_access_patterns
        return response
    except asyncio.TimeoutError as exc:
        logs.append(f"[ERROR] Code execution timed out after {timeout} seconds.")
        raise TimeoutError(
            f"Code execution timed out after {timeout} seconds."
        ) from exc
    except Exception as exc:
        logs.append(f"[ERROR] {exc}")
        raise
    finally:
        self._trace_sink = None
```

Now find the `@mcp.tool` wrapper for `call_tool_chain` in `tools.py` (grep: `rg -n 'def call_tool_chain' src/microsoft_mcp/tools.py`). Add the new parameter:

```python
@mcp.tool
def call_tool_chain(
    code: str,
    timeout: float = 30.0,
    include_interfaces: bool = False,
) -> dict[str, Any]:
    """Execute a multi-step Python workflow against the active Microsoft tool namespace.

    Set include_interfaces=True only when you need the generated TypedDict
    catalog in the response (default False to keep token cost low).
    """
    runtime = _get_code_mode_runtime()
    return _run_async(
        runtime.call_tool_chain(code, timeout=timeout, include_interfaces=include_interfaces)
    )
```

- [ ] **Step 4: Run tests to verify they pass**

```
uv run pytest tests/test_code_mode_tools.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/code_mode.py src/microsoft_mcp/tools.py tests/test_code_mode_tools.py
git commit -m "fix(code-mode): B2 make interface catalog opt-in in call_tool_chain

Every successful call_tool_chain response carried the full generated
TypedDict catalog (~50KB), wasting tokens on every iteration of an agent
loop. Catalog is now gated behind include_interfaces=True; discovery
continues through the existing tools_info and get_interfaces methods."
```

---

### Task 3: B3 — Populate `starts_in_minutes` on invite messages and events

Inbox ranker computes meeting-proximity tiers (25/15/5 bonuses) from `item.starts_in_minutes`, but the three `_*_to_inbox_items` helpers never set it. Every ranked item scores only the `unread` bonus. Reproduced live (seven consecutive items at score=10).

**Files:**
- Modify: `src/microsoft_mcp/tools.py:4028-4080` (`_invite_messages_to_inbox_items`, `_events_to_inbox_items`)
- Test: `tests/test_inbox_ranking.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_inbox_ranking.py`:

```python
import datetime as dt
from microsoft_mcp.inbox_ranking import _compute_score, rank_items
from microsoft_mcp.inbox_models import InboxItem
from microsoft_mcp import tools as tools_mod


def _future_iso(minutes: int) -> str:
    return (dt.datetime.now(dt.timezone.utc) + dt.timedelta(minutes=minutes)).isoformat()


def test_invite_message_populates_starts_in_minutes_under_15():
    raw = [{
        "id": "msg-1",
        "subject": "Imminent standup",
        "meetingMessageType": "meetingRequest",
        "startDateTime": {"dateTime": _future_iso(5)},
        "isRead": False,
    }]
    items = tools_mod._invite_messages_to_inbox_items(raw)
    assert items[0].starts_in_minutes is not None
    assert items[0].starts_in_minutes <= 15
    # Ranker must now award the 25-point near-meeting bonus on top of unread (+10).
    assert _compute_score(items[0]) >= 35


def test_event_populates_starts_in_minutes_1_to_2_hours():
    raw = [{
        "id": "evt-1",
        "subject": "Later meeting",
        "start": {"dateTime": _future_iso(90)},
    }]
    items = tools_mod._events_to_inbox_items(raw)
    assert 60 < items[0].starts_in_minutes <= 120
    assert _compute_score(items[0]) == 5.0


def test_past_events_have_none_starts_in_minutes():
    raw = [{
        "id": "evt-past",
        "subject": "Already happened",
        "start": {"dateTime": _future_iso(-30)},
    }]
    items = tools_mod._events_to_inbox_items(raw)
    assert items[0].starts_in_minutes is None or items[0].starts_in_minutes < 0
    # No proximity bonus for past events.
    assert _compute_score(items[0]) == 0.0
```

- [ ] **Step 2: Run test to verify it fails**

```
uv run pytest tests/test_inbox_ranking.py::test_invite_message_populates_starts_in_minutes_under_15 -v
```

Expected: FAIL — `starts_in_minutes` is `None`.

- [ ] **Step 3: Implement**

Add a parse helper near the top of `tools.py` (grep for `def _list_message_summaries` — place it above that or with the other inbox helpers):

```python
def _parse_graph_datetime(value: Any) -> dt.datetime | None:
    """Parse a Graph-supplied ISO datetime string into an aware UTC datetime."""
    if not value or not isinstance(value, str):
        return None
    # Graph returns formats like "2026-05-06T16:00:00.0000000" (no tz) or
    # "2026-05-06T16:00:00Z". Normalize both.
    cleaned = value.rstrip("Z")
    # Trim fractional seconds beyond 6 digits (Python's limit).
    if "." in cleaned:
        base, frac = cleaned.split(".", 1)
        frac = frac[:6]
        cleaned = f"{base}.{frac}"
    try:
        parsed = dt.datetime.fromisoformat(cleaned)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=dt.timezone.utc)
    return parsed


def _minutes_until(value: Any) -> float | None:
    parsed = _parse_graph_datetime(value)
    if parsed is None:
        return None
    delta = (parsed - dt.datetime.now(dt.timezone.utc)).total_seconds()
    if delta < 0:
        return None
    return delta / 60.0
```

Update `_invite_messages_to_inbox_items` (currently `tools.py:4028-4060`) so the `InboxItem(...)` call includes:

```python
starts_in_minutes=_minutes_until(message.get("startDateTime", {}).get("dateTime")),
```

Update `_events_to_inbox_items` (currently `tools.py:4063-4080`) so the `InboxItem(...)` call includes:

```python
starts_in_minutes=_minutes_until(ev.get("start", {}).get("dateTime")),
```

- [ ] **Step 4: Run tests to verify they pass**

```
uv run pytest tests/test_inbox_ranking.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/tools.py tests/test_inbox_ranking.py
git commit -m "fix(inbox): B3a populate starts_in_minutes on invites and events

Ranker's meeting-proximity tiers (+25/+15/+5) were dead code because
_invite_messages_to_inbox_items and _events_to_inbox_items never set
the signal. Also adds a Graph-tolerant ISO datetime parser that handles
both fractional-second (no tz) and Z-suffixed forms."
```

---

### Task 4: B3 — Populate `flagged` on emails

**Files:**
- Modify: `src/microsoft_mcp/tools.py:1629-1642` (`_list_message_summaries` and `MESSAGE_SUMMARY_SELECT_FIELDS`), `:4004-4025` (`_emails_to_inbox_items`)
- Test: `tests/test_inbox_ranking.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_inbox_ranking.py`:

```python
def test_email_flagged_status_feeds_ranker():
    raw = [{
        "id": "m-1",
        "subject": "Action needed",
        "isRead": True,
        "flag": {"flagStatus": "flagged"},
    }]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].flagged is True
    assert _compute_score(items[0]) == 8.0


def test_email_not_flagged_when_status_missing_or_none():
    raw = [
        {"id": "m-2", "subject": "None", "isRead": True, "flag": {"flagStatus": "notFlagged"}},
        {"id": "m-3", "subject": "Missing", "isRead": True},
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert all(not i.flagged for i in items)
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_inbox_ranking.py::test_email_flagged_status_feeds_ranker -v
```

Expected: FAIL — `flagged` is False.

- [ ] **Step 3: Implement**

Find `MESSAGE_SUMMARY_SELECT_FIELDS` in `tools.py` (grep for it). Add `"flag"` to the comma-separated `$select` string so Graph returns the field.

Update `_emails_to_inbox_items` in `tools.py:4004-4025` to include in the `InboxItem(...)` call:

```python
flagged=(e.get("flag", {}) or {}).get("flagStatus") == "flagged",
```

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_inbox_ranking.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/tools.py tests/test_inbox_ranking.py
git commit -m "fix(inbox): B3b populate flagged signal from message flag status"
```

---

### Task 5: B3 — Populate `is_newsletter` heuristic on emails

**Files:**
- Modify: `src/microsoft_mcp/tools.py` (add helper; update `_emails_to_inbox_items`)
- Test: `tests/test_inbox_ranking.py`

- [ ] **Step 1: Write the failing test**

```python
def test_newsletter_sender_heuristic_flags_item():
    raw = [{
        "id": "m-news",
        "subject": "Weekly digest",
        "isRead": False,
        "from": {"emailAddress": {"address": "noreply@substack.com", "name": "Substack"}},
    }]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].is_newsletter is True
    # unread(+10) plus newsletter(-20) = -10
    assert _compute_score(items[0]) == -10.0


def test_human_sender_not_newsletter():
    raw = [{
        "id": "m-human",
        "subject": "Hey",
        "isRead": False,
        "from": {"emailAddress": {"address": "alice@company.com", "name": "Alice"}},
    }]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].is_newsletter is False
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_inbox_ranking.py::test_newsletter_sender_heuristic_flags_item -v
```

Expected: FAIL.

- [ ] **Step 3: Implement**

Add to `tools.py` above `_emails_to_inbox_items`:

```python
_NEWSLETTER_LOCAL_PARTS = frozenset({
    "noreply", "no-reply", "donotreply", "do-not-reply",
    "newsletter", "news", "digest", "notifications", "notification",
    "updates", "mailer", "hello", "team",
})


def _is_newsletter_sender(raw_from: dict[str, Any] | None) -> bool:
    if not raw_from:
        return False
    address = (raw_from.get("emailAddress") or {}).get("address") or ""
    if "@" not in address:
        return False
    local, _, _domain = address.lower().partition("@")
    if local in _NEWSLETTER_LOCAL_PARTS:
        return True
    # Common patterns like "marketing+id@domain", "news-updates@domain"
    for token in _NEWSLETTER_LOCAL_PARTS:
        if local.startswith(token + "-") or local.startswith(token + "+") or local.startswith(token + "."):
            return True
    return False
```

Update `_emails_to_inbox_items` to include:

```python
is_newsletter=_is_newsletter_sender(e.get("from")),
```

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_inbox_ranking.py -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/tools.py tests/test_inbox_ranking.py
git commit -m "fix(inbox): B3c classify newsletter senders via local-part heuristic"
```

---

### Task 6: B3 — Populate `mentioned` on emails from `mentionsPreview`

**Files:**
- Modify: `src/microsoft_mcp/tools.py` (`MESSAGE_SUMMARY_SELECT_FIELDS`, `_emails_to_inbox_items`)
- Test: `tests/test_inbox_ranking.py`

- [ ] **Step 1: Write the failing test**

```python
def test_mentioned_signal_fires_when_mentionspreview_present():
    raw = [{
        "id": "m-ment",
        "subject": "FYI",
        "isRead": True,
        "mentionsPreview": {"isMentioned": True},
    }]
    items = tools_mod._emails_to_inbox_items(raw)
    assert items[0].mentioned is True
    # mentioned (+15) only, not unread
    assert _compute_score(items[0]) == 15.0


def test_not_mentioned_when_field_absent_or_false():
    raw = [
        {"id": "m-nm1", "subject": "a", "isRead": True},
        {"id": "m-nm2", "subject": "b", "isRead": True, "mentionsPreview": {"isMentioned": False}},
    ]
    items = tools_mod._emails_to_inbox_items(raw)
    assert all(not i.mentioned for i in items)
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_inbox_ranking.py::test_mentioned_signal_fires_when_mentionspreview_present -v
```

- [ ] **Step 3: Implement**

Add `"mentionsPreview"` to `MESSAGE_SUMMARY_SELECT_FIELDS`. Update `_emails_to_inbox_items`:

```python
mentioned=bool((e.get("mentionsPreview") or {}).get("isMentioned")),
```

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_inbox_ranking.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/tools.py tests/test_inbox_ranking.py
git commit -m "fix(inbox): B3d surface @mentions signal from Graph mentionsPreview"
```

---

### Task 7: B4 — Fix `action_hints` read in the inbox-triage example

**Files:**
- Modify: `examples/code-mode/inbox_triage.py:33-52`
- Test: none (example file; add a smoke test only if time permits — see Pass 5)

- [ ] **Step 1: Write the failing test**

In `tests/test_docs_contract.py` (or a new `tests/test_examples.py` if preferred) add:

```python
import pathlib
import re


def test_inbox_triage_example_reads_action_hints_from_summary_items():
    src = pathlib.Path("examples/code-mode/inbox_triage.py").read_text()
    # action_hints is a summary field; must not be read off the hydrated detail.
    assert "detail[\"action_hints\"]" not in src
    assert "detail['action_hints']" not in src
    # The summary-item variable is named `top_items` in the example.
    assert "top_items" in src and "action_hints" in src
    # Sanity: there must be a mapping from summary item to its first action hint.
    assert re.search(r"action_hints", src) is not None
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_docs_contract.py -k inbox_triage -v
```

Expected: FAIL — `detail["action_hints"]` is present.

- [ ] **Step 3: Implement**

Replace the `call_tool_chain` code block in `examples/code-mode/inbox_triage.py` (lines 33–52) with:

```python
    report = await mcp.call_tool_chain(
        """
summary = microsoft.list_inbox_items({"limit": 20})
top_items = summary["items"][:3]

hints = []
for item in top_items:
    action_hints = item.get("action_hints") or []
    hints.append(action_hints[0] if action_hints else "review")

details = []
for item in top_items:
    details.append(
        microsoft.get_inbox_item_detail({"item_id": item["id"], "kind": item["kind"]})
    )

return {
    "titles": [item["title"] for item in top_items],
    "actions": hints,
    "scores": [item["score"] for item in top_items],
}
""",
        timeout=30,
    )
```

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_docs_contract.py -v
```

- [ ] **Step 5: Commit**

```
git add examples/code-mode/inbox_triage.py tests/test_docs_contract.py
git commit -m "fix(examples): B4 read action_hints off inbox summary, not detail

get_inbox_item_detail does not include action_hints; reading it there
always falls through to 'review'. Move the hint lookup to the list
response and keep detail hydration purely for body/subject/location."
```

---

### Task 8: B9 — Dedupe `SCOPES` and freeze it

**Files:**
- Modify: `src/microsoft_mcp/auth.py:63-77`
- Test: `tests/test_auth.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_auth.py`:

```python
from microsoft_mcp.auth import SCOPES


def test_scopes_has_no_duplicates():
    assert len(SCOPES) == len(set(SCOPES))


def test_scopes_contains_required_delegated_permissions():
    for required in ("User.Read", "Mail.Read", "Calendars.Read", "Files.Read"):
        assert required in SCOPES
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_auth.py::test_scopes_has_no_duplicates -v
```

Expected: FAIL — `Chat.Read` appears twice.

- [ ] **Step 3: Implement**

In `src/microsoft_mcp/auth.py`, replace the `SCOPES = [...]` block (lines 63–77) with:

```python
SCOPES: list[str] = sorted({
    "User.Read",
    "User.ReadBasic.All",
    "Chat.Read",
    "Mail.Read",
    "Team.ReadBasic.All",
    "TeamMember.ReadWrite.All",
    "Calendars.Read",
    "Files.Read",
    "ChannelMessage.Read.All",
    "Sites.Read.All",
    "Files.Read.All",
})
```

Delete the commented-out "Scopes useful for full search:" block (lines 79–89) — it's drift bait.

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_auth.py -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/auth.py tests/test_auth.py
git commit -m "fix(auth): B9 dedupe SCOPES via set literal and drop stale comment"
```

---

### Task 9: B13 — Remove MSAL duplicate `account_identifier` assignment

**Files:**
- Modify: `src/microsoft_mcp/auth_msal.py:128-189`
- Test: `tests/test_auth_msal.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_auth_msal.py`:

```python
import inspect
from microsoft_mcp.auth_msal import MSALRefreshTokenAuth


def test_init_assigns_account_identifier_exactly_once():
    source = inspect.getsource(MSALRefreshTokenAuth.__init__)
    # Exactly one `self.account_identifier = ...` assignment.
    occurrences = source.count("self.account_identifier =")
    assert occurrences == 1, f"expected 1 assignment, found {occurrences}"


def test_init_uses_default_when_no_identifier_given(tmp_path):
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, client_id="test-cid")
    assert auth.account_identifier == "default"


def test_init_preserves_explicit_identifier(tmp_path):
    auth = MSALRefreshTokenAuth(
        tokens_dir=tmp_path, client_id="test-cid", account_identifier="user@example.com"
    )
    assert auth.account_identifier == "user@example.com"
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_auth_msal.py::test_init_assigns_account_identifier_exactly_once -v
```

Expected: FAIL — two assignments present.

- [ ] **Step 3: Implement**

In `src/microsoft_mcp/auth_msal.py`, delete the duplicate assignment at line 177 (`self.account_identifier = account_identifier or "default"`). Keep the earlier one at line 165. The `_load_outlook_creds_account_metadata` lookup at lines 166–175 must remain after the first assignment (it depends on `self.account_identifier`).

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_auth_msal.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/auth_msal.py tests/test_auth_msal.py
git commit -m "chore(auth-msal): B13 drop duplicate account_identifier assignment"
```

---

## Pass 2 — Convention drift

### Task 10: A1 — Add `response_profile` to the 13 missing list/search tools

`list_emails`, `list_events`, `list_contacts`, `list_chat_messages` already support it. The other 13 do not: `list_mail_folders`, `list_master_categories`, `list_invite_messages`, `list_files`, `unified_search`, `search_files`, `search_emails`, `search_events`, `search_contacts`, `list_channel_messages`, `search_chat_messages`, `search_channel_messages`, `list_inbox_items`.

The add is mechanical per tool:
1. Add `response_profile: str = "auto"` parameter.
2. Add a `profile = get_response_profile(response_profile)` line near the top of the function body.
3. Under `assistant` profile, switch to the `shape_*_summary` variant (or drop non-essential fields for tools without a shaper).

**Files:**
- Modify: `src/microsoft_mcp/tools.py` (13 tool functions, line numbers below)
- Test: `tests/test_response_shaping.py`

Line references from the audit:
- `list_mail_folders` `:695`, `list_master_categories` `:866`, `list_invite_messages` `:1932`, `list_files` `:2417`, `unified_search` `:2684`, `search_files` `:3081`, `search_emails` `:3135`, `search_events` `:3201`, `search_contacts` `:3246`, `list_channel_messages` `:3450`, `search_chat_messages` `:3785`, `search_channel_messages` `:3855`, `list_inbox_items` `:4084`.

- [ ] **Step 1: Write the failing contract test**

Append to `tests/test_response_shaping.py`:

```python
import inspect
from microsoft_mcp import tools as tools_mod

LIST_OR_SEARCH_TOOLS_THAT_MUST_ACCEPT_PROFILE = [
    "list_emails",
    "list_events",
    "list_contacts",
    "list_chat_messages",
    "list_mail_folders",
    "list_master_categories",
    "list_invite_messages",
    "list_files",
    "unified_search",
    "search_files",
    "search_emails",
    "search_events",
    "search_contacts",
    "list_channel_messages",
    "search_chat_messages",
    "search_channel_messages",
    "list_inbox_items",
]


def test_all_list_and_search_tools_accept_response_profile():
    missing = []
    for name in LIST_OR_SEARCH_TOOLS_THAT_MUST_ACCEPT_PROFILE:
        tool = getattr(tools_mod, name, None)
        assert tool is not None, f"{name} not exported"
        fn = getattr(tool, "fn", tool)
        sig = inspect.signature(fn)
        if "response_profile" not in sig.parameters:
            missing.append(name)
    assert not missing, f"tools missing response_profile param: {missing}"
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_response_shaping.py::test_all_list_and_search_tools_accept_response_profile -v
```

Expected: FAIL — lists the 13 offenders.

- [ ] **Step 3: Implement — template**

For each of the 13 tools, apply this mechanical transform. Template, shown for `list_mail_folders`:

```python
@mcp.tool
def list_mail_folders(
    parent_folder: Any = None,
    recursive: bool = False,
    include_hidden: bool = False,
    limit: int = 100,
    response_profile: str = "auto",
) -> list[dict[str, Any]]:
    """..."""
    profile = get_response_profile(response_profile)
    # existing body...
    folders = _fetch_mail_folders(...)
    if profile == "assistant":
        return [
            {"id": f["id"], "name": f["displayName"], "unread": f.get("unreadItemCount", 0)}
            for f in folders
        ]
    return folders
```

Apply across all 13 tools. Where a summary shaper already exists (`shape_email_summary`, `shape_event_summary`, `shape_contact_summary`), prefer it under `assistant` profile. Tools without a shaper (`list_mail_folders`, `list_master_categories`, `list_files`) should return a minimal dict — id + display_name + the one or two counters the assistant usually needs.

Update each tool's docstring `Args:` block to document the new parameter (use the exact wording from `list_emails` at `tools.py:1116` so it's uniform).

Preserve the `"auto"` default — `get_response_profile` resolves it against `MICROSOFT_MCP_RESPONSE_PROFILE`.

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_response_shaping.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/tools.py tests/test_response_shaping.py
git commit -m "feat(shaping): A1 add response_profile to 13 list/search tools

Uniform support for legacy and assistant profiles across all list/search
endpoints. A contract test now enforces the parameter presence so new
tools cannot skip it."
```

---

### Task 11: A2 — Remove unused `ResponseProfile` enum and `BudgetHints`

**Files:**
- Modify: `src/microsoft_mcp/response_shaping.py:10-27`
- Test: `tests/test_response_shaping.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_response_shaping.py`:

```python
import microsoft_mcp.response_shaping as rs


def test_response_shaping_does_not_export_dead_types():
    assert not hasattr(rs, "ResponseProfile"), \
        "ResponseProfile enum was unused and should be removed"
    assert not hasattr(rs, "BudgetHints"), \
        "BudgetHints dataclass was unused and should be removed"
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_response_shaping.py::test_response_shaping_does_not_export_dead_types -v
```

- [ ] **Step 3: Implement**

In `src/microsoft_mcp/response_shaping.py`, delete lines 3–4 (the unused `Enum` and `dataclass` imports if nothing else needs them — verify first with `rg 'Enum|dataclass' src/microsoft_mcp/response_shaping.py`) and delete the `ResponseProfile` class + `BudgetHints` class at lines 10–27.

Grep the repo for any surviving imports:

```
rg -n 'ResponseProfile|BudgetHints' src/ tests/
```

Any hits must be removed.

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/response_shaping.py tests/test_response_shaping.py
git commit -m "refactor(shaping): A2 remove unused ResponseProfile enum and BudgetHints

Both types were defined but never read. The active profile system lives
on string tokens ('legacy'/'assistant') routed through get_response_profile;
keeping two disjoint vocabularies invited future drift."
```

---

### Task 12: B5 — Make list tools pre-populate the search_cache

**Files:**
- Modify: `src/microsoft_mcp/tools.py` (`list_emails`, `list_events`, `list_chat_messages`, `list_channel_messages`)
- Modify: `src/microsoft_mcp/search_cache.py` (add a normalize helper if useful)
- Test: new `tests/test_search_cache.py`

- [ ] **Step 1: Write the failing test**

Create `tests/test_search_cache.py`:

```python
from unittest.mock import patch
from microsoft_mcp import tools as tools_mod
from microsoft_mcp.search_cache import get_global_cache


def _reset_cache():
    cache = get_global_cache()
    with cache._lock:
        cache._store.clear()


def test_list_emails_populates_cache(monkeypatch):
    _reset_cache()

    def fake_paginated(path, params=None, limit=None, auth=None):
        yield {
            "id": "m-1",
            "subject": "Project alpha kickoff",
            "bodyPreview": "Details inside",
            "from": {"emailAddress": {"address": "a@example.com", "name": "A"}},
            "receivedDateTime": "2026-04-23T00:00:00Z",
            "isRead": False,
            "conversationId": "c-1",
        }

    monkeypatch.setattr("microsoft_mcp.graph.request_paginated", fake_paginated)

    tools_mod.list_emails.fn(folder="inbox", limit=5)

    hits = get_global_cache().search("alpha", kinds=["message"])
    assert any(h.get("id") == "m-1" for h in hits)


def test_cache_hit_survives_between_calls(monkeypatch):
    _reset_cache()

    def fake_paginated(path, params=None, limit=None, auth=None):
        yield {"id": "m-2", "subject": "Budget review", "bodyPreview": "Q3 budget"}

    monkeypatch.setattr("microsoft_mcp.graph.request_paginated", fake_paginated)
    tools_mod.list_emails.fn(folder="inbox", limit=5)

    hits = get_global_cache().search("budget", kinds=["message"])
    assert len(hits) >= 1
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_search_cache.py -v
```

Expected: FAIL — cache is empty after list_emails.

- [ ] **Step 3: Implement**

Add a small normalizer to `search_cache.py` so all callers agree on shape:

```python
def normalize_for_cache(kind: str, raw: dict[str, Any]) -> dict[str, Any]:
    """Normalize a raw Graph item into the cache's {id, kind, title, snippet} schema."""
    if kind == "message":
        return {
            "id": raw.get("id", ""),
            "kind": "message",
            "title": raw.get("subject") or "",
            "snippet": (raw.get("bodyPreview") or "")[:200],
        }
    if kind == "event":
        return {
            "id": raw.get("id", ""),
            "kind": "event",
            "title": raw.get("subject") or "",
            "snippet": (raw.get("bodyPreview") or "")[:200],
        }
    if kind == "chatMessage":
        body = raw.get("body") or {}
        return {
            "id": raw.get("id", ""),
            "kind": "chatMessage",
            "title": raw.get("subject") or (body.get("content") or "")[:60],
            "snippet": (body.get("content") or "")[:200],
        }
    return {
        "id": raw.get("id", ""),
        "kind": kind,
        "title": raw.get("subject") or raw.get("displayName") or "",
        "snippet": (raw.get("bodyPreview") or "")[:200],
    }
```

In `tools.py`, at the end of `list_emails`'s `try:` block (right before `return results`, after the Graph fetch succeeded), add:

```python
from .search_cache import normalize_for_cache
get_global_cache().store("message", [normalize_for_cache("message", e) for e in raw_emails])
```

Do the same for `list_events` (kind `"event"`), `list_chat_messages` (kind `"chatMessage"`), `list_channel_messages` (kind `"chatMessage"`). Import `normalize_for_cache` at the top of `tools.py` next to `get_global_cache`.

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_search_cache.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/search_cache.py src/microsoft_mcp/tools.py tests/test_search_cache.py
git commit -m "fix(search): B5 list tools now populate the degraded-search cache

Cache docstring promised population by list_emails/list_events/
list_chat_messages/list_channel_messages, but the only writer was
unified_search itself — making the cache useless as a 403/404 fallback.
Wires the promised writers and adds a shared normalizer."
```

---

### Task 13: B16 — Strip Safelinks/Mimecast URLs in email body cleanup

**Files:**
- Modify: `src/microsoft_mcp/response_shaping.py:75-131`
- Test: `tests/test_response_shaping.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_response_shaping.py`:

```python
from microsoft_mcp.response_shaping import _clean_body_text


def test_clean_body_strips_safelinks_wrapper_and_decodes_target():
    wrapped = (
        "See https://eur01.safelinks.protection.outlook.com/?url=https%3A%2F%2Fexample.com%2Fpath"
        "&data=abc for details."
    )
    cleaned = _clean_body_text(wrapped)
    assert "safelinks" not in cleaned.lower()
    assert "https://example.com/path" in cleaned


def test_clean_body_strips_mimecast_url():
    wrapped = "Read https://url.ca1.mimecast.com/v1/token now."
    cleaned = _clean_body_text(wrapped)
    assert "mimecast" not in cleaned.lower()
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_response_shaping.py -k clean_body -v
```

- [ ] **Step 3: Implement**

In `src/microsoft_mcp/response_shaping.py`, modify `_clean_body_text` (currently at lines 129–131) to invoke the two unused regexes:

```python
from urllib.parse import unquote  # add to the import block at the top if not present


def _unwrap_safelinks(text: str) -> str:
    def _replace(match: re.Match[str]) -> str:
        return unquote(match.group(1))
    return _SAFE_LINK_RE.sub(_replace, text)


def _strip_mimecast(text: str) -> str:
    # Drop the wrapper URL entirely — there is no recoverable target in the link.
    return _MIMECAST_RE.sub("", text)


def _clean_body_text(text: str) -> str:
    text = _SECURITY_BANNER_RE.sub("", text)
    text = _unwrap_safelinks(text)
    text = _strip_mimecast(text)
    return text.strip()
```

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_response_shaping.py -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/response_shaping.py tests/test_response_shaping.py
git commit -m "fix(shaping): B16 wire Safelinks/Mimecast regexes into body cleanup

Both regexes were defined at module scope but never invoked, so detail
email bodies still contained wrapped URLs. _unwrap_safelinks decodes the
actual target; _strip_mimecast removes the wrapper entirely since the
link payload is opaque."
```

---

## Pass 3 — Robustness

### Task 14: B6 — Remove the double-retry path in `graph.request`

**Files:**
- Modify: `src/microsoft_mcp/graph.py:93-132`
- Test: `tests/test_graph.py`

- [ ] **Step 1: Write the failing test**

Append to `tests/test_graph.py`:

```python
from unittest.mock import patch, MagicMock
import httpx
from microsoft_mcp import graph


def test_request_retries_exactly_max_retries_plus_one_on_persistent_500(mock_auth):
    graph.set_auth_instance(mock_auth)

    call_count = {"n": 0}

    def fake_request(method, url, headers=None, params=None, json=None, content=None):
        call_count["n"] += 1
        resp = MagicMock(spec=httpx.Response)
        resp.status_code = 500
        resp.headers = {}
        resp.content = b""
        resp.raise_for_status.side_effect = httpx.HTTPStatusError(
            "500", request=MagicMock(), response=resp
        )
        return resp

    with patch.object(graph._client, "request", side_effect=fake_request), \
         patch("time.sleep", return_value=None):
        with pytest.raises(httpx.HTTPStatusError):
            graph.request("GET", "/me", max_retries=3)

    # Initial attempt + 3 retries = 4 total. No more.
    assert call_count["n"] == 4
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_graph.py::test_request_retries_exactly_max_retries_plus_one_on_persistent_500 -v
```

Expected: FAIL — current code produces ~8 attempts because both the inline 5xx branch and the except branch retry.

- [ ] **Step 3: Implement**

Replace the `request` function body (lines 93–132) with a single retry loop:

```python
retry_count = 0
while True:
    response = _client.request(
        method=method,
        url=f"{BASE_URL}{path}",
        headers=headers,
        params=params,
        json=json,
        content=data,
    )

    if response.status_code == 429 and retry_count < max_retries:
        retry_after_header = response.headers.get("Retry-After", "5")
        try:
            retry_after = int(retry_after_header)
        except ValueError:
            retry_after = 5
        time.sleep(min(retry_after, 60))
        retry_count += 1
        continue

    if response.status_code >= 500 and retry_count < max_retries:
        time.sleep((2 ** retry_count) * 1)
        retry_count += 1
        continue

    response.raise_for_status()

    if response.content:
        return response.json()
    return None
```

Remove the `except httpx.HTTPStatusError` retry block entirely — `raise_for_status` is only reached after retries are exhausted, and it should just bubble up.

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_graph.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/graph.py tests/test_graph.py
git commit -m "fix(graph): B6 collapse double retry path into single loop

request() previously retried 5xx both inline and in its own except
handler, leading to ~2x the intended attempts. Also hardens Retry-After
parse against non-integer values (HTTP-date form)."
```

---

### Task 15: B7 — Defensive `params` copy in `graph.request`

**Files:**
- Modify: `src/microsoft_mcp/graph.py:61-91`
- Test: `tests/test_graph.py`

- [ ] **Step 1: Write the failing test**

```python
def test_request_does_not_mutate_caller_params(mock_auth, monkeypatch):
    graph.set_auth_instance(mock_auth)

    captured = {}

    def fake_request(method, url, headers=None, params=None, json=None, content=None):
        captured["params"] = dict(params) if params else {}
        resp = MagicMock(spec=httpx.Response)
        resp.status_code = 200
        resp.content = b"{}"
        resp.json.return_value = {}
        resp.raise_for_status = lambda: None
        return resp

    monkeypatch.setattr(graph._client, "request", fake_request)

    user_params = {"$search": "foo"}
    before = dict(user_params)
    graph.request("GET", "/me/messages", params=user_params)

    # The caller's dict must be unchanged.
    assert user_params == before
    # But the request must still have gone out with $count=true.
    assert captured["params"].get("$count") == "true"
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_graph.py::test_request_does_not_mutate_caller_params -v
```

- [ ] **Step 3: Implement**

In `graph.request`, add as the first line of the function body (before the `auth_instance = ...` line):

```python
params = dict(params) if params else None
```

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_graph.py -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/graph.py tests/test_graph.py
git commit -m "fix(graph): B7 defensive copy of caller params to prevent leak"
```

---

### Task 16: B8 — Narrow Azure auth cache-wipe to `ClientAuthenticationError`

**Files:**
- Modify: `src/microsoft_mcp/auth.py:227-317`
- Test: `tests/test_auth.py`

- [ ] **Step 1: Write the failing test**

```python
from unittest.mock import MagicMock
import pytest
from azure.core.exceptions import ClientAuthenticationError, ServiceRequestError
from microsoft_mcp.auth import AzureAuthentication


def test_transient_service_error_does_not_clear_auth_record(tmp_path, monkeypatch):
    record_file = tmp_path / "auth.json"
    record_file.write_text("{\"version\": \"1.0\"}")

    auth = AzureAuthentication(auth_record_file=record_file)

    fake_cred = MagicMock()
    fake_cred.get_token.side_effect = ServiceRequestError("DNS blip")
    monkeypatch.setattr(auth, "get_credential", lambda: fake_cred)

    with pytest.raises(ServiceRequestError):
        auth.get_token()

    # Record must survive the transient failure.
    assert record_file.exists()


def test_client_auth_error_clears_record(tmp_path, monkeypatch):
    record_file = tmp_path / "auth.json"
    record_file.write_text("{\"version\": \"1.0\"}")

    auth = AzureAuthentication(auth_record_file=record_file)

    fake_cred = MagicMock()
    fake_cred.get_token.side_effect = ClientAuthenticationError("invalid_grant")
    monkeypatch.setattr(auth, "get_credential", lambda: fake_cred)

    with pytest.raises(Exception):
        auth.get_token()

    # invalid_grant is terminal; cache should be cleared for re-auth.
    assert not record_file.exists()
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_auth.py::test_transient_service_error_does_not_clear_auth_record -v
```

Expected: FAIL — record is wiped.

- [ ] **Step 3: Implement**

At the top of `auth.py`, import:

```python
from azure.core.exceptions import ClientAuthenticationError
```

Replace the `except Exception as e` branch of both `get_token` (lines 285–317) and `get_token_with_details` (lines 239–271) with:

```python
except ClientAuthenticationError as e:
    logger.error(f"Client authentication failed (terminal): {e}")
    if not self.auth_record_file.exists():
        logger.info("No AuthenticationRecord found, attempting interactive authentication")
        self.authenticate()
        token: AccessToken = credential.get_token(*SCOPES)
        return token.token  # or (token.token, token.expires_on) for the _with_details variant
    logger.info("Clearing cached data after terminal auth failure")
    self.clear_cache()
    self._credential_instance = None
    raise RuntimeError(f"Client authentication failed: {e}") from e
except Exception as e:
    logger.error(f"Transient token acquisition failure: {e}")
    raise
```

Note the key change: transient (non-`ClientAuthenticationError`) failures propagate without clearing cache.

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_auth.py -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/auth.py tests/test_auth.py
git commit -m "fix(auth): B8 stop wiping auth record on transient token failures

Any failure (network blip, DNS hiccup) previously cleared
~/.ms-graph-mcp-azure-auth-record.json, forcing a full interactive
re-auth. Now only terminal ClientAuthenticationError wipes state."
```

---

### Task 17: B10 — Preserve original scopes on MSAL refresh

**Files:**
- Modify: `src/microsoft_mcp/auth_msal.py:339-396`
- Test: `tests/test_auth_msal.py`

- [ ] **Step 1: Write the failing test**

```python
def test_msal_refresh_preserves_saved_scopes(tmp_path, monkeypatch):
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    (tmp_path / "x@y.com_refresh_only.txt").write_text("rt-123")
    (tmp_path / "x@y.com_access_token.json").write_text(
        '{"email": "x@y.com", "access_token": "old", "token_type": "Bearer", '
        '"expires_in": 0, "expires_at": "2020-01-01T00:00:00Z", '
        '"refreshed_at": "2020-01-01T00:00:00Z", '
        '"scopes": "Mail.Read Files.Read offline_access", "api_type": "graph"}'
    )

    captured_scope = {}

    def fake_urlopen(req, timeout=30):
        import io, json as _json
        body = req.data.decode()
        from urllib.parse import parse_qs
        params = parse_qs(body)
        captured_scope["scope"] = params["scope"][0]
        return io.BytesIO(_json.dumps({
            "access_token": "new", "refresh_token": "rt-123",
            "expires_in": 3600, "scope": params["scope"][0],
        }).encode())

    monkeypatch.setattr("microsoft_mcp.auth_msal.urllib.request.urlopen",
                        lambda req, timeout=30: fake_urlopen(req, timeout))

    auth.get_token()
    # Scopes from the cached access_token.json must be reused.
    assert "Mail.Read" in captured_scope["scope"]
    assert "Files.Read" in captured_scope["scope"]
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_auth_msal.py::test_msal_refresh_preserves_saved_scopes -v
```

Expected: FAIL — current refresh hardcodes `.default offline_access`.

- [ ] **Step 3: Implement**

In `_refresh_access_token` (lines 339–396), change the `scopes` local from:

```python
scopes = "https://graph.microsoft.com/.default offline_access"
```

to:

```python
saved = self._load_access_token_data() or {}
saved_scopes = saved.get("scopes") or ""
if saved_scopes:
    scopes = saved_scopes
    if "offline_access" not in scopes.split():
        scopes = scopes + " offline_access"
else:
    scopes = "https://graph.microsoft.com/.default offline_access"
```

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_auth_msal.py -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/auth_msal.py tests/test_auth_msal.py
git commit -m "fix(auth-msal): B10 preserve saved scopes on token refresh

Refresh was hardcoded to '.default offline_access' regardless of the
scope set the original token was issued under. Now reads the saved
scope string back from access_token.json and only falls back to
.default when nothing is cached."
```

---

### Task 18: B11 — Tolerant ISO parse for MSAL expiry

**Files:**
- Modify: `src/microsoft_mcp/auth_msal.py:306-337` (`_is_token_valid`)
- Test: `tests/test_auth_msal.py`

- [ ] **Step 1: Write the failing test**

```python
def test_is_token_valid_accepts_microseconds_and_offset(tmp_path):
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    now = datetime.now(timezone.utc) + timedelta(hours=1)
    for fmt in ["%Y-%m-%dT%H:%M:%S.%fZ", "%Y-%m-%dT%H:%M:%S+00:00"]:
        data = {
            "access_token": "tok",
            "expires_at": now.strftime(fmt),
        }
        path = tmp_path / "x@y.com_access_token.json"
        path.write_text(json.dumps(data))
        assert auth._is_token_valid(), f"failed on format {fmt}"
```

(Add `from datetime import datetime, timedelta, timezone` and `import json` at the top of the test file if not present.)

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_auth_msal.py::test_is_token_valid_accepts_microseconds_and_offset -v
```

- [ ] **Step 3: Implement**

Replace the `try:` block inside `_is_token_valid` (lines 320–337) with:

```python
try:
    raw = expires_at_str.replace("Z", "+00:00")
    expires_at = datetime.fromisoformat(raw)
    if expires_at.tzinfo is None:
        expires_at = expires_at.replace(tzinfo=timezone.utc)
    now = datetime.now(timezone.utc)
    remaining = (expires_at - now).total_seconds()
    if remaining > TOKEN_EXPIRY_BUFFER_SECONDS:
        logger.info(f"Token valid for {remaining:.0f} more seconds")
        return True
    logger.info(f"Token expired or expiring soon ({remaining:.0f}s remaining)")
    return False
except ValueError as e:
    logger.warning(f"Error parsing token expiration: {e}")
    return False
```

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_auth_msal.py -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/auth_msal.py tests/test_auth_msal.py
git commit -m "fix(auth-msal): B11 tolerate ISO variants in expires_at parse

datetime.strptime(%Y-%m-%dT%H:%M:%SZ) rejected microseconds and offset
forms. Swaps to fromisoformat with a Z->+00:00 normalization so cached
tokens written by external tools (outlook-creds) still parse."
```

---

### Task 19: B12 — Lock around MSAL token refresh

**Files:**
- Modify: `src/microsoft_mcp/auth_msal.py:128-570`
- Test: `tests/test_auth_msal.py`

- [ ] **Step 1: Write the failing test**

```python
import threading


def test_concurrent_get_token_refreshes_exactly_once(tmp_path, monkeypatch):
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path, account_identifier="x@y.com")
    (tmp_path / "x@y.com_refresh_only.txt").write_text("rt-1")
    past = datetime.now(timezone.utc) - timedelta(minutes=5)
    (tmp_path / "x@y.com_access_token.json").write_text(json.dumps({
        "access_token": "stale", "expires_at": past.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "scopes": "Mail.Read offline_access",
    }))

    calls = {"n": 0}
    lock = threading.Lock()

    def fake_refresh(refresh_token):
        with lock:
            calls["n"] += 1
        return {"access_token": "fresh", "refresh_token": refresh_token,
                "expires_in": 3600, "scope": "Mail.Read offline_access"}

    monkeypatch.setattr(auth, "_refresh_access_token", fake_refresh)

    tokens = []
    def worker():
        tokens.append(auth.get_token())

    threads = [threading.Thread(target=worker) for _ in range(8)]
    for t in threads: t.start()
    for t in threads: t.join()

    assert calls["n"] == 1, f"expected 1 refresh, got {calls['n']}"
    assert all(t == "fresh" for t in tokens)
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_auth_msal.py::test_concurrent_get_token_refreshes_exactly_once -v
```

Expected: FAIL — multiple refreshes happen.

- [ ] **Step 3: Implement**

In `MSALRefreshTokenAuth.__init__`, add after `self._msal_app = None`:

```python
self._refresh_lock = threading.Lock()
```

(Add `import threading` at module scope.)

Rewrap `get_token` (lines 523–570) with double-checked locking:

```python
def get_token(self) -> str:
    logger.info("Getting access token")
    if self._is_token_valid():
        token_data = self._load_access_token_data()
        if token_data and token_data.get("access_token"):
            return token_data["access_token"]

    with self._refresh_lock:
        # Re-check after acquiring the lock — another thread may have refreshed.
        if self._is_token_valid():
            token_data = self._load_access_token_data()
            if token_data and token_data.get("access_token"):
                return token_data["access_token"]

        refresh_token = self._load_refresh_token()
        if not refresh_token:
            logger.error("No refresh token found. Authentication required.")
            raise Exception(
                "No refresh token found. Run authentication first: "
                "MICROSOFT_MCP_AUTH_METHOD=msal uv run authenticate.py"
            )

        try:
            result = self._refresh_access_token(refresh_token)
            self._save_tokens(
                access_token=result["access_token"],
                refresh_token=result.get("refresh_token", refresh_token),
                expires_in=result.get("expires_in", 3600),
                scopes=result.get("scope", "https://graph.microsoft.com/.default"),
            )
            return result["access_token"]
        except Exception as e:
            logger.error(f"Token refresh failed: {e}")
            self.clear_cache()
            raise Exception(
                f"Token refresh failed: {e}. Please re-authenticate: "
                "MICROSOFT_MCP_AUTH_METHOD=msal uv run authenticate.py"
            ) from e
```

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_auth_msal.py -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/auth_msal.py tests/test_auth_msal.py
git commit -m "fix(auth-msal): B12 serialize concurrent token refresh

Refresh was unlocked. Two threads both seeing an expired token would
both POST the same refresh_token; on tenants with single-use refresh
tokens the second call 400s and wipes the user's credentials. Uses
double-checked locking so the common valid-token path stays lock-free."
```

---

### Task 20: B14 — Stop mutating `account_identifier` in `_save_tokens`

**Files:**
- Modify: `src/microsoft_mcp/auth_msal.py:253-305`
- Test: `tests/test_auth_msal.py`

- [ ] **Step 1: Write the failing test**

```python
def test_account_identifier_immutable_after_save(tmp_path):
    auth = MSALRefreshTokenAuth(tokens_dir=tmp_path)
    assert auth.account_identifier == "default"
    auth._save_tokens(
        access_token="t",
        refresh_token="r",
        expires_in=3600,
        scopes="Mail.Read offline_access",
        email="new@example.com",
    )
    # Must remain "default" — path stability for any caller that cached a path.
    assert auth.account_identifier == "default"
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_auth_msal.py::test_account_identifier_immutable_after_save -v
```

- [ ] **Step 3: Implement**

In `_save_tokens` (lines 253–305), delete lines 273–275:

```python
# Update account identifier if email provided
if email and self.account_identifier == "default":
    self.account_identifier = email
```

The `access_token_data["email"]` field already records the authenticated email; no instance-state mutation needed.

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_auth_msal.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/auth_msal.py tests/test_auth_msal.py
git commit -m "fix(auth-msal): B14 stop mutating account_identifier after save

_save_tokens silently rewrote self.account_identifier when email was
discovered, breaking any caller that cached a token-file path. Email
is still preserved in the JSON payload itself."
```

---

## Pass 4 — Hygiene

### Task 21: A3 — Collapse `get_token` / `get_token_with_details` in both auth modules

Both `auth.py:227-317` and `auth_msal.py:523-599` have near-identical copies.

**Files:**
- Modify: `src/microsoft_mcp/auth.py:227-317`
- Modify: `src/microsoft_mcp/auth_msal.py:523-599`
- Test: `tests/test_auth.py`, `tests/test_auth_msal.py`

- [ ] **Step 1: Write the failing test**

```python
# tests/test_auth.py — behavior lock-in
def test_get_token_and_get_token_with_details_share_value(tmp_path, monkeypatch):
    auth = AzureAuthentication(auth_record_file=tmp_path / "ar.json")
    from unittest.mock import MagicMock
    cred = MagicMock()
    cred.get_token.return_value = MagicMock(token="abc", expires_on=9999999)
    monkeypatch.setattr(auth, "get_credential", lambda: cred)
    t = auth.get_token()
    t2, exp = auth.get_token_with_details()
    assert t == t2 == "abc"
    assert exp == 9999999
```

Same for `tests/test_auth_msal.py` against `MSALRefreshTokenAuth`.

- [ ] **Step 2: Run to verify failure**

Tests may already pass — that's fine. The goal is to lock behavior before refactoring.

- [ ] **Step 3: Implement**

In `auth.py`, replace both `get_token` and `get_token_with_details` with one shared helper:

```python
def _acquire_token(self) -> AccessToken:
    """All token acquisition flows through here so retry/cleanup stay uniform."""
    credential = self.get_credential()
    try:
        return credential.get_token(*SCOPES)
    except ClientAuthenticationError as e:
        logger.error(f"Client authentication failed (terminal): {e}")
        if not self.auth_record_file.exists():
            logger.info("No AuthenticationRecord — running interactive auth")
            self.authenticate()
            return credential.get_token(*SCOPES)
        self.clear_cache()
        self._credential_instance = None
        raise RuntimeError(f"Client authentication failed: {e}") from e


def get_token(self) -> str:
    return self._acquire_token().token


def get_token_with_details(self) -> tuple[str, int]:
    token = self._acquire_token()
    return token.token, token.expires_on
```

Same pattern for `MSALRefreshTokenAuth`: extract the lock+check+refresh logic into `_acquire_token()`, have `get_token` and `get_token_with_details` both call it.

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_auth.py tests/test_auth_msal.py -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/auth.py src/microsoft_mcp/auth_msal.py tests/test_auth.py tests/test_auth_msal.py
git commit -m "refactor(auth): A3 collapse duplicate get_token methods"
```

---

### Task 22: A4 — Lock around `_global_auth`

**Files:**
- Modify: `src/microsoft_mcp/graph.py:17-58`
- Test: `tests/test_graph.py`

- [ ] **Step 1: Write the failing test**

```python
def test_set_and_get_auth_instance_is_thread_safe():
    import threading
    from microsoft_mcp import graph as g

    auths = [object() for _ in range(20)]
    seen = []

    def writer(a): g.set_auth_instance(a)
    def reader(): seen.append(g._global_auth)

    threads = [threading.Thread(target=writer, args=(a,)) for a in auths] + \
              [threading.Thread(target=reader) for _ in range(20)]
    for t in threads: t.start()
    for t in threads: t.join()
    # Every observed value must be either None or one of the real auth objects.
    assert all(s is None or s in auths for s in seen)
```

- [ ] **Step 2-4:** Add `import threading` + `_auth_lock = threading.Lock()`; guard both setters and getter.

```python
_auth_lock = threading.Lock()


def set_auth_instance(auth: "AuthProvider") -> None:
    global _global_auth
    with _auth_lock:
        _global_auth = auth


def get_auth_instance() -> "AuthProvider":
    with _auth_lock:
        if _global_auth is not None:
            return _global_auth
    # Slow path outside the lock — construction may do disk I/O.
    # Re-check under the lock before installing.
    constructed = _construct_default_auth()
    with _auth_lock:
        global _global_auth
        if _global_auth is None:
            _global_auth = constructed
        return _global_auth
```

Extract the existing env-dispatch body into `_construct_default_auth`.

- [ ] **Step 5: Commit**

```
git commit -m "refactor(graph): A4 lock around _global_auth setter/getter"
```

---

### Task 23: A5 — Move `load_dotenv()` out of module scope in `auth.py`

**Files:**
- Modify: `src/microsoft_mcp/auth.py:61`
- Test: `tests/test_auth.py`

- [ ] **Step 1: Write the failing test**

```python
def test_importing_auth_module_does_not_mutate_env(clean_env, monkeypatch, tmp_path):
    envfile = tmp_path / ".env"
    envfile.write_text("MICROSOFT_MCP_CLIENT_ID=from-dotenv\n")
    monkeypatch.chdir(tmp_path)
    # Force a fresh import.
    import importlib, microsoft_mcp.auth as m
    importlib.reload(m)
    # Module import must not have side-loaded .env.
    assert "MICROSOFT_MCP_CLIENT_ID" not in os.environ
```

- [ ] **Steps 2-4:** Delete `load_dotenv()` at line 61. Call it once inside `AzureAuthentication.__init__` instead.

- [ ] **Step 5: Commit**

```
git commit -m "refactor(auth): A5 defer load_dotenv() to instance construction"
```

---

### Task 24: A6 — Use a loop-aware runner in the sandbox awaitable path

**Files:**
- Modify: `src/microsoft_mcp/code_mode.py:657-680`
- Test: `tests/test_code_mode_tools.py`

- [ ] **Step 1: Write the failing test**

```python
import asyncio
import pytest


@pytest.mark.asyncio
async def test_async_tool_invocation_from_sandbox(mcp_with_runtime, monkeypatch):
    """Simulate a tool whose .fn returns a coroutine and verify it still works
    when call_tool_chain is invoked while an event loop is running."""
    runtime = mcp_with_runtime

    async def fake_fn(**kwargs):
        return {"echo": kwargs}

    # Inject a fake tool into the registry under a stable name.
    runtime._tool_cache["echo"] = type("T", (), {"fn": staticmethod(fake_fn)})
    runtime._tool_namespace.echo = runtime._make_tool_wrapper("echo")

    result = await runtime.call_tool_chain(
        "return microsoft.echo({\"x\": 1})"
    )
    assert result["result"] == {"echo": {"x": 1}}
```

- [ ] **Step 2: Run to verify failure**

```
uv run pytest tests/test_code_mode_tools.py::test_async_tool_invocation_from_sandbox -v
```

Expected: FAIL or flaky — `asyncio.run` may misbehave depending on the event-loop state when the sandbox thread resolves the coroutine.

- [ ] **Step 3: Implement**

Add at the top of `code_mode.py` next to the other imports:

```python
import concurrent.futures
```

Add a module-level helper above `class CodeModeRuntime`:

```python
def _run_coroutine_sync(coro: Awaitable[Any]) -> Any:
    """Run a coroutine to completion whether or not a loop is already running.

    Mirrors tools._run_async so sandboxed user code can invoke async tools
    regardless of how call_tool_chain itself was driven.
    """
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return asyncio.run(coro)
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        return executor.submit(lambda: asyncio.run(coro)).result()
```

Replace the body of `_make_tool_wrapper.call_tool` (currently `code_mode.py:657-677`) with:

```python
def call_tool(args: dict[str, Any] | None = None, /, **kwargs: Any) -> Any:
    tool = self._tool_cache.get(tool_name)
    if tool is None:
        tool = self._lookup_tool(tool_name)
        self._tool_cache[tool_name] = tool
    payload = dict(args or {})
    payload.update(kwargs)
    if isinstance(self._trace_sink, list):
        self._trace_sink.append({"tool": tool_name, "args": payload})

    result = tool.fn(**payload)
    if inspect.isawaitable(result):
        return _run_coroutine_sync(result)
    return result
```

Remove the now-unused `_await_result` method at `code_mode.py:679-680`.

- [ ] **Step 4: Run to verify pass**

```
uv run pytest tests/test_code_mode_tools.py -v
uv run pytest tests/ -v
```

- [ ] **Step 5: Commit**

```
git add src/microsoft_mcp/code_mode.py tests/test_code_mode_tools.py
git commit -m "refactor(code-mode): A6 loop-aware awaitable handling in tool wrapper

asyncio.run fails if called from a thread that already has a running
loop. Mirrors the detection pattern from tools._run_async so sandboxed
user code can invoke async tools regardless of how call_tool_chain was
driven."
```

---

### Task 25: A7 — Rename `convert_to_markdown` to `_convert_to_markdown`

**Files:**
- Modify: `src/microsoft_mcp/tools.py:529`
- Test: `tests/test_tools_simple.py` (or whichever test currently references it)

- [ ] **Steps 1-4:** Rename. Grep `rg 'convert_to_markdown' src/ tests/` to find callers; rename them too. The function is not `@mcp.tool`-decorated, so no public contract.

- [ ] **Step 5: Commit**

```
git commit -m "refactor(tools): A7 mark convert_to_markdown as private"
```

---

### Task 26: A8 — URL-encode `email` in `get_user_details`

**Files:**
- Modify: `src/microsoft_mcp/tools.py:578`
- Test: `tests/test_tool_contracts.py`

- [ ] **Step 1: Write the failing test**

```python
def test_get_user_details_url_encodes_email(monkeypatch):
    captured = {}
    def fake_request(method, path, **kwargs):
        captured["path"] = path
        return {"id": "x"}
    monkeypatch.setattr("microsoft_mcp.graph.request", fake_request)
    from microsoft_mcp import tools
    tools.get_user_details.fn(email="weird+name@example.com")
    assert "%2B" in captured["path"]  # '+' must be URL-encoded
```

- [ ] **Steps 2-4:** Change `f"/users/{email}"` to `f"/users/{quote(email, safe='')}"`.

- [ ] **Step 5: Commit**

```
git commit -m "fix(tools): A8 URL-encode email path segment in get_user_details"
```

---

### Task 27: A9 — Atomic token-file creation in MSAL

**Files:**
- Modify: `src/microsoft_mcp/auth_msal.py:222-228` (`_secure_write_file`)
- Test: `tests/test_auth_msal.py`

- [ ] **Steps 1-4:** Replace `_secure_write_file` body with:

```python
fd = os.open(str(path), os.O_WRONLY | os.O_CREAT | os.O_TRUNC, TOKEN_FILE_MODE)
with os.fdopen(fd, "w") as f:
    f.write(content)
```

Add a test that creates a token file and verifies the mode is `0o600` via `path.stat().st_mode & 0o777`.

- [ ] **Step 5: Commit**

```
git commit -m "fix(auth-msal): A9 atomic create+mode for token files"
```

---

### Task 28: A10 — `shutil.which("npx")` default in UTCP bridge

**Files:**
- Modify: `src/microsoft_mcp/utcp_bridge_config.py:11-13`
- Test: `tests/test_utcp_bridge_config.py`

- [ ] **Steps 1-4:** Replace the hardcoded path constant:

```python
import shutil

DEFAULT_BRIDGE_COMMAND = (
    os.getenv("MICROSOFT_MCP_UTCP_BRIDGE_COMMAND")
    or shutil.which("npx")
    or "npx"
)
```

Add a test that asserts `DEFAULT_BRIDGE_COMMAND` is not a hardcoded user path. Use a regex like `re.search(r"/(?:Users|home)/[A-Za-z0-9_.-]+/", DEFAULT_BRIDGE_COMMAND)` so the guard catches any contributor's home dir (e.g. `~/…` → `/Users/<name>/…` or `/home/<name>/…`), not just one developer's.

- [ ] **Step 5: Commit**

```
git commit -m "fix(utcp): A10 resolve npx via shutil.which instead of hardcoded path"
```

---

### Task 29: A11 — `raise ... from e` sweep in auth modules

**Files:**
- Modify: `src/microsoft_mcp/auth.py`, `src/microsoft_mcp/auth_msal.py`
- Test: none (lint-style sweep)

- [ ] **Steps 1-4:** Grep `rg -n 'raise Exception\(' src/microsoft_mcp/auth*.py`. For each hit, rewrite as `raise RuntimeError(...) from e` (or a more specific exception class). Verify no test breakage.

- [ ] **Step 5: Commit**

```
git commit -m "refactor(auth): A11 chain exceptions with 'raise ... from e'"
```

---

## Pass 5 — Regression guardrails

### Task 30: Add contract test for response_profile + Graph-call-path coverage

**Files:**
- Create: `tests/test_tool_surface_contract.py`

- [ ] **Step 1: Write the (passing but protective) test**

```python
"""Contract tests enforcing tool-surface audit invariants. These lock in
the fixes from the 2026-04-23 audit so regressions are caught in CI."""

import inspect
import re
from microsoft_mcp import tools as tools_mod


LIST_OR_SEARCH_TOOLS = [
    "list_emails", "list_events", "list_contacts", "list_chat_messages",
    "list_mail_folders", "list_master_categories", "list_invite_messages",
    "list_files", "unified_search", "search_files", "search_emails",
    "search_events", "search_contacts", "list_channel_messages",
    "search_chat_messages", "search_channel_messages", "list_inbox_items",
]


def test_all_list_search_tools_accept_response_profile():
    missing = []
    for name in LIST_OR_SEARCH_TOOLS:
        tool = getattr(tools_mod, name)
        fn = getattr(tool, "fn", tool)
        if "response_profile" not in inspect.signature(fn).parameters:
            missing.append(name)
    assert not missing, missing


def test_no_direct_httpx_calls_in_tools_module():
    src = inspect.getsource(tools_mod)
    # tools.py imports httpx but must not instantiate its own clients or make raw HTTP.
    assert not re.search(r"httpx\.(Async)?Client\(", src)
    assert not re.search(r"httpx\.(get|post|put|delete|patch)\(", src)


def test_inbox_ranker_signals_are_populated():
    # At least one call to each attribute in tools.py body.
    src = inspect.getsource(tools_mod)
    for signal in ("mentioned=", "flagged=", "is_newsletter=", "starts_in_minutes="):
        assert signal in src, f"inbox ranker signal {signal!r} has no populator"


def test_call_tool_chain_default_response_is_lean():
    # Guards B2 regression — the meta-tool should not return the catalog by default.
    tool = tools_mod.call_tool_chain
    fn = getattr(tool, "fn", tool)
    sig = inspect.signature(fn)
    assert "include_interfaces" in sig.parameters
    assert sig.parameters["include_interfaces"].default is False
```

- [ ] **Step 2: Run to verify all pass**

```
uv run pytest tests/test_tool_surface_contract.py -v
```

- [ ] **Step 3-4: n/a** — no code change; this is the regression net.

- [ ] **Step 5: Commit**

```
git add tests/test_tool_surface_contract.py
git commit -m "test(contract): add tool-surface audit regression guards

Locks in the fixes from the 2026-04-23 audit so any future regression
(missing response_profile, raw httpx call, unpopulated ranker signal,
re-introduced envelope bloat) fails CI."
```

---

### Task 31: Wire `/techdebt` into a weekly reminder

**Files:**
- Create: `.claude/commands/weekly-audit.md` (mirrors `/techdebt` but scoped to the full surface)
- Update: `CLAUDE.md` — add a one-line pointer in the "Claude Code Setup" section.

- [ ] **Step 1-4:** Write `.claude/commands/weekly-audit.md`:

```markdown
---
description: Weekly tool-surface audit pass (bugs, response_profile drift, cache integrity)
---

Run `/techdebt src/microsoft_mcp/` and compare against the findings in
`docs/superpowers/plans/2026-04-23-tool-surface-audit.md`. Flag:

1. New list/search tools without `response_profile`.
2. New raw httpx calls.
3. New silent `except Exception: pass` blocks.
4. Tool functions that look public (no leading underscore) but lack `@mcp.tool`.
5. Any `InboxItem(...)` construction that skips `mentioned`/`flagged`/`is_newsletter`/`starts_in_minutes`.

Report as `pass/fail` against each of the five categories. Suggest fixes, do not apply them automatically.
```

Add to `CLAUDE.md` under "Claude Code Setup":

```markdown
Run `/weekly-audit` monthly (or after any sweep of new tools) to guard
against regression of the 2026-04-23 audit findings.
```

- [ ] **Step 5: Commit**

```
git add .claude/commands/weekly-audit.md CLAUDE.md
git commit -m "chore(claude): add weekly-audit slash command for regression guarding"
```

---

## Appendix A — Finding-to-task cross-reference

| ID | Severity | File:line | Task |
|----|----------|-----------|------|
| B1 | bug | code_mode.py:541-616 | Task 1 |
| B2 | bug | code_mode.py:361-369 | Task 2 |
| B3 | bug | tools.py:4004-4080, inbox_ranking.py:8-36 | Tasks 3-6 |
| B4 | bug | examples/code-mode/inbox_triage.py:45 | Task 7 |
| B5 | bug | search_cache.py:1-6, tools.py list tools | Task 12 |
| B6 | bug | graph.py:93-132 | Task 14 |
| B7 | bug | graph.py:85-91 | Task 15 |
| B8 | bug | auth.py:268-271, 313-317 | Task 16 |
| B9 | bug | auth.py:73 | Task 8 |
| B10 | bug | auth_msal.py:355-356 | Task 17 |
| B11 | bug | auth_msal.py:320-321 | Task 18 |
| B12 | bug | auth_msal.py:523-570 | Task 19 |
| B13 | bug | auth_msal.py:165,177 | Task 9 |
| B14 | bug | auth_msal.py:273-275 | Task 20 |
| B15 | minor | response_shaping.py:47,57 | (deferred — no fix planned unless a real case emerges) |
| B16 | bug | response_shaping.py:75-131 | Task 13 |
| A1 | drift | tools.py (13 tools) | Task 10 |
| A2 | drift | response_shaping.py:10-27 | Task 11 |
| A3 | drift | auth.py & auth_msal.py | Task 21 |
| A4 | drift | graph.py:17-58 | Task 22 |
| A5 | drift | auth.py:61 | Task 23 |
| A6 | drift | code_mode.py:673-674 | Task 24 |
| A7 | drift | tools.py:529 | Task 25 |
| A8 | drift | tools.py:578 | Task 26 |
| A9 | drift | auth_msal.py:222-228 | Task 27 |
| A10 | drift | utcp_bridge_config.py:11-13 | Task 28 |
| A11 | drift | auth*.py | Task 29 |

## Appendix B — Pass sequencing rationale

- Pass 1 is pure user-visible wins; ship first for fastest feedback.
- Pass 2 depends on Pass 1 only for momentum. Could ship in parallel from a separate worktree.
- Pass 3 touches auth state — do these in sequence, not parallel, to keep the token-file format stable between commits.
- Pass 4 is safe to parallelize across contributors.
- Pass 5 should land last so the regression guards reflect the world after the fixes, not before.

## Appendix C — What this plan does NOT cover

Deliberately out of scope (documented for future work):
- `search_query` cursor migration from `from`-offset to `moreResultsAvailable`.
- MarkItDown async conversion.
- httpx connection-pool rework.
- B15 (`cleanup_graph_payload` `in _EMPTY`) — kept as fragile-but-correct.
- Comprehensive rewrite of the inbox ranker (could benefit from ML-learned weights; current heuristic tiers are a reasonable baseline).
