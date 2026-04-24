# MCP Response Shaping And Inbox Tools Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Reduce unnecessary tool-response volume before it reaches the AI assistant while adding assistant-native inbox tooling for email, calendar, and Teams workflows.

**Architecture:** Introduce a central response-shaping layer that converts raw Microsoft Graph resources into compact assistant-facing payloads with consistent contracts. Keep existing raw-ish tools for compatibility behind an opt-in response profile, then add higher-level inbox tools that cluster, rank, deduplicate, and selectively hydrate data so the assistant sees only high-signal items.

**Tech Stack:** Python 3.13, FastMCP, httpx, Microsoft Graph, pytest

**Graph request minimization rules:**
- Use `$select` as the primary server-side payload reduction tool on read paths where the endpoint supports projection. This is the correct optimization for mailbox and file listing APIs and should be applied before local response shaping.
- Use `Prefer: outlook.body-content-type="text"` only when a body field is intentionally requested. Do not request `body` in summary/list/search profiles.
- Use `Prefer: return=minimal` only on supported write operations and specific delta endpoints. Do not plan around it for normal `GET` list/detail calls because it does not replace `$select`.
- Treat per-operation Microsoft Graph docs as the source of truth for query-option support. The local `docs/graph_metadata.xml` CSDL file is useful for capability discovery, but it does not reliably advertise `$select` support on the Outlook and Teams collection endpoints used by this server.
- For this codebase specifically:
  - Mail list and message delta flows should use `$select` when projecting summary fields.
  - Calendar list flows can use `$select`, but calendarView delta does not support `$select`.
  - Drive child listing supports `$select`.
  - Chat and channel message listing should not rely on `$select`; reduce volume there with `$top`, targeted container selection, and local shaping.

**Cross-cutting assistant UX rules:**
- Summary/list/search outputs should read like inbox previews, not transport payloads. The shaping layer should optimize for readability and follow-up actions, not field completeness.
- Email shaping must remove common external-mail security banners and disclaimer boilerplate from snippets and summary bodies. In detail mode, keep the cleaned body by default and expose raw content only through `response_profile="raw"`.
- Email shaping should normalize or collapse known wrapped-link formats such as Mimecast and Microsoft Safe Links when the destination URL is recoverable. Summary outputs should never surface 150-character tracking URLs when a canonical destination or short placeholder is available.
- Email detail in `thread_mode="latest"` should strip quoted reply history and signature blocks so the assistant sees the new content first. A full-thread/raw option can preserve the original payload when explicitly requested.
- Permission/auth/search failures should use one normalized error envelope across tools, for example `code`, `message`, `required_scopes`, `remediation`, and `degraded_mode_available`, instead of surfacing raw HTTP exceptions or endpoint-specific error strings.
- Contact shaping should prefer SMTP-style addresses in `email_addresses`. If an Exchange DN cannot be resolved, keep it in `unresolved_addresses` rather than mixing it into the primary email list.

---

## Phase 0: Baseline And Compatibility Guardrails

### Task 1: Add response profile and budget primitives

**Files:**
- Create: `src/microsoft_mcp/response_shaping.py`
- Modify: `src/microsoft_mcp/tools.py:353-2302`
- Test: `tests/test_response_shaping.py`

**Step 1: Write the failing tests**

```python
from microsoft_mcp.response_shaping import ResponseProfile, BudgetHints


def test_response_profile_defaults_to_assistant_summary():
    assert ResponseProfile.default_for_operation("list") == ResponseProfile.SUMMARY


def test_budget_hints_exposes_body_and_item_limits():
    hints = BudgetHints.for_operation("list_emails")
    assert hints.include_body is False
    assert hints.max_items <= 25
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_response_shaping.py -k "response_profile or budget_hints" -v`

Expected: FAIL because `response_shaping.py` does not exist.

**Step 3: Write minimal implementation**

```python
from enum import Enum
from dataclasses import dataclass


class ResponseProfile(str, Enum):
    RAW = "raw"
    DETAIL = "detail"
    SUMMARY = "summary"

    @classmethod
    def default_for_operation(cls, operation: str) -> "ResponseProfile":
        return cls.SUMMARY if operation in {"list", "search"} else cls.DETAIL


@dataclass(frozen=True)
class BudgetHints:
    include_body: bool
    max_items: int

    @classmethod
    def for_operation(cls, tool_name: str) -> "BudgetHints":
        return cls(include_body=False, max_items=25)
```

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_response_shaping.py -k "response_profile or budget_hints" -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_response_shaping.py src/microsoft_mcp/response_shaping.py
git commit -m "feat: add response profile primitives"
```

---

## Phase 1: Central Response Shaping

### Task 2: Implement global Graph payload cleanup helpers

**Files:**
- Modify: `src/microsoft_mcp/response_shaping.py`
- Test: `tests/test_response_shaping.py`

**Step 1: Write the failing tests**

```python
from microsoft_mcp.response_shaping import cleanup_graph_payload


def test_cleanup_graph_payload_strips_odata_and_empty_values():
    raw = {
        "@odata.context": "x",
        "@odata.etag": "y",
        "displayName": "John",
        "mobilePhone": None,
        "otherAddress": {},
        "businessPhones": [],
    }
    assert cleanup_graph_payload(raw) == {"displayName": "John"}


def test_cleanup_graph_payload_keeps_false_and_zero():
    raw = {"isRead": False, "size": 0, "subject": "Test"}
    assert cleanup_graph_payload(raw) == raw
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_response_shaping.py -k cleanup_graph_payload -v`

Expected: FAIL because helper is not implemented.

**Step 3: Write minimal implementation**

```python
ODATA_KEYS = {"@odata.context", "@odata.etag"}
NOISE_KEYS = {"changeKey", "parentFolderId", "calendar@odata.associationLink", "calendar@odata.navigationLink"}


def cleanup_graph_payload(value):
    if isinstance(value, dict):
        cleaned = {}
        for key, child in value.items():
            if key in ODATA_KEYS or key in NOISE_KEYS:
                continue
            next_value = cleanup_graph_payload(child)
            if next_value in (None, "", [], {}):
                continue
            cleaned[key] = next_value
        return cleaned
    if isinstance(value, list):
        return [item for item in (cleanup_graph_payload(v) for v in value) if item not in (None, "", [], {})]
    return value
```

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_response_shaping.py -k cleanup_graph_payload -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_response_shaping.py src/microsoft_mcp/response_shaping.py
git commit -m "feat: add global graph payload cleanup"
```

### Task 3: Add email/event/contact/message-specific shapers

**Files:**
- Modify: `src/microsoft_mcp/response_shaping.py`
- Test: `tests/test_response_shaping.py`

**Step 1: Write the failing tests**

```python
from microsoft_mcp.response_shaping import (
    shape_email_summary,
    shape_event_summary,
    shape_contact_summary,
)


def test_shape_email_summary_drops_body_and_flattens_sender():
    raw = {
        "id": "1",
        "subject": "Hello",
        "from": {"emailAddress": {"name": "JP", "address": "jp@example.com"}},
        "body": {"content": "huge"},
        "conversationId": "abc",
    }
    shaped = shape_email_summary(raw)
    assert shaped == {
        "id": "1",
        "subject": "Hello",
        "from": "JP <jp@example.com>",
        "conversation_url": "https://outlook.office.com/mail/deeplink/readconv/abc",
    }


def test_shape_event_summary_keeps_actionable_fields():
    raw = {
        "id": "evt-1",
        "subject": "AI Pilot Intro",
        "start": {"dateTime": "2026-03-24T18:00:00", "timeZone": "UTC"},
        "location": {"displayName": "Microsoft Teams Meeting", "uniqueId": "huge-bing-url"},
    }
    assert shape_event_summary(raw)["id"] == "evt-1"


def test_shape_contact_summary_filters_empty_email_entries():
    raw = {
        "id": "c-1",
        "displayName": "Brian Roach",
        "emailAddresses": [{"address": "user@example.com"}, {}, {}],
    }
    assert shape_contact_summary(raw)["email_addresses"] == ["user@example.com"]
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_response_shaping.py -k "shape_email_summary or shape_event_summary or shape_contact_summary" -v`

Expected: FAIL

**Step 3: Write minimal implementation**

Implement:
- `flatten_email_address()`
- `compact_location()`
- `extract_teams_meeting_info()`
- `shape_email_summary()`, `shape_email_detail()`
- `shape_event_summary()`, `shape_event_detail()`
- `shape_contact_summary()`, `shape_contact_detail()`
- `shape_message_summary()`, `shape_message_detail()`

Specific rules:
- Drop `body` from list/search outputs by default.
- Flatten nested recipient objects to strings.
- Filter empty email objects and Exchange DN entries into `unresolved_addresses`.
- Convert event/message HTML to markdown before truncation.
- Extract `join_url`, `meeting_id`, `passcode`, `dial_in` from event HTML.

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_response_shaping.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_response_shaping.py src/microsoft_mcp/response_shaping.py
git commit -m "feat: add assistant-facing resource shapers"
```

---

## Phase 2: Refactor Existing Tools To Use Shapers

### Task 4: Refactor auth, user, availability, files, and contacts to consistent compact responses

**Files:**
- Modify: `src/microsoft_mcp/tools.py:93-330`
- Modify: `src/microsoft_mcp/tools.py:671-941`
- Test: `tests/test_tool_contracts.py`

**Step 1: Write the failing tests**

```python
def test_get_user_details_strips_odata_and_nulls():
    result = get_user_details()
    assert "@odata.context" not in result


def test_check_availability_returns_compact_schedule_items():
    result = check_availability("2026-03-23T16:00:00Z", "2026-03-23T17:00:00Z")
    assert "summary" in result
    assert "value" not in result


def test_list_contacts_summary_mode_is_compact():
    results = list_contacts(limit=5)
    assert set(results[0]).issubset({"id", "displayName", "jobTitle", "companyName", "email_addresses", "businessPhones", "mobilePhone"})
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_tool_contracts.py -k "get_user_details or check_availability or list_contacts" -v`

Expected: FAIL on current raw payload shapes.

**Step 3: Write minimal implementation**

Changes:
- `get_user_details()` uses `cleanup_graph_payload()` and returns a compact user dict.
- `check_availability()` returns:
  - `summary`
  - `participants`
  - `slots`
  - `working_hours`
- `list_contacts()` and `get_contact()` delegate to contact shapers.
- `search_contacts()` filters empty objects and unresolved Exchange DN addresses and should only place SMTP-style addresses in `email_addresses`.
- `list_files()` and `search_files()` keep current compact shape but add `kind` and `web_url` if available.

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_tool_contracts.py -k "get_user_details or check_availability or list_contacts" -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_tool_contracts.py src/microsoft_mcp/tools.py
git commit -m "refactor: compact auth contact and availability tool responses"
```

### Task 5: Refactor email tools to summary/detail contracts and safer defaults

**Files:**
- Modify: `src/microsoft_mcp/tools.py:353-557`
- Test: `tests/test_email_tools.py`

**Step 1: Write the failing tests**

```python
def test_list_emails_defaults_to_no_body():
    result = list_emails(limit=10)
    assert "body" not in result[0]


def test_get_email_drops_body_preview_when_body_present():
    result = get_email("message-id")
    assert "bodyPreview" not in result


def test_get_email_supports_latest_message_only_mode():
    result = get_email("message-id", thread_mode="latest")
    assert result["body"]["thread_mode"] == "latest"
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_email_tools.py -v`

Expected: FAIL because defaults and thread controls do not exist.

**Step 3: Write minimal implementation**

Changes:
- Change `list_emails(include_body=False)` default.
- Add `response_profile: str = "summary"` to `list_emails`, `get_email`, `search_emails`.
- Add `thread_mode: str = "latest"` to `get_email`.
- Add explicit summary-field `$select` projections for list and search reads, for example `id,subject,from,toRecipients,receivedDateTime,isRead,hasAttachments,bodyPreview,conversationId,webLink`.
- Apply `shape_email_summary()` to list/search and `shape_email_detail()` to `get_email`.
- Strip common security-warning banners, duplicate `sender`, duplicate `webLink`, `bodyPreview` when body is present, signature/cid noise, and quoted reply history in `thread_mode="latest"`.
- Collapse known wrapped-link formats such as Mimecast and Safe Links to canonical destinations where possible, and otherwise shorten them to readable placeholders in summary/snippet fields.
- Add optional `group_by_conversation: bool = True` to `list_emails`.
- Preserve `Prefer: outlook.body-content-type="text"` only when detail mode requests `body`.

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_email_tools.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_email_tools.py src/microsoft_mcp/tools.py
git commit -m "refactor: make email tools assistant-first"
```

### Task 6: Refactor calendar tools to action-oriented event outputs

**Files:**
- Modify: `src/microsoft_mcp/tools.py:558-670`
- Test: `tests/test_event_tools.py`

**Step 1: Write the failing tests**

```python
def test_get_event_converts_html_to_markdown():
    result = get_event("event-id")
    assert result["body"]["contentType"] == "text/markdown"


def test_get_event_extracts_teams_meeting_info():
    result = get_event("event-id")
    assert "meeting" in result
    assert "join_url" in result["meeting"]


def test_search_events_returns_id_location_and_organizer():
    result = search_events("tour")
    assert {"id", "subject", "location", "organizer"} <= set(result[0])
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_event_tools.py -v`

Expected: FAIL

**Step 3: Write minimal implementation**

Changes:
- `list_events()` delegates to `shape_event_summary()` unless `response_profile="detail"`.
- Add explicit `$select` projections to calendar list reads so summary mode only asks Graph for the event fields it returns.
- `get_event()` converts body HTML, compacts attendees, extracts Teams metadata, removes raw Graph noise.
- `search_events()` must no longer return raw `graph.search_query()` output directly. Map to event summary shape and guarantee `id`, `subject`, `start`, `end`, `location`, `organizer`, `meeting`.
- If delta-based event caching is added later, do not assume `$select` is available on calendarView delta; use date-range scoping and `Prefer: odata.maxpagesize` there instead.

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_event_tools.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_event_tools.py src/microsoft_mcp/tools.py
git commit -m "refactor: shape calendar tools for assistant workflows"
```

---

## Phase 3: Search And Teams Reliability

### Task 7: Replace raw Graph search passthrough with normalized search results

**Files:**
- Modify: `src/microsoft_mcp/graph.py:329-430`
- Modify: `src/microsoft_mcp/tools.py:1079-1690`
- Test: `tests/test_search_tools.py`

**Step 1: Write the failing tests**

```python
def test_unified_search_defaults_to_inbox_entities():
    result = unified_search("AI pilot")
    assert result["summary"]["entity_types_searched"] == ["message", "event", "chatMessage"]


def test_unified_search_results_are_normalized():
    result = unified_search("budget", entity_types=["message"])
    assert set(result["results"][0]) >= {"id", "kind", "title", "snippet", "score"}


def test_search_emails_and_search_events_share_common_contract():
    email = search_emails("meeting", limit=1)[0]
    event = search_events("meeting", limit=1)[0]
    assert {"id", "kind", "title"} <= set(email)
    assert {"id", "kind", "title"} <= set(event)
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_search_tools.py -v`

Expected: FAIL due to raw resource passthrough and file-first defaults.

**Step 3: Write minimal implementation**

Changes:
- `graph.search_query()` accepts explicit stored fields per entity type instead of one generic list.
- `unified_search()` defaults to inbox-first entities when no entity types are supplied.
- `_process_search_hit()` no longer returns raw resource copies. It should:
  - normalize to a shared contract
  - attach `kind`
  - include `score`
  - include `snippet`
  - avoid `body` unless explicit detail hydration requested
- `search_emails()` and `search_events()` reuse the same shaping contract.
- Any permission or capability failure should return the shared normalized error envelope with likely missing scopes and remediation guidance.

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_search_tools.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_search_tools.py src/microsoft_mcp/graph.py src/microsoft_mcp/tools.py
git commit -m "refactor: normalize search outputs and defaults"
```

### Task 8: Add capability-aware degraded search fallback

**Files:**
- Create: `src/microsoft_mcp/search_cache.py`
- Modify: `src/microsoft_mcp/tools.py:1079-2302`
- Test: `tests/test_search_fallback.py`

**Step 1: Write the failing tests**

```python
def test_unified_search_returns_degraded_mode_when_graph_search_is_forbidden():
    result = unified_search("AI pilot", entity_types=["message"])
    assert result["summary"]["mode"] in {"graph_search", "degraded_cache_search"}


def test_search_chat_messages_uses_cache_fallback_on_403():
    result = search_chat_messages("budget", limit=5)
    assert "degraded" in result["meta"]
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_search_fallback.py -v`

Expected: FAIL because no cache fallback exists.

**Step 3: Write minimal implementation**

Implement an in-memory TTL cache of normalized recent items populated by:
- `list_emails`
- `list_events`
- `list_chat_messages`
- `list_channel_messages`

Fallback behavior:
- On Graph Search 403/404, search the normalized cache.
- Return `meta.mode`, `meta.data_freshness`, and `meta.degraded_reason`.
- Never fall back to raw all-item scans from inside search handlers.
- If fallback is unavailable, still return the shared normalized error envelope rather than a raw HTTP exception string.
- If the cache later gains delta refreshers, use message delta with `$select` for mail and avoid assuming calendarView delta supports `$select`.

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_search_fallback.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_search_fallback.py src/microsoft_mcp/search_cache.py src/microsoft_mcp/tools.py
git commit -m "feat: add degraded cache search fallback"
```

### Task 9: Stop brute-force Teams fan-out and replace it with bounded recency collection

**Files:**
- Modify: `src/microsoft_mcp/tools.py:1692-2302`
- Test: `tests/test_teams_tools.py`

**Step 1: Write the failing tests**

```python
def test_list_chat_messages_does_not_scan_every_chat_when_limit_is_small():
    result = list_chat_messages(limit=10)
    assert result["meta"]["containers_scanned"] <= 10


def test_list_channel_messages_requires_scope_or_target():
    result = list_channel_messages(limit=10)
    assert "meta" in result
    assert result["meta"]["mode"] in {"targeted", "recent-index"}
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_teams_tools.py -v`

Expected: FAIL because current implementation scans all chats and channels.

**Step 3: Write minimal implementation**

Changes:
- Add explicit `chat_id`, `team_id`, `channel_id`, and `recent_container_limit` controls.
- For untargeted calls, fetch only top-N recent chats/teams/channels based on container metadata, not every container.
- Return normalized message summaries only.
- Add capability metadata to permission errors with likely missing scopes, and return it via the shared normalized error envelope used by other tools.
- Do not spend implementation effort on `$select` for chat/channel list endpoints. Those endpoints should be treated as non-projectable reads and optimized through bounded traversal, `$top`, targeting, and local shaping instead.

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_teams_tools.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_teams_tools.py src/microsoft_mcp/tools.py
git commit -m "refactor: bound Teams traversal and normalize outputs"
```

---

## Phase 4: Assistant-Native Inbox Tools

### Task 10: Add a normalized inbox item model and ranking rules

**Files:**
- Create: `src/microsoft_mcp/inbox_models.py`
- Create: `src/microsoft_mcp/inbox_ranking.py`
- Test: `tests/test_inbox_ranking.py`

**Step 1: Write the failing tests**

```python
from microsoft_mcp.inbox_models import InboxItem
from microsoft_mcp.inbox_ranking import rank_items


def test_rank_items_prioritizes_unread_mentions_and_soon_events():
    ranked = rank_items([
        InboxItem(kind="email", title="FYI", unread=False, mentioned=False, urgency=0),
        InboxItem(kind="event", title="Starts soon", starts_in_minutes=10, unread=False, mentioned=False, urgency=0),
    ])
    assert ranked[0].title == "Starts soon"
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_inbox_ranking.py -v`

Expected: FAIL

**Step 3: Write minimal implementation**

Define a shared `InboxItem` dataclass with:
- `id`
- `kind`
- `source_tool`
- `title`
- `snippet`
- `participants`
- `when`
- `state`
- `score`
- `reason`
- `action_hints`
- `web_url`

Implement ranking heuristics for:
- unread
- mentions
- meeting start proximity
- flagged/important
- sender/organizer priority
- newsletter suppression

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_inbox_ranking.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_inbox_ranking.py src/microsoft_mcp/inbox_models.py src/microsoft_mcp/inbox_ranking.py
git commit -m "feat: add inbox item model and ranking"
```

### Task 11: Add assistant-native inbox tools

**Files:**
- Modify: `src/microsoft_mcp/tools.py`
- Test: `tests/test_inbox_tools.py`

**Step 1: Write the failing tests**

```python
def test_list_inbox_items_returns_mixed_ranked_items():
    result = list_inbox_items(limit=20)
    assert "items" in result
    assert "meta" in result
    assert {"id", "kind", "title", "score"} <= set(result["items"][0])


def test_get_inbox_item_detail_hydrates_one_item_only():
    result = get_inbox_item_detail(item_id="x", kind="email")
    assert result["kind"] == "email"
    assert "body" in result
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_inbox_tools.py -v`

Expected: FAIL because tools do not exist.

**Step 3: Write minimal implementation**

Add:
- `list_inbox_items(limit=20, include_kinds=None, response_profile="summary")`
- `search_inbox_items(query, limit=20, include_kinds=None)`
- `get_inbox_item_detail(item_id, kind)`
- `list_actionable_today(limit=20)`

Implementation rules:
- Use normalized shapers and ranking layer.
- Cluster by email conversation, event invite chain, and Teams thread.
- Return only one representative item per cluster in list/search.
- Hydrate details only in `get_inbox_item_detail`.

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_inbox_tools.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_inbox_tools.py src/microsoft_mcp/tools.py
git commit -m "feat: add assistant-native inbox tools"
```

---

## Phase 5: Code Mode And Documentation

### Task 12: Add Code Mode guidance for orchestration, not server-side shaping

**Files:**
- Create: `docs/code-mode-inbox-orchestration.md`
- Create: `examples/code-mode/inbox_triage.ts`
- Modify: `README.md`
- Modify: `IMPLEMENTATION.md`
- Test: `tests/README.md` (if example docs are referenced there)

**Step 1: Write the failing test**

```python
def test_readme_mentions_inbox_tools_and_code_mode_usage():
    text = Path("README.md").read_text()
    assert "list_inbox_items" in text
    assert "Code Mode" in text
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_docs_contract.py -v`

Expected: FAIL

**Step 3: Write minimal implementation**

Documentation must state:
- Code Mode is recommended to batch normalized MCP calls and compute triage decisions.
- Code Mode is not the primary fix for raw Graph payload bloat.
- Preferred flow:
  1. `list_inbox_items`
  2. `search_inbox_items`
  3. `get_inbox_item_detail`
  4. Code Mode batches follow-up actions over only selected items

Example TypeScript should show:
- registering this MCP server
- fetching summary items
- hydrating top 3
- returning a compact triage report

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_docs_contract.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_docs_contract.py README.md IMPLEMENTATION.md docs/code-mode-inbox-orchestration.md examples/code-mode/inbox_triage.ts
git commit -m "docs: add code mode orchestration guidance"
```

---

## Phase 6: Rollout And Compatibility

### Task 13: Add opt-in rollout flag, flip defaults, and validate token budgets

**Files:**
- Modify: `src/microsoft_mcp/tools.py`
- Modify: `README.md`
- Test: `tests/test_rollout_flags.py`
- Test: `tests/test_token_budgets.py`

**Step 1: Write the failing tests**

```python
def test_assistant_profile_can_be_enabled_with_env_flag(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_RESPONSE_PROFILE", "assistant")
    result = list_emails(limit=5)
    assert "body" not in result[0]


def test_list_emails_summary_stays_under_budget():
    result = list_emails(limit=10)
    serialized = json.dumps(result)
    assert len(serialized) < 12000
```

**Step 2: Run test to verify it fails**

Run: `uv run pytest tests/test_rollout_flags.py tests/test_token_budgets.py -v`

Expected: FAIL

**Step 3: Write minimal implementation**

Rollout rules:
- Add `MICROSOFT_MCP_RESPONSE_PROFILE=legacy|assistant`
- First release:
  - default env = `legacy`
  - new inbox tools always assistant-shaped
  - existing tools accept `response_profile`
- Second release:
  - default env = `assistant`
  - `legacy` remains available temporarily

Budget assertions:
- `list_emails(limit=10)` summary target `< 12k chars`
- `list_events(limit=10)` summary target `< 8k chars`
- `list_contacts(limit=20)` summary target `< 10k chars`
- `list_chat_messages(limit=10)` summary target `< 12k chars`

**Step 4: Run test to verify it passes**

Run: `uv run pytest tests/test_rollout_flags.py tests/test_token_budgets.py -v`

Expected: PASS

**Step 5: Commit**

```bash
git add tests/test_rollout_flags.py tests/test_token_budgets.py src/microsoft_mcp/tools.py README.md
git commit -m "feat: add assistant response rollout and budget gates"
```

---

## Final Verification Checklist

Run:

```bash
uv run pytest tests/ -v
uv run pyright
uvx ruff format .
uvx ruff check --fix --unsafe-fixes .
```

Expected:
- All tests PASS
- Pyright reports no type errors
- Ruff formatting and lint pass cleanly

Additional manual checks:
- Verify [mcp-tool-responses/v1](../../mcp-tool-responses/v1) fixtures can be re-generated and still satisfy compact contracts.
- Verify `list_inbox_items` returns mixed email/event/Teams items with stable `kind`, `score`, and `action_hints`.
- Verify `search_events` returns actionable IDs and `get_event` extracts Teams meeting metadata.
- Verify `list_emails` summary/detail output no longer includes common external-mail warning banners, unreadable wrapped URLs, or full quoted reply chains unless raw/full-thread mode is requested.
- Verify 403 degraded mode returns guidance and cache search metadata rather than raw HTTP noise.

## Notes For The Implementer

- Do not put assistant-shaping logic directly in `graph.py`. Keep `graph.py` transport-oriented and put all assistant contracts in `response_shaping.py`, `inbox_models.py`, and `inbox_ranking.py`.
- Do not use Code Mode to paper over raw server outputs. First make the server produce compact normalized responses; then use Code Mode as a batching and orchestration layer on top.
- Prefer adding `response_profile` parameters and new inbox tools before breaking old tool contracts.
- Do not add database persistence unless the in-memory TTL cache proves insufficient. YAGNI applies.

Plan complete and saved to `docs/plans/2026-03-23-mcp-response-shaping-and-inbox-tools.md`. Two execution options:

1. Subagent-Driven (this session) - I dispatch fresh subagent per task, review between tasks, fast iteration
2. Parallel Session (separate) - Open new session with executing-plans, batch execution with checkpoints

Which approach?
