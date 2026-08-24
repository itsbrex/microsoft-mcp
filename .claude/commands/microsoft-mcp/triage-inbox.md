---
description: Precise inbox triage via microsoft-mcp — buckets unread threads into REPLY / MAYBE / NO / GRP / COLD without re-suggesting already-handled mail
allowed-tools: ["mcp__microsoft-mcp__call_tool_chain", "mcp__microsoft-mcp__list_tools", "mcp__microsoft-mcp__search_tools", "mcp__microsoft-mcp__tools_info", "Write"]
---

Triage the active account's inbox with precision. The goal: surface only threads that
genuinely need **the user's** reply — never re-suggest mail already replied to, already
drafted, addressed to someone else, or handled by another participant.

## Principles (learned the hard way — see CLAUDE.md "Code-mode sandbox gotchas")

- `call_tool_chain` results truncate at ~1800 chars. **Compute everything in-sandbox and `print` compact output** — never expect full email dumps back.
- `list_emails` returns a **list directly** (not `{"result": [...]}`). Guard: `x if isinstance(x,list) else x.get("result",[])`.
- Dunder attribute access (`x.__name__`) is banned in the sandbox.
- Email `body` is a dict `{contentType, content}` — unwrap `.get("content")` before regex.
- No raw `conversationId` is exposed. Thread token = `conversation_url.split("readconv/")[1].split("?")[0]`.
- Emails render `[REDACTED_EMAIL]` in tool output, but matching runs server-side on real values, so accuracy is unaffected.

## Algorithm

1. Pull **inbox (75)**, **sent (100)**, **drafts (50)** with `include_body=False`.
2. **Drop handled threads**: my latest Sent in a conversation ≥ the inbox message → replied; or a draft exists in that conversation → in progress.
3. **Drop noise**: `is_read`; AUTO senders (no-reply, CoStar, Ren Systems, CompStak, Zoom, Salesforce notifications, postmaster/quarantine); NEWS subjects (newsletter, digest, webinar, "Automatic reply:", rent index); CAL subjects (Canceled/Accepted/Invitation:).
4. Build a **known-contacts allowlist** from Sent recipients. External sender NOT in it AND not a `RE:/FW:` continuation → **COLD** (first-touch outreach).
5. Threads with **8+ recipients** → **GRP** (watch, others likely handle).
6. For each remaining thread's **newest** message, fetch the body and judge:
   - Strip quoted history + header lines (`From:/To:/Cc:/Subject:/Sent:/_____`) BEFORE analysis so quoted recipient lists & signatures don't leak the user's name.
   - **Greeting target**: first name after Hi/Hey/Hello/Dear. If it's someone else ("Hey Rylan!") → **NO**.
   - **Sender-defer** ("I'll follow up", "we'll handle it", "Not yet, I will…") → **NO**.
   - **Acknowledgment-only** ("great news", "thanks", "keep us posted after…") with no question → **NO** (maybe a future task, not a reply now).
   - Addressed/named the user (distinctive surname, not bare first name) + question → **REPLY**.
   - Sole/near-sole recipient + question → **REPLY**.
   - CC-only → **NO**. Group To + open question → **MAYBE**.
7. **Dedup by conversation** (keep newest). Sort REPLY → MAYBE → NO.

## Output

Print a compact verdict line per thread (`VERDICT | from | subject | date | why`) plus a
short evidence snippet for every non-NO item. Then save a dated report to
`inbox-triage-<YYYY-MM-DD>.md` in the repo root with the full bucket breakdown,
the methodology, and caveats (Sent/draft lookback depth).

## Guardrails

- **Read-only.** This command never sends, archives, deletes, or drafts. To act, the user must explicitly ask (and replies are draft-first — only `send_email_draft` sends).
- State the lookback caveat: a reply older than the 100 most recent Sent items could be missed (low risk for unread inbox items).
- "asks"/greeting detection is a heuristic hint, not certainty — show the evidence snippet so the user can confirm borderline (MAYBE) calls.
