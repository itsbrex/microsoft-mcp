---
description: Identify and propose cleanup of inbox clutter (newsletters, cold outreach, automated notifications, stale read mail) via microsoft-mcp — confirm-first, never auto-deletes
allowed-tools: ["mcp__microsoft-mcp__call_tool_chain", "mcp__microsoft-mcp__list_tools", "mcp__microsoft-mcp__tools_info", "Write"]
---

Find clutter in the active account's inbox and propose a cleanup plan. Uses the same
in-sandbox classification approach as `/triage-inbox` (read those Principles + sandbox
gotchas first — they apply identically here).

## What counts as clutter

Classify inbox messages into actionable cleanup buckets (compute in-sandbox, print compact):

1. **NEWSLETTERS / marketing** — AUTO/NEWS senders & subjects (CoStar, Cresa Communications, Ren Systems/Alerts, CompStak, Zoom marketing, "digest", "webinar", "unsubscribe"). Candidate: **archive** (or create an inbox rule — see `/rules`).
2. **COLD outreach** — external first-touch sender NOT in the Sent known-contacts allowlist, non-RE:/FW:. Candidate: **archive / delete**.
3. **AUTOMATED notifications** — Salesforce notifications, postmaster/quarantine digests, calendar Canceled/Accepted, OOO "Automatic reply:". Candidate: **archive**.
4. **STALE read** — `is_read=True`, older than N days (default 14), not flagged, not in an active thread you replied to recently. Candidate: **archive**.
5. **HANDLED** — already replied (latest Sent ≥ inbox msg) or has a draft. Candidate: **archive** (thread resolved on your side).

Always **exclude** anything the `/triage-inbox` REPLY/MAYBE buckets would surface, anything
flagged/important, and anything from a known human in an open thread.

## Output → confirm → act

1. Print counts per bucket + a sample (sender|subject) per bucket. Identify **repeat senders** (3+ messages) as inbox-rule candidates.
2. Save a plan to `inbox-declutter-<YYYY-MM-DD>.md`: per bucket, the proposed action and the message IDs.
3. **Stop and present the plan. Do nothing destructive yet.**
4. Only after the user explicitly approves a bucket, act on it using the relevant tools
   (`archive_email`, `bulk_manage_emails`, `move_email`, `delete_email`, or propose an
   `/rules create` for recurring senders). Prefer **archive over delete**. Confirm scope
   ("archive 23 newsletters from the last 14 days?") before bulk operations.

## Guardrails

- **Destructive-by-nature command → confirm-first, always.** Never delete/archive without explicit per-bucket approval. Default to archive, not delete.
- Show message counts and a representative sample before any bulk action; never silently act on a whole bucket.
- For recurring clutter, recommend a durable **inbox rule** (`/rules`) over repeated manual cleanup.
- Read-only classification phase is safe; the act phase is gated on user approval.
