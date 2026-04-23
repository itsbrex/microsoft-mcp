---
description: Hunt and fix technical debt in recently-changed code
argument-hint: "[path or glob, defaults to src/]"
---

Scan the target (default `src/`) for tech debt and either fix it or list the top 5 items with the ROI of each. Focus on:

1. **Duplication** — identical blocks that should be extracted. Use `rg` to find clones.
2. **Dead code** — unused imports/functions/flags (pyflakes is your friend).
3. **Silent failures** — bare `except`, `except Exception: pass`, swallowed Graph errors.
4. **Graph-API anti-patterns** — hand-rolled pagination that ignores `@odata.nextLink`, HTTP calls that bypass `graph.request`, missing retry on 429/5xx.
5. **Auth duplication** — bypassing `get_auth_instance()` and re-instantiating credentials.
6. **Response shaping drift** — tools that ignore `response_profile` on list/search endpoints.
7. **Test gaps** — tool functions with no corresponding test in `tests/`.

Rules:
- No rewrites for its own sake. Only touch something if the fix is <20 lines and has a clear win.
- If the finding is large (>1 file or >50 lines), open a TODO in `TODO.md` with a one-paragraph writeup instead of fixing it inline.
- After edits, run `/test` and `/lint` before handing back.

Target: $ARGUMENTS
