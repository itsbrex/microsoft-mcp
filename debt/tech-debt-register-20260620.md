# Tech Debt Register -- 2026-06-20

**Repo:** `microsoft-mcp` | **Commit date:** 2026-06-20 | **Method:** read-only static analysis (ruff, radon cc/mi, AST function-length scan, grep)

## Severity scale

| Score | Meaning |
|-------|---------|
| 5 | Blocks releases |
| 4 | Recurring incidents / major maintenance drag |
| 3 | Ongoing friction for contributors |
| 2 | Minor annoyance, low risk |
| 1 | Cosmetic |

## Effort scale

| Label | Meaning |
|-------|---------|
| S | < 1 day |
| M | 1--3 days |
| L | 3--5 days |
| XL | > 1 week |

---

## Register

| id | category | severity | effort | evidence | remediation |
|----|----------|----------|--------|----------|-------------|
| TD-01 | size | 4 | XL | `tools.py` is 7 586 lines (41% of all source), maintainability index 0.00 (grade C), 16 functions >80 lines, 109 `@mcp.tool` decorators. Largest file by 6.4x. | Split into domain submodules (`tools_email.py`, `tools_calendar.py`, `tools_search.py`, `tools_teams.py`, `tools_rules.py`, `tools_todo.py`, etc.) behind a barrel re-export. |
| TD-02 | complexity | 4 | L | `unified_search` (tools.py:4844) -- 310 lines, cyclomatic complexity 44 (grade F). Highest CC in the codebase. | Extract into a dedicated `search.py` module; break into `_build_search_request`, `_process_search_hit`, `_format_results` helpers. |
| TD-03 | dependencies | 4 | S | `azure-identity` and `msgraph-sdk` are completely unpinned (no version constraint). A breaking major bump silently breaks the project. | Add upper-bound pins: `azure-identity>=1.17,<2`, `msgraph-sdk>=1.0,<2` (adjust to current major). |
| TD-04 | duplication | 3 | M | 58 of 91 `except Exception` blocks in `tools.py` are identical `logger.error(f"Failed to ...", exc_info=True); raise` wrappers around `graph.request` calls. 104 call sites, 102 except blocks, 101 logger.error calls. | Introduce a `@graph_tool(action="...")` decorator or `async def graph_call(action, coro)` helper that handles logging + re-raise uniformly. |
| TD-05 | coverage | 3 | S | No test coverage measurement configured: no `pytest-cov` dependency, no `.coveragerc`, no `[tool.coverage]` in `pyproject.toml`, no `pragma: no cover` markers. Coverage is unknown. | Add `pytest-cov` to dev deps, add `[tool.coverage.run]` section, run `pytest --cov=microsoft_mcp` to establish a baseline. |
| TD-06 | complexity | 3 | M | `graph.search_query` (graph.py:383) -- 125 lines, CC=30 (grade D). `graph.request` (graph.py:90) -- 81 lines, CC=24 (grade D). | Extract retry logic into a `_retry_with_backoff` helper; split `search_query` hit-processing into a separate function. |
| TD-07 | size | 3 | L | 31 functions exceed 80 lines across the codebase. Top offenders beyond TD-01/TD-02: `list_channel_messages` (175), `list_chat_messages` (154), `create_email_draft` (152), `import_inbox_rules` (138), `verify_account_tokens` (131). | Apply extract-method refactoring to each; target <60 lines per function. Prioritize the 10 functions >100 lines first. |
| TD-08 | dependencies | 3 | S | `black>=23.3.0` is listed in runtime `[project.dependencies]` but is a dev-only formatting tool. End users installing `microsoft-mcp` pull in black unnecessarily. | Move `black` to `[dependency-groups] dev`. |
| TD-09 | complexity | 3 | M | `signature_parser.parse_email_body` CC=27 (grade D, 87 lines), `rules.rule_to_template` CC=24, `rules.validate_template` CC=24. | Break conditional chains into helper functions; consider table-driven dispatch for rule template conversion. |
| TD-10 | duplication | 2 | S | Environment variable names (`MICROSOFT_MCP_TOKENS_DIR`, `_CLIENT_ID`, `_TENANT_ID`, `_ACCOUNT_ID`) appear as raw string literals 5--8 times each across modules. | Define as module-level constants in a `constants.py` or at the top of each consuming module. |
| TD-11 | duplication | 2 | S | 4 pagination sites in `tools.py` manually reimplement `while next_link` loop despite `graph.request_paginated` existing. | Migrate manual pagination loops to use `graph.request_paginated`. |
| TD-12 | dependencies | 2 | S | `markitdown>=0.1.3` and `RestrictedPython>=6.0` have floor-only pins. `markitdown` is pre-1.0 (semver allows breaking changes). | Add upper-bound: `markitdown>=0.1.3,<1`, `RestrictedPython>=6.0,<8`. |
| TD-13 | dependencies | 2 | S | `pytest-asyncio>=1.0.0` in dev deps -- latest published release is 0.x. Floor version matches no existing release (works only if uv resolves differently or a fork exists). | Verify intended version; likely should be `>=0.21.0,<1`. |
| TD-14 | complexity | 2 | M | `code_mode.py` -- 878 lines, maintainability index 9.32 (grade B, borderline). `call_tool_chain` is 97 lines. | Extract sandboxing and tool-dispatch logic into helper functions. |
| TD-15 | size | 2 | M | `auth_msal.py` is 1 179 lines with 4 functions >80 lines (`verify_account_tokens` 131, `authenticate` 124, `refresh_all_accounts` 84, `_refresh_access_token` 83). | Extract token file I/O into a `_token_store.py` helper; simplify multi-step verify logic. |
| TD-16 | smells | 2 | S | 3 `# noqa` suppressions in src/: `F401` on deferred import in `tools.py`, `F401` on type-annotation import in `bounces_cli.py`, `ARG001` on unused parameter in `signature_parser.py`. | Review each: the `F401` in tools.py is documented (ruff hook race); the ARG001 suggests `default_region` may need a real implementation. |
| TD-17 | smells | 1 | S | 2 `# type: ignore[attr-defined]` in `graph.py` for `force_refresh()` -- duck-typing the auth provider. 5 additional suppressions in test files (standard test-monkey-patching). | Add `force_refresh()` to the `AuthProvider` protocol as an optional method, or use `hasattr` guard, to eliminate src suppressions. |
| TD-18 | smells | 1 | S | 14 ruff violations in test files: 12 `E402` (late imports) + 1 `F811` (re-import) + 1 duplicate import in `test_auth.py`. | Move imports to top of file or add `# noqa: E402` with justification. Quick cleanup. |
| TD-19 | size | 1 | S | `tools.py` comment density is 3% (195 comments / 7 586 lines). Low inline documentation for the largest module. | Add section-header docstrings and inline comments for non-obvious Graph API patterns as part of TD-01 split. |

---

## Metrics snapshot

| Metric | Value |
|--------|-------|
| Source files (src/) | 35 |
| Source lines (src/) | ~19 300 |
| Test files | 61 |
| Test lines | ~16 800 |
| Test-to-source ratio | 0.87 |
| Registered MCP tools | 109 |
| Functions >80 lines | 31 |
| Cyclomatic complexity F-grade | 1 (`unified_search`) |
| Cyclomatic complexity D-grade | 8 |
| Maintainability index C-grade | 1 (`tools.py`) |
| Ruff violations (src/) | 0 |
| Ruff violations (tests/) | 14 |
| Smell markers (TODO/FIXME/HACK/XXX) | 0 |
| Type-ignore / noqa suppressions (src/) | 5 |
| Unpinned dependencies | 2 |
| Test coverage config | None |
| Global mutable singletons | 2 (`_global_auth`, `_global_cache`) |

---

## Summary

The codebase is **clean on the surface** -- zero ruff violations in source, zero TODO/FIXME markers, zero circular imports, and a well-designed dependency-injection architecture. The debt is **structural**: `tools.py` at 7 586 lines with a maintainability index of 0.00 is the dominant burden (TD-01), compounded by a 310-line grade-F function (TD-02) and pervasive try/except boilerplate (TD-04). Two fully unpinned dependencies (TD-03) pose the highest-likelihood incident risk for the lowest fix cost. Standing up test coverage measurement (TD-05) is a one-hour task that would establish a baseline for all future refactoring. Addressing TD-01 through TD-05 would resolve the top five items and meaningfully reduce maintenance friction.
