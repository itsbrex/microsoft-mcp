# Add Code Mode Plan

**Goal:** Rework `microsoft-mcp` so the existing Microsoft Graph tool surface remains intact while the server also exposes a real, integrated code-mode orchestration layer with discovery, interface introspection, prompt guidance, and sandboxed multi-tool execution.

**Architecture:** Keep the current Graph/auth/tool business logic in Python. Add a Python-native code-mode runtime that reflects the live FastMCP registry, exposes bridge-style orchestration tools, and executes Python code against the active tool set inside a restricted sandbox. Avoid turning the repo into a generic UTCP bridge or introducing a mandatory Node runtime.

**Key constraints:**
- Preserve current Microsoft Graph behavior and auth semantics.
- Use the active FastMCP tool registry so auth-mode-specific tool visibility remains accurate.
- Mirror the useful bridge semantics from `universal-tool-calling-protocol/code-mode`:
  `search_tools`, `list_tools`, `tools_info`, `get_required_keys_for_tool`, `call_tool_chain`, and usage prompt guidance.
- Document the sandbox honestly as cooperative, Python-native, and weaker than hardened process isolation.

---

## Phase 1: Foundation And Runtime Shape

### Objectives
- Define the integrated code-mode model for this server.
- Separate orchestration/runtime logic from Graph tool implementations.
- Establish the live tool-registry adapter that powers discovery and execution.

### Tasks
1. Add a new runtime module under `src/microsoft_mcp/` for code-mode behavior.
2. Build adapters that read the current FastMCP registry and expose:
   - active tool name
   - sanitized namespace access name
   - description
   - input schema
   - output schema
   - required environment keys
3. Generate Python interface text from the live JSON schemas.
4. Implement runtime search ranking over tool names, descriptions, and schema fields.
5. Define a fixed built-in namespace for the server tools, e.g. `microsoft.<tool>()`.

### Deliverables
- Runtime module with tool registry discovery and interface generation.
- Stable internal contract for the orchestration-facing tool metadata.

---

## Phase 2: Sandboxed Execution Engine

### Objectives
- Provide a Python-native `call_tool_chain` experience aligned with code-mode behavior.
- Support multi-step orchestration against the existing tool set in one execution.

### Tasks
1. Add `RestrictedPython` dependency to the project.
2. Implement a cooperative sandbox with:
   - restricted imports
   - limited safe builtins
   - print/log capture
   - configurable timeout
   - runtime context variables mirroring code-mode concepts:
     - `__interfaces`
     - `__get_tool_interface(...)`
3. Add namespaced tool wrappers that invoke the underlying FastMCP `FunctionTool.fn` callables directly.
4. Ensure result + logs are returned in a compact, structured payload.
5. Handle execution failures with explicit runtime errors and captured logs.

### Deliverables
- Sandboxed executor with timeout and observability.
- Tool wrappers bound to the active Microsoft MCP tool set.

---

## Phase 3: MCP Surface Integration

### Objectives
- Expose the new code-mode layer as first-class MCP tools and prompts inside `microsoft-mcp`.

### Tasks
1. Register a code-mode usage prompt analogous to `utcp_codemode_usage`.
2. Add MCP tools:
   - `search_tools`
   - `list_tools`
   - `tools_info`
   - `get_required_keys_for_tool`
   - `call_tool_chain`
3. Keep the existing business tools unchanged.
4. Make the new tool set reflect the active auth mode so hidden Teams tools stay hidden under MSAL.
5. Keep tool outputs compact and machine-friendly.

### Deliverables
- Integrated code-mode MCP surface in the same server process.
- Prompt guidance that teaches `search_tools -> tools_info/__interfaces -> call_tool_chain`.

---

## Phase 4: Packaging, Docs, And Examples

### Objectives
- Make the integrated code-mode surface installable, discoverable, and usable.

### Tasks
1. Update `pyproject.toml` with the runtime dependency.
2. Update `README.md` to describe:
   - the integrated code-mode capability
   - when to use direct tool calls vs `call_tool_chain`
   - auth behavior with the orchestration layer
3. Update `IMPLEMENTATION.md` to reflect the new architecture.
4. Replace the current documentation-only framing in `docs/code-mode-inbox-orchestration.md` with real integrated usage.
5. Replace the simulated `examples/code-mode/inbox_triage.ts` story with executable integrated examples, likely Python-first.
6. Add at least one end-to-end orchestration example that exercises selective hydration over existing inbox/search tools.

### Deliverables
- Updated package metadata.
- Updated architecture docs.
- Real code-mode examples instead of simulated stubs.

---

## Phase 5: Validation And Regression Coverage

### Objectives
- Prove the new layer works without regressing the current server behavior.

### Tasks
1. Add focused tests for runtime metadata:
   - live tool enumeration
   - auth-aware tool visibility
   - interface generation
   - required env key reporting
2. Add sandbox tests for:
   - successful multi-tool orchestration
   - timeout behavior
   - error capture
   - restricted imports / blocked unsafe operations
3. Add doc/example contract tests:
   - README mentions the integrated code-mode surface
   - referenced examples/docs exist
   - example tool names match the actual registry
4. Retain current inbox/search shaping tests to ensure orchestration stays token-efficient.
5. Run targeted and full test suites and fix regressions.

### Deliverables
- New runtime/tool tests.
- Updated doc contract tests.
- Passing verification suite.

---

## Parallel Execution Waves

### Wave 1: Analysis And Planning
- Main agent: architecture decision, plan authoring, runtime contract.
- Explorer agents: current server integration points, code-mode capability extraction, existing doc/test gap analysis.

### Wave 2: Independent Implementation Slices
- Worker A: runtime metadata + schema/interface generation.
- Worker B: sandbox executor + timeout/log handling.
- Worker C: docs/examples rewrite once the public surface is stable.

### Wave 3: Verification And Stabilization
- Worker D: tests for runtime and execution engine.
- Worker E: doc/example contract tests and README assertions.
- Main agent: integrate, resolve conflicts, run final verification, update architecture notes.

---

## Acceptance Criteria

- Existing Microsoft Graph tools still work with current auth flows.
- The server exposes integrated code-mode tools and a usage prompt.
- `search_tools` returns relevant tool matches from the active registry.
- `tools_info` and `__interfaces` expose generated Python interfaces.
- `call_tool_chain` can orchestrate multiple existing tools in one execution and returns both `result` and `logs`.
- Teams tool visibility remains auth-aware.
- Docs and examples describe the actual integrated flow, not an external simulated one.
- Tests cover runtime discovery, execution, failure handling, and documentation contracts.
