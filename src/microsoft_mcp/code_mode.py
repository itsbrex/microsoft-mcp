"""Python-native code-mode runtime for Microsoft MCP.

This module provides a cooperative, Python-only code-execution layer over the
live FastMCP tool registry. It is intentionally scoped to the current server's
registered tools and does not try to become a generic UTCP implementation.

The runtime mirrors the useful parts of the `code-mode` pattern:

- discover tools dynamically from the active MCP registry
- search tools by name, description, tags, parameters, and schema text
- generate interface text from live JSON schemas
- expose namespaced access as ``microsoft.<tool>()``
- execute multi-step workflows in one sandboxed pass
- return structured execution results and logs

Sandbox model
-------------
The execution sandbox is cooperative, not hardened multi-tenant isolation.
It uses RestrictedPython when available, restricted imports, limited builtins,
and an execution timeout. It should be treated as a guardrail for trusted
agent-generated code, not as a security boundary for hostile code.
"""

from __future__ import annotations

import asyncio
import dataclasses
import importlib
import inspect
import json
import logging
import os
import re
import textwrap
import warnings
from collections.abc import Awaitable, Callable, Mapping, Sequence
from dataclasses import dataclass
from types import SimpleNamespace
from typing import Any

logger = logging.getLogger(__name__)


@dataclass(frozen=True, slots=True)
class CodeModeToolSummary:
    """Compact metadata for a live MCP tool."""

    name: str
    access_pattern: str
    description: str
    tags: tuple[str, ...]
    score: float = 0.0


@dataclass(frozen=True, slots=True)
class CodeModeToolDetails(CodeModeToolSummary):
    """Detailed tool metadata including schemas and required keys."""

    input_schema: dict[str, Any] | None = None
    output_schema: dict[str, Any] | None = None
    required_keys: tuple[str, ...] = ()
    python_interface: str = ""


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
        return x**y
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


def _run_coroutine_sync(coro: Awaitable[Any]) -> Any:
    """Run a coroutine to completion whether or not a loop is already running.

    Mirrors tools._run_async so sandboxed user code can invoke async
    tools regardless of how call_tool_chain itself was driven.
    """
    try:
        asyncio.get_running_loop()
    except RuntimeError:
        return asyncio.run(coro)
    with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
        return executor.submit(lambda: asyncio.run(coro)).result()


class CodeModeRuntime:
    """Runtime adapter for the active Microsoft MCP FastMCP registry."""

    AGENT_PROMPT_TEMPLATE = """
## Microsoft MCP Code Mode Usage Guide

You have access to a Microsoft MCP code-mode runtime that lets you execute
Python code with direct access to the live tool registry.

### 1. Discover tools first
- Use `search_tools()` to find relevant tools.
- Use `list_tools()` to inspect the active registry.
- Use `tools_info()` or `interfaces` to understand exact contracts.

### 2. Execute code in one pass
- Use `microsoft.<tool>(...)` for tool calls.
- Chain operations locally instead of asking for repeated tool discovery.
- Execute the workflow with `call_tool_chain()`.
- Return the final value from your code block.

### 3. Sandbox expectations
- The sandbox is cooperative, not hardened against hostile code.
- File system and network access are not provided by this runtime.
- Only a small set of safe imports and builtins are exposed.
- Timeouts are enforced for runaway code.

### 4. Output discipline
- Keep intermediate logging concise.
- Prefer compact summaries over raw Graph payloads.
- Return only the data needed for the next step.
""".strip()

    def __init__(
        self,
        mcp: Any,
        namespace: str = "microsoft",
        default_timeout: float = 30.0,
        memory_limit: int | None = 128,
        excluded_tools: Sequence[str] = (),
        tool_provider: Callable[[], Sequence[Any] | Awaitable[Sequence[Any]]]
        | None = None,
    ) -> None:
        self._mcp = mcp
        self._namespace = namespace
        self._default_timeout = default_timeout
        self._memory_limit = memory_limit
        self._excluded_tools = frozenset(excluded_tools)
        self._tool_provider = tool_provider
        self._tool_cache: dict[str, Any] = {}
        self._current_registry: list[Any] = []
        self._tool_summaries: list[CodeModeToolSummary] = []
        self._tool_details: dict[str, CodeModeToolDetails] = {}
        self._tool_namespace = SimpleNamespace()
        self._trace_sink: list[dict[str, Any]] | None = None

    @classmethod
    async def create(
        cls,
        mcp: Any,
        namespace: str = "microsoft",
        default_timeout: float = 30.0,
        memory_limit: int | None = 128,
    ) -> "CodeModeRuntime":
        runtime = cls(
            mcp=mcp,
            namespace=namespace,
            default_timeout=default_timeout,
            memory_limit=memory_limit,
        )
        await runtime.refresh()
        return runtime

    @property
    def namespace(self) -> str:
        return self._namespace

    @property
    def tool_namespace(self) -> SimpleNamespace:
        return self._tool_namespace

    async def refresh(self) -> None:
        """Refresh the live tool registry from the FastMCP instance."""

        registry = await self._list_registered_tools()
        self._current_registry = registry
        self._tool_cache.clear()
        summaries: list[CodeModeToolSummary] = []
        details: dict[str, CodeModeToolDetails] = {}
        namespace = SimpleNamespace()

        for tool in registry:
            summary, detail = self._build_tool_metadata(tool)
            summaries.append(summary)
            details[summary.name] = detail
            setattr(
                namespace,
                self._sanitize_identifier(summary.name),
                self._make_tool_wrapper(summary.name),
            )

        self._tool_summaries = summaries
        self._tool_details = details
        self._tool_namespace = namespace
        logger.debug("Code mode registry refreshed with %d tools", len(summaries))

    async def list_tools(self) -> list[dict[str, Any]]:
        """Return compact metadata for the active registry."""

        if not self._tool_summaries:
            await self.refresh()
        return [dataclasses.asdict(tool) for tool in self._tool_summaries]

    async def search_tools(self, query: str, limit: int = 10) -> list[dict[str, Any]]:
        """Search the live registry using a simple lexical ranking model."""

        if not self._tool_summaries:
            await self.refresh()

        tokens = [token for token in self._tokenize(query) if token]
        scored: list[CodeModeToolSummary] = []
        for tool in self._tool_summaries:
            score = self._score_tool(tool, tokens)
            if score > 0:
                scored.append(dataclasses.replace(tool, score=score))

        scored.sort(key=lambda item: (-item.score, item.name))
        results: list[dict[str, Any]] = []
        for tool in scored[:limit]:
            detail = self._tool_details.get(tool.name)
            item = dataclasses.asdict(tool)
            if detail is not None:
                item["python_interface"] = detail.python_interface
                item["required_keys"] = list(detail.required_keys)
            results.append(item)
        return results

    async def tools_info(self, tool_names: Sequence[str]) -> list[dict[str, Any]]:
        """Return detailed metadata and generated interfaces for named tools."""

        if not self._tool_details:
            await self.refresh()

        result: list[dict[str, Any]] = []
        for name in tool_names:
            detail = self._resolve_detail(name)
            if detail is None:
                result.append(
                    {
                        "name": name,
                        "found": False,
                        "error": f"Tool '{name}' is not registered in the active MCP registry.",
                    }
                )
                continue

            result.append(
                {
                    "name": detail.name,
                    "access_pattern": detail.access_pattern,
                    "description": detail.description,
                    "tags": list(detail.tags),
                    "found": True,
                    "required_keys": list(detail.required_keys),
                    "input_schema": detail.input_schema,
                    "output_schema": detail.output_schema,
                    "python_interface": detail.python_interface,
                }
            )
        return result

    async def get_required_keys_for_tool(self, tool_name: str) -> dict[str, Any]:
        """Report required env keys or other declared requirements for a tool."""

        if not self._tool_details:
            await self.refresh()

        detail = self._resolve_detail(tool_name)
        if detail is None:
            return {
                "tool": tool_name,
                "found": False,
                "required_keys": [],
                "required_env_keys": [],
                "notes": f"Tool '{tool_name}' is not registered in the active MCP registry.",
            }

        declared_keys = list(detail.required_keys)
        if not declared_keys:
            declared_keys = ["MICROSOFT_MCP_CLIENT_ID"]
            if os.getenv("MICROSOFT_MCP_AUTH_METHOD", "azure").lower() == "msal":
                declared_keys.extend(
                    [
                        "MICROSOFT_MCP_AUTH_METHOD",
                        "MICROSOFT_MCP_ACCOUNT_ID",
                    ]
                )

        return {
            "tool": detail.name,
            "access_pattern": detail.access_pattern,
            "found": True,
            "required_keys": declared_keys,
            "required_env_keys": declared_keys,
            "notes": (
                "Required keys are derived from tool metadata when available, "
                "otherwise they fall back to the current authentication mode defaults."
            ),
        }

    async def get_interfaces(self) -> str:
        """Generate Python type hints for all active tools."""

        if not self._tool_details:
            await self.refresh()

        lines = [
            "# Auto-generated Python interfaces for Microsoft MCP code mode",
            "from __future__ import annotations",
            "",
            "from typing import Any, Literal, Optional, TypedDict",
            "",
        ]
        for detail in self._tool_details.values():
            lines.append(detail.python_interface.rstrip())
            lines.append("")
        return "\n".join(lines).rstrip()

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

        if not self._tool_details:
            await self.refresh()

        timeout = self._default_timeout if timeout is None else timeout
        sandbox = self._build_sandbox()
        interfaces = await self.get_interfaces()
        available_tools = [tool.name for tool in self._tool_summaries]
        available_access_patterns = [
            tool.access_pattern for tool in self._tool_summaries
        ]
        interface_map = self._build_interface_lookup_map()
        interface_map_json = json.dumps(interface_map, sort_keys=True)

        def get_tool_interface(name: str) -> str | None:
            detail = self._resolve_detail(name)
            return detail.python_interface if detail else None

        # Preferred names for RestrictedPython user code.
        sandbox["interfaces"] = interfaces
        sandbox["available_tools"] = available_tools
        sandbox["availableTools"] = available_access_patterns
        sandbox["get_tool_interface"] = get_tool_interface
        sandbox["getToolInterface"] = get_tool_interface
        sandbox["interface_map"] = interface_map
        sandbox["interface_map_json"] = interface_map_json
        sandbox["interfaceMapJson"] = interface_map_json

        # Legacy aliases (kept for parity with code-mode conventions in non-Restricted environments).
        sandbox["__interfaces"] = interfaces
        sandbox["__available_tools"] = available_tools
        sandbox["__availableTools"] = available_access_patterns
        sandbox["__get_tool_interface"] = get_tool_interface
        sandbox["__getToolInterface"] = get_tool_interface
        sandbox["__interface_map_json"] = interface_map_json
        sandbox["__interfaceMapJson"] = interface_map_json
        sandbox[self._namespace] = self._tool_namespace

        wrapped = self._wrap_user_code(code)
        compiled = self._compile_code(wrapped)

        trace: list[dict[str, Any]] = []
        logs: list[str] = []
        sandbox["__trace__"] = trace
        sandbox["__logs__"] = logs
        self._trace_sink = trace
        sandbox["print"] = self._make_print(logs)

        def execute() -> Any:
            exec(compiled, sandbox, sandbox)
            fn = sandbox.get("user_code_function")
            if not callable(fn):
                return None
            return fn()

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

    def _build_tool_metadata(
        self, tool: Any
    ) -> tuple[CodeModeToolSummary, CodeModeToolDetails]:
        name = getattr(tool, "name", "")
        description = (getattr(tool, "description", "") or "").strip()
        tags_value = getattr(tool, "tags", ())
        if isinstance(tags_value, str):
            tags = tuple(sorted(tag for tag in re.split(r"[,\s]+", tags_value) if tag))
        else:
            tags = tuple(sorted(str(tag) for tag in tags_value or ()))

        access_pattern = f"{self._namespace}.{self._sanitize_identifier(name)}"
        input_schema = self._jsonish(getattr(tool, "parameters", None))
        output_schema = self._jsonish(getattr(tool, "output_schema", None))
        required_keys = self._extract_required_keys(tool)
        python_interface = self._build_python_interface(
            name, description, input_schema, output_schema, access_pattern
        )

        summary = CodeModeToolSummary(
            name=name,
            access_pattern=access_pattern,
            description=description,
            tags=tags,
        )
        details = CodeModeToolDetails(
            name=name,
            access_pattern=access_pattern,
            description=description,
            tags=tags,
            input_schema=input_schema,
            output_schema=output_schema,
            required_keys=required_keys,
            python_interface=python_interface,
        )
        return summary, details

    def _build_python_interface(
        self,
        tool_name: str,
        description: str,
        input_schema: dict[str, Any] | None,
        output_schema: dict[str, Any] | None,
        access_pattern: str,
    ) -> str:
        input_type_name = f"{self._sanitize_identifier(tool_name)}Input"
        output_type_name = f"{self._sanitize_identifier(tool_name)}Output"
        input_body = self._schema_to_typed_dict_body(input_schema)
        output_body = self._schema_to_typed_dict_body(output_schema)
        description_line = description.splitlines()[0] if description else ""

        lines = [
            f"class {input_type_name}(TypedDict, total=False):",
            *self._indent_interface_body(input_body),
            "",
            f"class {output_type_name}(TypedDict, total=False):",
            *self._indent_interface_body(output_body),
            "",
            f"# {description_line}" if description_line else "#",
            f"# Access as: {access_pattern}(...)",
        ]
        return "\n".join(lines).strip()

    def _schema_to_typed_dict_body(self, schema: dict[str, Any] | None) -> str:
        if not schema:
            return "pass"

        schema_type = schema.get("type")
        if schema_type != "object":
            return "pass"

        properties = schema.get("properties") or {}
        required = set(schema.get("required") or [])
        lines: list[str] = []
        for key, value in properties.items():
            py_type = self._schema_to_python_type(value)
            if key not in required:
                py_type = f"Optional[{py_type}]"
            lines.append(f"{self._sanitize_identifier(key)}: {py_type}")

        return "\n".join(lines) if lines else "pass"

    def _indent_interface_body(self, body: str) -> list[str]:
        if not body.strip():
            return ["    pass"]
        return [f"    {line}" for line in body.splitlines()]

    def _schema_to_python_type(self, schema: Any) -> str:
        if not isinstance(schema, dict):
            return "Any"

        if "enum" in schema and isinstance(schema["enum"], list):
            literals = ", ".join(repr(value) for value in schema["enum"])
            return f"Literal[{literals}]"

        schema_type = schema.get("type")
        if isinstance(schema_type, list):
            return " | ".join(
                self._map_json_type_to_python(part) for part in schema_type
            )
        if schema_type == "array":
            item_schema = schema.get("items")
            return (
                f"list[{self._schema_to_python_type(item_schema)}]"
                if item_schema
                else "list[Any]"
            )
        if schema_type == "object":
            return "dict[str, Any]"
        return self._map_json_type_to_python(schema_type)

    def _map_json_type_to_python(self, json_type: Any) -> str:
        return {
            "string": "str",
            "integer": "int",
            "number": "float",
            "boolean": "bool",
            "null": "None",
            "array": "list[Any]",
            "object": "dict[str, Any]",
        }.get(json_type, "Any")

    def _extract_required_keys(self, tool: Any) -> tuple[str, ...]:
        meta_candidates: list[Any] = []
        for attr in ("meta", "annotations", "task_config"):
            value = getattr(tool, attr, None)
            if value is not None:
                meta_candidates.append(value)

        discovered: list[str] = []
        for candidate in meta_candidates:
            if isinstance(candidate, Mapping):
                for key in ("required_keys", "required_env_keys", "env_keys"):
                    value = candidate.get(key)
                    if isinstance(value, Sequence) and not isinstance(
                        value, (str, bytes)
                    ):
                        discovered.extend(str(item) for item in value)

        return tuple(dict.fromkeys(discovered))

    def _wrap_user_code(self, code: str) -> str:
        stripped = code.strip()
        indented = textwrap.indent(stripped or "pass", "    ")
        return f"def user_code_function():\n{indented}\n"

    def _compile_code(self, wrapped_code: str) -> Any:
        compiler = self._load_restricted_python()
        if compiler is None:
            raise RuntimeError(
                "RestrictedPython is required for call_tool_chain but is not installed."
            )
        with warnings.catch_warnings():
            warnings.filterwarnings(
                "ignore",
                message=r"Line \d+: Prints, but never reads 'printed' variable\.",
                category=SyntaxWarning,
            )
            return compiler(wrapped_code, "<code-mode>", "exec")

    def _build_sandbox(self) -> dict[str, Any]:
        restricted_globals = self._load_restricted_python_globals()
        sandbox: dict[str, Any] = dict(restricted_globals)
        builtins = sandbox.setdefault("__builtins__", {})
        if isinstance(builtins, dict):
            builtins.update(
                {
                    "abs": abs,
                    "all": all,
                    "any": any,
                    "bool": bool,
                    "dict": dict,
                    "enumerate": enumerate,
                    "filter": filter,
                    "float": float,
                    "int": int,
                    "len": len,
                    "list": list,
                    "max": max,
                    "min": min,
                    "print": lambda *args, **kwargs: None,
                    "range": range,
                    "reversed": reversed,
                    "round": round,
                    "set": set,
                    "sorted": sorted,
                    "str": str,
                    "sum": sum,
                    "tuple": tuple,
                    "zip": zip,
                    "__import__": self._restricted_import,
                }
            )

        sandbox.update(
            {
                "Any": Any,
                "Callable": Callable,
                "Awaitable": Awaitable,
                "json": json,
                "math": importlib.import_module("math"),
                "re": re,
                "SimpleNamespace": SimpleNamespace,
                "datetime": importlib.import_module("datetime"),
                "timedelta": importlib.import_module("datetime").timedelta,
                "time": importlib.import_module("time"),
                "typing": importlib.import_module("typing"),
                "_getattr_": getattr,
                "_getitem_": lambda obj, key: obj[key],
                "_write_": lambda obj: obj,
            }
        )
        return sandbox

    def _load_restricted_python(self) -> Callable[[str, str, str], Any] | None:
        try:
            module = importlib.import_module("RestrictedPython")
        except ModuleNotFoundError:
            return None
        return getattr(module, "compile_restricted", None)

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
        safe_globals["_getiter_"] = eval_mod.default_guarded_getiter
        safe_globals["_getitem_"] = eval_mod.default_guarded_getitem
        safe_globals["_iter_unpack_sequence_"] = getattr(
            guards,
            "guarded_iter_unpack_sequence",
            lambda it, spec, _getiter_: tuple(it),
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

    def _restricted_import(self, name: str, *args: Any, **kwargs: Any) -> Any:
        allowed = {
            "asyncio",
            "collections",
            "datetime",
            "functools",
            "json",
            "math",
            "operator",
            "re",
            "time",
            "typing",
            "uuid",
        }
        if name not in allowed:
            raise ImportError(
                f"Import of '{name}' is not allowed in the code-mode sandbox."
            )
        return importlib.import_module(name)

    async def _list_registered_tools(self) -> list[Any]:
        if self._tool_provider is not None:
            provided_tools = self._tool_provider()
            if inspect.isawaitable(provided_tools):
                provided_tools = await provided_tools
            return [
                tool
                for tool in provided_tools
                if getattr(tool, "name", None) not in self._excluded_tools
            ]

        tools = await self._mcp._list_tools_middleware()
        return [
            tool
            for tool in tools
            if getattr(tool, "enabled", True)
            and getattr(tool, "name", None) not in self._excluded_tools
        ]

    def _make_tool_wrapper(self, tool_name: str) -> Callable[..., Any]:
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

        return call_tool

    def _lookup_tool(self, tool_name: str) -> Any:
        for tool in getattr(self, "_current_registry", []):
            if getattr(tool, "name", None) == tool_name:
                return tool
        # Fallback to a refreshed registry if needed.
        raise KeyError(f"Tool '{tool_name}' is not available in the active registry.")

    def _tokenize(self, query: str) -> list[str]:
        return [token for token in re.split(r"[^a-zA-Z0-9_]+", query.lower()) if token]

    def _score_tool(self, tool: CodeModeToolSummary, tokens: Sequence[str]) -> float:
        if not tokens:
            return 0.0

        haystack_parts = [
            tool.name.lower(),
            tool.access_pattern.lower(),
            tool.description.lower(),
            " ".join(tool.tags).lower(),
        ]
        detail = self._tool_details.get(tool.name)
        if detail and detail.input_schema:
            haystack_parts.append(
                json.dumps(detail.input_schema, sort_keys=True).lower()
            )
        haystack = " \n ".join(haystack_parts)

        score = 0.0
        for token in tokens:
            if token == tool.name.lower():
                score += 20.0
            elif tool.name.lower().startswith(token):
                score += 10.0
            elif token in tool.name.lower():
                score += 7.0
            elif token in tool.access_pattern.lower():
                score += 6.0
            elif token in haystack:
                score += 2.0

        return score

    def _jsonish(self, value: Any) -> dict[str, Any] | None:
        if value is None:
            return None
        if hasattr(value, "model_dump"):
            try:
                dumped = value.model_dump()
                return (
                    dumped
                    if isinstance(dumped, dict)
                    else json.loads(json.dumps(dumped))
                )
            except Exception:
                pass
        if isinstance(value, Mapping):
            return dict(value)
        try:
            return json.loads(json.dumps(value, default=str))
        except Exception:
            return {"value": str(value)}

    def _sanitize_identifier(self, value: str) -> str:
        cleaned = re.sub(r"[^0-9a-zA-Z_]", "_", value)
        if cleaned and cleaned[0].isdigit():
            cleaned = f"_{cleaned}"
        return cleaned or "_tool"

    def _build_interface_lookup_map(self) -> dict[str, str]:
        interface_map: dict[str, str] = {}
        for detail in self._tool_details.values():
            interface_map[detail.name] = detail.python_interface
            interface_map[detail.access_pattern] = detail.python_interface
        return interface_map

    def _resolve_detail(self, name: str) -> CodeModeToolDetails | None:
        raw_name = str(name or "").strip()
        if not raw_name:
            return None

        candidates: list[str] = []
        seen: set[str] = set()

        def add_candidate(value: str) -> None:
            candidate = str(value).strip()
            if candidate and candidate not in seen:
                seen.add(candidate)
                candidates.append(candidate)

        add_candidate(raw_name)
        if ":" in raw_name:
            add_candidate(raw_name.split(":", 1)[1])

        namespace_prefix = f"{self._namespace}."
        for candidate in list(candidates):
            if candidate.startswith(namespace_prefix):
                add_candidate(candidate[len(namespace_prefix) :])

        for candidate in candidates:
            detail = self._tool_details.get(candidate)
            if detail is not None:
                return detail

        for candidate in candidates:
            for detail in self._tool_details.values():
                if detail.access_pattern == candidate:
                    return detail
        return None

    def _make_print(self, logs: list[str]) -> Callable[..., None]:
        def emit(*args: Any, **kwargs: Any) -> None:
            sep = kwargs.get("sep", " ")
            end = kwargs.get("end", "\n")
            text = sep.join(str(arg) for arg in args) + end
            for line in text.rstrip("\n").splitlines():
                logs.append(line)

        return emit


async def build_code_mode_runtime(
    mcp: Any,
    *,
    excluded_tools: Sequence[str] = (),
    tool_provider: Callable[[], Sequence[Any] | Awaitable[Sequence[Any]]] | None = None,
) -> CodeModeRuntime:
    """Helper for callers that want a ready-to-use runtime."""

    runtime = CodeModeRuntime(
        mcp,
        excluded_tools=excluded_tools,
        tool_provider=tool_provider,
    )
    await runtime.refresh()
    return runtime
