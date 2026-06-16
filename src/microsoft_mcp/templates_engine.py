"""
YAML template loader and renderer for Microsoft MCP.

Search path (user dir first, then bundled):
- $MICROSOFT_MCP_TEMPLATES_DIR or ~/.config/microsoft-mcp/templates/
- <this_module>/templates_data/
"""

from __future__ import annotations

import html
import os
import re
from pathlib import Path
from typing import Any

import yaml

_BUILTIN_DATA = Path(__file__).parent / "templates_data"


def template_dirs() -> list[Path]:
    """Return the template search path: user dir first, then bundled."""
    user_root_str = os.environ.get("MICROSOFT_MCP_TEMPLATES_DIR")
    user_root = (
        Path(user_root_str)
        if user_root_str
        else Path.home() / ".config" / "microsoft-mcp" / "templates"
    )
    dirs: list[Path] = []
    if user_root.exists():
        dirs.append(user_root)
    if _BUILTIN_DATA.exists():
        dirs.append(_BUILTIN_DATA)
    return dirs


def _category_dirs(category: str | None) -> list[tuple[Path, str]]:
    roots = template_dirs()
    result: list[tuple[Path, str]] = []
    for idx, root in enumerate(roots):
        source = "user" if (idx == 0 and root != _BUILTIN_DATA) else "builtin"
        if category is None:
            if root.exists():
                for subdir in sorted(root.iterdir()):
                    if subdir.is_dir():
                        result.append((subdir, source))
        else:
            candidate = root / category
            if candidate.exists():
                result.append((candidate, source))
    return result


def list_templates(category: str | None = None) -> list[dict[str, Any]]:
    """List available templates, optionally filtered by category."""
    templates: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()

    for cat_dir, source in _category_dirs(category):
        cat_name = cat_dir.name
        for yaml_file in sorted(cat_dir.glob("*.yaml")):
            if yaml_file.name.startswith("_"):
                continue
            tpl_name = yaml_file.stem
            key = (cat_name, tpl_name)
            if key in seen:
                continue
            seen.add(key)
            try:
                with yaml_file.open(encoding="utf-8") as fh:
                    data: dict[str, Any] = yaml.safe_load(fh) or {}
            except (yaml.YAMLError, OSError):
                continue
            placeholders = [
                p.get("name")
                for p in data.get("placeholders", [])
                if isinstance(p, dict)
            ]
            templates.append(
                {
                    "name": tpl_name,
                    "description": data.get("description", ""),
                    "version": data.get("version", "1.0"),
                    "category": cat_name,
                    "source": source,
                    "placeholders": placeholders,
                }
            )

    templates.sort(key=lambda t: (t["category"], t["name"]))
    return templates


def load_template(category: str, name: str) -> dict[str, Any]:
    """Load and validate a template by category and name. Raises ValueError if not found or invalid."""
    for cat_dir, _source in _category_dirs(category):
        candidate = cat_dir / f"{name}.yaml"
        if candidate.exists():
            try:
                with candidate.open(encoding="utf-8") as fh:
                    tpl: dict[str, Any] = yaml.safe_load(fh) or {}
            except yaml.YAMLError as exc:
                raise ValueError(f"Invalid YAML in template '{name}': {exc}") from exc
            if not tpl.get("name"):
                raise ValueError(
                    f"Template '{name}' (category '{category}') missing required field: name"
                )
            if not tpl.get("html_template"):
                raise ValueError(
                    f"Template '{name}' (category '{category}') missing required field: html_template"
                )
            tpl["_path"] = str(candidate)
            tpl["_category"] = category
            return tpl
    raise ValueError(f"Template '{name}' not found in category '{category}'")


def validate_template_data(tpl: dict[str, Any], data: dict[str, Any]) -> list[str]:
    """Return list of error strings for missing required placeholders (empty = valid)."""
    errors: list[str] = []
    for ph in tpl.get("placeholders", []):
        if (
            isinstance(ph, dict)
            and ph.get("required")
            and not data.get(ph.get("name", ""))
        ):
            errors.append(f"Missing required placeholder: {ph['name']}")
    return errors


def render_template(category: str, name: str, data: dict[str, Any]) -> str:
    """Load and render a template. Raises ValueError on validation failure."""
    tpl = load_template(category, name)
    errors = validate_template_data(tpl, data)
    if errors:
        raise ValueError(
            "Template validation failed:\n" + "\n".join(f"  - {e}" for e in errors)
        )
    conditional_values = _render_conditional_sections(tpl, data)
    render_data = {**data, **conditional_values}
    rendered = _substitute_placeholders(tpl.get("html_template", ""), render_data, tpl)
    lines = [line for line in rendered.split("\n") if line.strip()]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

_PRE_RENDERED_KEYS: frozenset[str] = frozenset(
    {
        "agenda_items",
        "interviewer_items",
        "focus_area_items",
        "action_item_list",
        "format_items",
        "meta_items",
        "candidate_info",
        "client_info_items",
        "severity_status_info",
        "notes_content",
    }
)

_PLACEHOLDER_RE = re.compile(r"\{(\w+)\}")


def _escape(value: str) -> str:
    return html.escape(str(value))


def _parse_comma_list(value: str | list[str]) -> list[str]:
    if isinstance(value, list):
        return value
    return [item.strip() for item in str(value).split(",") if item.strip()]


def _render_list_items(
    items: list[str], numbered: bool = True, bullet: str = "•"
) -> str:
    rows: list[str] = []
    for i, item in enumerate(items, 1):
        esc = _escape(item.strip())
        rows.append(
            f"<tr><td>{i}. {esc}</td></tr>"
            if numbered
            else f"<tr><td>{bullet} {esc}</td></tr>"
        )
    return "\n".join(rows)


def _evaluate_condition(condition: str, data: dict[str, Any]) -> bool:
    if "|" in condition:
        return any(data.get(f.strip()) for f in condition.split("|"))
    if "&" in condition:
        return all(data.get(f.strip()) for f in condition.split("&"))
    return bool(data.get(condition.strip()))


def _raw_substitute(s: str, data: dict[str, Any]) -> str:
    return _PLACEHOLDER_RE.sub(lambda m: str(data.get(m.group(1), "")), s)


def _render_conditional_sections(
    tpl: dict[str, Any], data: dict[str, Any]
) -> dict[str, str]:
    sections: dict[str, str] = {}
    for sname, sdef in tpl.get("conditional_sections", {}).items():
        if not isinstance(sdef, dict):
            sections[sname] = ""
            continue
        if _evaluate_condition(sdef.get("condition", ""), data):
            sections[sname] = _raw_substitute(sdef.get("template", ""), data)
        else:
            sections[sname] = ""
    return sections


def _substitute_placeholders(
    html_str: str, data: dict[str, Any], tpl: dict[str, Any]
) -> str:
    defaults: dict[str, Any] = {
        ph["name"]: ph["default"]
        for ph in tpl.get("placeholders", [])
        if isinstance(ph, dict) and "default" in ph
    }
    subs: dict[str, Any] = {**defaults, **data}

    # Pre-render list fields
    for src, (dest, numbered, bullet) in {
        "agenda": ("agenda_items", True, "•"),
        "interviewers": ("interviewer_items", False, "•"),
        "focus_areas": ("focus_area_items", False, "•"),
        "action_items": ("action_item_list", False, "•"),
        "format": ("format_items", True, "•"),
    }.items():
        if subs.get(src):
            subs[dest] = _render_list_items(
                _parse_comma_list(subs[src]), numbered=numbered, bullet=bullet
            )

    # Keys that must NOT be HTML-escaped (pre-rendered HTML or conditional outputs)
    pre_rendered = set(_PRE_RENDERED_KEYS) | set(tpl.get("conditional_sections", {}))

    safe: dict[str, str] = {
        k: (str(v) if k in pre_rendered or not isinstance(v, str) else _escape(v))
        for k, v in subs.items()
        if v is not None
    }

    return _PLACEHOLDER_RE.sub(lambda m: safe.get(m.group(1), ""), html_str)


# ---------------------------------------------------------------------------
# {{var}} variable substitution + CSV recipients (task 7.2)
# ---------------------------------------------------------------------------

# Matches {{word_chars}} in plain text
_DOUBLE_BRACE_RE = re.compile(r"\{\{(\w+)\}\}")

# Matches HTML-encoded variants: &#123;&#123;name&#125;&#125; (and &#x7b;/&#x7d; forms)
_HTML_ENCODED_VAR_RE = re.compile(
    r"(?:&#123;|&#x7b;|\{)(?:&#123;|&#x7b;|\{)"
    r"(\w+)"
    r"(?:&#125;|&#x7d;|\})(?:&#125;|&#x7d;|\})"
)


class VariableSubstitutionError(Exception):
    """Raised when strict mode is enabled and a referenced variable is missing."""


def find_template_variables(content: str, decode_html: bool = True) -> list[str]:
    """Return unique ``{{var}}`` names in first-appearance order.

    Args:
        content: Plain-text or HTML content to scan.
        decode_html: When *True*, also detect HTML-encoded variants such as
            ``&#123;&#123;name&#125;&#125;``.

    Returns:
        Ordered list of unique variable names found.
    """
    if not content:
        return []

    seen: set[str] = set()
    result: list[str] = []

    for m in _DOUBLE_BRACE_RE.finditer(content):
        name = m.group(1)
        if name not in seen:
            seen.add(name)
            result.append(name)

    if decode_html:
        for m in _HTML_ENCODED_VAR_RE.finditer(content):
            name = m.group(1)
            if name not in seen:
                seen.add(name)
                result.append(name)

    return result


def substitute_variables(
    content: str,
    values: dict[str, str],
    strict: bool = False,
) -> str:
    """Replace ``{{var}}`` tokens (and HTML-encoded equivalents) with *values*.

    Args:
        content: Content containing ``{{var}}`` tokens.
        values: Mapping of variable name → replacement string.
        strict: When *True*, raise :class:`VariableSubstitutionError` if any
            referenced variable is absent from *values*; otherwise leave the
            token unchanged.

    Returns:
        Content with variables substituted.

    Raises:
        VariableSubstitutionError: If *strict* is ``True`` and a variable has
            no matching entry in *values*.
    """
    if not content:
        return content

    missing: list[str] = []

    def _replace(m: re.Match[str]) -> str:
        var = m.group(1)
        if var in values:
            return values[var]
        missing.append(var)
        return m.group(0)  # leave unchanged

    result = _DOUBLE_BRACE_RE.sub(_replace, content)
    result = _HTML_ENCODED_VAR_RE.sub(_replace, result)

    if strict and missing:
        unique_missing = list(dict.fromkeys(missing))
        raise VariableSubstitutionError(
            f"Missing values for template variables: {', '.join(unique_missing)}"
        )

    return result


def parse_recipients_csv(path: str) -> list[dict[str, str]]:
    """Read a CSV file and return one dict per row keyed by header.

    Encoding is attempted in order: ``utf-8``, ``utf-8-sig``, ``latin-1``.

    Args:
        path: Filesystem path to the CSV file.

    Returns:
        List of row dicts; keys are stripped column headers, values are
        stripped cell strings.

    Raises:
        FileNotFoundError: If *path* does not exist.
        ValueError: If none of the supported encodings can decode the file.
    """
    import csv
    from pathlib import Path as _Path

    csv_path = _Path(path)
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file not found: {path}")

    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            with csv_path.open(encoding=encoding, newline="") as fh:
                reader = csv.DictReader(fh)
                rows: list[dict[str, str]] = []
                for row in reader:
                    clean = {
                        k.strip(): (v.strip() if v else "") for k, v in row.items() if k
                    }
                    rows.append(clean)
            return rows
        except UnicodeDecodeError:
            continue

    raise ValueError(f"Could not decode CSV file with supported encodings: {path}")
