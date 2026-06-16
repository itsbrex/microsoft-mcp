"""Local plain-text email signature store and renderer.

Microsoft Graph does not expose Outlook signature settings, so signatures
are managed entirely client-side: stored as files under
``~/.config/microsoft-mcp/signatures/`` and appended to draft bodies before
they are POST/PATCHed to Graph.

File naming: ``<account-slug>-<signature-name>.txt`` with an optional
``<account-slug>-<signature-name>.html`` sibling used verbatim for HTML
drafts.
"""

from __future__ import annotations

import html as _html
import os
import re
from dataclasses import dataclass
from pathlib import Path

__all__ = [
    "SignatureInfo",
    "SignatureNotFoundError",
    "InvalidSignatureNameError",
    "NoAccountError",
    "NONE_SENTINEL",
    "resolve_dir",
    "account_slug",
    "signature_path",
    "list_signatures",
    "read_signature",
    "write_signature",
    "delete_signature",
    "apply_signature",
    "is_none",
]


NONE_SENTINEL = "none"

_DEFAULT_DIR = Path.home() / ".config" / "microsoft-mcp" / "signatures"

# Signature names: lowercase letters, digits, dot, underscore, dash.
_NAME_RE = re.compile(r"^[a-z0-9._-]+$")
# Account slugs follow the same rule (after slugification).
_SLUG_INVALID_RE = re.compile(r"[^a-z0-9-]+")
_SLUG_DASH_COLLAPSE_RE = re.compile(r"-+")


class SignatureNotFoundError(FileNotFoundError):
    def __init__(self, name: str, account: str, html: bool) -> None:
        self.name = name
        self.account = account
        self.html = html
        suffix = ".html" if html else ".txt"
        super().__init__(f"signature not found: {account}/{name}{suffix}")


class InvalidSignatureNameError(ValueError):
    pass


class NoAccountError(ValueError):
    pass


@dataclass(frozen=True)
class SignatureInfo:
    account: str
    name: str
    path: Path
    has_html: bool
    size: int
    modified: float


# --- environment / path resolution ---------------------------------------


def _rfc3676_enabled() -> bool:
    return os.getenv("MICROSOFT_MCP_SIGNATURE_RFC3676", "").strip().lower() in {
        "1",
        "true",
        "yes",
        "on",
    }


def resolve_dir() -> Path:
    """Return the directory where signature files live.

    Honors ``MICROSOFT_MCP_SIGNATURES_DIR`` and falls back to
    ``~/.config/microsoft-mcp/signatures/``.
    """
    override = os.getenv("MICROSOFT_MCP_SIGNATURES_DIR")
    return Path(override).expanduser() if override else _DEFAULT_DIR


def _slugify_email(value: str) -> str:
    lowered = value.strip().lower()
    replaced = lowered.replace("@", "-").replace(".", "-")
    cleaned = _SLUG_INVALID_RE.sub("-", replaced)
    collapsed = _SLUG_DASH_COLLAPSE_RE.sub("-", cleaned).strip("-")
    return collapsed


def account_slug(
    account_id: str | None = None,
    override: str | None = None,
) -> str:
    """Resolve the account slug used in signature filenames.

    Priority:
      1. *override* argument (sanitized — lowercased, dash-cleaned).
      2. ``MICROSOFT_MCP_SIGNATURE_ACCOUNT`` env var (sanitized).
      3. *account_id* argument or ``MICROSOFT_MCP_ACCOUNT_ID`` env var,
         slugified from email form (``brian@work.com`` →
         ``brian-work-com``).

    Raises ``NoAccountError`` when nothing resolves.
    """
    for candidate in (override, os.getenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT")):
        if candidate and candidate.strip():
            slug = _slugify_email(candidate)
            if slug:
                return slug

    raw = (
        account_id if account_id is not None else os.getenv("MICROSOFT_MCP_ACCOUNT_ID")
    )
    if raw and raw.strip():
        slug = _slugify_email(raw)
        if slug:
            return slug

    raise NoAccountError(
        "no signature account; set MICROSOFT_MCP_SIGNATURE_ACCOUNT, "
        "MICROSOFT_MCP_ACCOUNT_ID, or pass account="
    )


def _normalize_name(name: str) -> str:
    lowered = (name or "").strip().lower()
    if not lowered:
        raise InvalidSignatureNameError("signature name cannot be empty")
    if lowered == NONE_SENTINEL:
        raise InvalidSignatureNameError(
            "'none' is a reserved sentinel meaning 'do not inject'"
        )
    if not _NAME_RE.match(lowered):
        raise InvalidSignatureNameError(
            f"invalid signature name: {name!r} (allowed: [a-z0-9._-]+)"
        )
    return lowered


def is_none(name: str | None) -> bool:
    return bool(name and name.strip().lower() == NONE_SENTINEL)


def signature_path(
    name: str,
    *,
    account: str | None = None,
    html: bool = False,
) -> Path:
    """Return the resolved file path for a signature (file may not exist)."""
    normalized = _normalize_name(name)
    slug = account_slug(override=account)
    suffix = ".html" if html else ".txt"
    return resolve_dir() / f"{slug}-{normalized}{suffix}"


# --- store I/O -----------------------------------------------------------


def list_signatures(account: str | None = None) -> list[SignatureInfo]:
    """List signatures.

    ``account`` may be:
      - ``None``: use the resolved slug for the current env.
      - a specific slug/email-ish string: list signatures for that slug.
      - ``"*"`` or ``"all"`` (case-insensitive): list signatures across
        every account found in the directory.
    """
    directory = resolve_dir()
    if not directory.exists():
        return []

    if account is not None and account.strip().lower() in {"*", "all"}:
        wanted_slug = None
    else:
        wanted_slug = account_slug(override=account)

    # Group .txt entries with sibling .html detection.
    by_key: dict[tuple[str, str], SignatureInfo] = {}
    html_keys: set[tuple[str, str]] = set()

    for entry in sorted(directory.iterdir()):
        if not entry.is_file():
            continue
        stem = entry.stem
        suffix = entry.suffix.lower()
        if suffix not in (".txt", ".html"):
            continue
        # stem is "<slug>-<name>"; split on last dash before the name.
        # Because both slugs and names may contain dashes, we split on the
        # first dash that produces a valid (non-empty) name match — but the
        # simpler rule: account slug is everything up to the last dash, name
        # is the remainder. Names contain only [a-z0-9._-], same as slugs,
        # so the split is genuinely ambiguous. We instead require that the
        # file's stem starts with the resolved slug + "-" when filtering.
        if "-" not in stem:
            continue
        if wanted_slug is not None:
            if not stem.startswith(wanted_slug + "-"):
                continue
            slug = wanted_slug
            sig_name = stem[len(slug) + 1 :]
        else:
            # Best-effort split for "all" listings: rightmost segment after
            # the last dash is the name; everything before is the slug.
            slug, _, sig_name = stem.rpartition("-")
            if not slug or not sig_name:
                continue
        if not sig_name:
            continue
        key = (slug, sig_name)
        if suffix == ".html":
            html_keys.add(key)
            # Track HTML-only signatures too, so they show up in listings.
            if key not in by_key:
                stat = entry.stat()
                by_key[key] = SignatureInfo(
                    account=slug,
                    name=sig_name,
                    path=entry,
                    has_html=True,
                    size=stat.st_size,
                    modified=stat.st_mtime,
                )
            continue

        stat = entry.stat()
        by_key[key] = SignatureInfo(
            account=slug,
            name=sig_name,
            path=entry,
            has_html=False,
            size=stat.st_size,
            modified=stat.st_mtime,
        )

    return [
        SignatureInfo(
            account=info.account,
            name=info.name,
            path=info.path,
            has_html=(info.account, info.name) in html_keys,
            size=info.size,
            modified=info.modified,
        )
        for info in by_key.values()
    ]


def read_signature(
    name: str,
    *,
    account: str | None = None,
    html: bool = False,
) -> str | None:
    """Return signature contents, or ``None`` if the file does not exist."""
    path = signature_path(name, account=account, html=html)
    if not path.exists():
        return None
    return path.read_text(encoding="utf-8")


def write_signature(
    name: str,
    content: str,
    *,
    account: str | None = None,
    html: bool = False,
) -> Path:
    """Persist a signature to disk; creates the directory if needed."""
    path = signature_path(name, account=account, html=html)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")
    return path


def delete_signature(
    name: str,
    *,
    account: str | None = None,
    html: bool = False,
) -> bool:
    """Delete a signature file. Returns True if removed, False if absent."""
    path = signature_path(name, account=account, html=html)
    if not path.exists():
        return False
    path.unlink()
    return True


# --- rendering -----------------------------------------------------------


def _separator(content_type: str) -> str:
    if _rfc3676_enabled():
        if content_type == "html":
            return "<br><br>-- <br>\n"
        return "\n\n-- \n"
    return "\n\n"


def _text_to_html(text: str) -> str:
    escaped = _html.escape(text, quote=False)
    converted = escaped.replace("\n", "<br>\n")
    return f'<div class="signature">{converted}</div>'


def apply_signature(
    body: str | None,
    body_content_type: str,
    name: str,
    *,
    account: str | None = None,
) -> str:
    """Return *body* with the named signature appended.

    Raises ``SignatureNotFoundError`` if the file is missing. Callers that
    want lenient behavior should catch that and surface a warning.
    """
    content_type = (body_content_type or "text").strip().lower()
    if content_type not in {"text", "html"}:
        content_type = "text"

    if is_none(name):
        return body or ""

    normalized = _normalize_name(name)
    slug = account_slug(override=account)

    if content_type == "html":
        html_path = resolve_dir() / f"{slug}-{normalized}.html"
        if html_path.exists():
            sig_content = html_path.read_text(encoding="utf-8")
        else:
            txt_path = resolve_dir() / f"{slug}-{normalized}.txt"
            if not txt_path.exists():
                raise SignatureNotFoundError(normalized, slug, html=False)
            sig_content = _text_to_html(txt_path.read_text(encoding="utf-8"))
    else:
        txt_path = resolve_dir() / f"{slug}-{normalized}.txt"
        if not txt_path.exists():
            raise SignatureNotFoundError(normalized, slug, html=False)
        sig_content = txt_path.read_text(encoding="utf-8")

    if not body:
        return sig_content

    return f"{body}{_separator(content_type)}{sig_content}"


def info_for(name: str, *, account: str | None = None) -> SignatureInfo | None:
    """Return a SignatureInfo for the named signature, or None if missing."""
    slug = account_slug(override=account)
    normalized = _normalize_name(name)
    txt_path = resolve_dir() / f"{slug}-{normalized}.txt"
    html_path = resolve_dir() / f"{slug}-{normalized}.html"
    has_txt = txt_path.exists()
    has_html = html_path.exists()
    if not has_txt and not has_html:
        return None
    primary = txt_path if has_txt else html_path
    stat = primary.stat()
    return SignatureInfo(
        account=slug,
        name=normalized,
        path=primary,
        has_html=has_html,
        size=stat.st_size,
        modified=stat.st_mtime,
    )
