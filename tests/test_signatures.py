"""Tests for the local signature file store and renderer."""

from __future__ import annotations

import pytest

from microsoft_mcp import signatures


@pytest.fixture(autouse=True)
def _isolate_signatures_dir(tmp_path, monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURES_DIR", str(tmp_path))
    monkeypatch.delenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_ACCOUNT_ID", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_SIGNATURE_RFC3676", raising=False)
    yield


# --- account_slug --------------------------------------------------------


def test_account_slug_uses_signature_account_env(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "Brian-Work")
    assert signatures.account_slug() == "brian-work"


def test_account_slug_overrides_env(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    assert signatures.account_slug(override="jp-work") == "jp-work"


def test_account_slug_slugifies_email(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_ACCOUNT_ID", "brian@work.com")
    assert signatures.account_slug() == "brian-work-com"


def test_account_slug_requires_some_account(monkeypatch):
    monkeypatch.delenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", raising=False)
    monkeypatch.delenv("MICROSOFT_MCP_ACCOUNT_ID", raising=False)
    with pytest.raises(signatures.NoAccountError):
        signatures.account_slug()


# --- path construction ---------------------------------------------------


def test_signature_path_lowercases_name(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    p = signatures.signature_path("Default")
    assert p.name == "brian-work-default.txt"


def test_signature_path_html_variant(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    p = signatures.signature_path("default", html=True)
    assert p.name == "brian-work-default.html"


def test_signature_name_rejects_invalid_chars(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    with pytest.raises(signatures.InvalidSignatureNameError):
        signatures.signature_path("bad name!")


def test_signature_name_rejects_none_sentinel(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    with pytest.raises(signatures.InvalidSignatureNameError):
        signatures.signature_path("none")


# --- store round-trip ----------------------------------------------------


def test_write_read_delete_round_trip(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    signatures.write_signature("default", "Cheers,\nBrian")
    assert signatures.read_signature("default") == "Cheers,\nBrian"
    assert signatures.delete_signature("default") is True
    assert signatures.read_signature("default") is None
    assert signatures.delete_signature("default") is False


def test_list_signatures_groups_html_siblings(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    signatures.write_signature("default", "Cheers,\nBrian")
    signatures.write_signature("default", "<b>Cheers</b>", html=True)
    signatures.write_signature("replies", "Brian")

    rows = signatures.list_signatures()
    by_name = {r.name: r for r in rows}
    assert set(by_name) == {"default", "replies"}
    assert by_name["default"].has_html is True
    assert by_name["replies"].has_html is False


def test_list_signatures_all_accounts(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    signatures.write_signature("default", "brian")
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "jp-work")
    signatures.write_signature("default", "jp")

    rows = signatures.list_signatures(account="*")
    pairs = {(r.account, r.name) for r in rows}
    assert pairs == {("brian-work", "default"), ("jp-work", "default")}


# --- apply_signature -----------------------------------------------------


def test_apply_signature_text_default_separator(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    signatures.write_signature("default", "Cheers,\nBrian")
    out = signatures.apply_signature("Hi there", "text", "default")
    assert out == "Hi there\n\nCheers,\nBrian"


def test_apply_signature_text_rfc3676(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_RFC3676", "1")
    signatures.write_signature("default", "Cheers,\nBrian")
    out = signatures.apply_signature("Hi there", "text", "default")
    assert out == "Hi there\n\n-- \nCheers,\nBrian"


def test_apply_signature_html_uses_sibling(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    signatures.write_signature("default", "Brian (text)")
    signatures.write_signature("default", "<b>Brian</b>", html=True)
    out = signatures.apply_signature("<p>Hi</p>", "html", "default")
    assert out == "<p>Hi</p>\n\n<b>Brian</b>"


def test_apply_signature_html_falls_back_to_text_with_conversion(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    signatures.write_signature("default", "Cheers,\nBrian & co")
    out = signatures.apply_signature("<p>Hi</p>", "html", "default")
    assert '<div class="signature">' in out
    # < and > escaped, ampersand HTML-escaped, newlines converted.
    assert "Cheers,<br>\nBrian &amp; co" in out


def test_apply_signature_none_sentinel_is_noop(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    # No file written; "none" must not trigger a read.
    assert signatures.apply_signature("Hi", "text", "none") == "Hi"
    assert signatures.apply_signature(None, "text", "NONE") == ""


def test_apply_signature_missing_raises(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    with pytest.raises(signatures.SignatureNotFoundError) as exc_info:
        signatures.apply_signature("Hi", "text", "default")
    assert exc_info.value.account == "brian-work"
    assert exc_info.value.name == "default"


def test_apply_signature_empty_body_yields_signature_only(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    signatures.write_signature("default", "Cheers,\nBrian")
    assert signatures.apply_signature("", "text", "default") == "Cheers,\nBrian"
    assert signatures.apply_signature(None, "text", "default") == "Cheers,\nBrian"


def test_is_none():
    assert signatures.is_none("none")
    assert signatures.is_none("None")
    assert signatures.is_none("  NONE  ")
    assert not signatures.is_none("")
    assert not signatures.is_none(None)
    assert not signatures.is_none("default")
