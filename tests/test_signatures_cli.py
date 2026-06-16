"""Tests for the signatures CLI."""

from __future__ import annotations

import io
import json

import pytest

from microsoft_mcp import signatures, signatures_cli


@pytest.fixture(autouse=True)
def _isolate(tmp_path, monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURES_DIR", str(tmp_path))
    monkeypatch.setenv("MICROSOFT_MCP_SIGNATURE_ACCOUNT", "brian-work")
    monkeypatch.delenv("MICROSOFT_MCP_SIGNATURE_RFC3676", raising=False)
    yield


def _run(argv, monkeypatch=None, stdin: str | None = None):
    if stdin is not None:
        monkeypatch.setattr("sys.stdin", io.StringIO(stdin))
    return signatures_cli.main(argv)


def test_dir_prints_resolved_directory(capsys):
    rc = _run(["dir"])
    assert rc == 0
    out = capsys.readouterr().out.strip()
    assert out == str(signatures.resolve_dir())


def test_path_with_name(capsys, monkeypatch):
    rc = _run(["path", "default"])
    assert rc == 0
    expected = signatures.signature_path("default")
    assert capsys.readouterr().out.strip() == str(expected)


def test_path_without_name_returns_dir(capsys):
    rc = _run(["path"])
    assert rc == 0
    assert capsys.readouterr().out.strip() == str(signatures.resolve_dir())


def test_set_from_file_and_show(tmp_path, capsys, monkeypatch):
    src = tmp_path / "src.txt"
    src.write_text("Cheers,\nBrian", encoding="utf-8")

    rc = _run(["set", "default", "--from-file", str(src)])
    assert rc == 0
    capsys.readouterr()  # discard "wrote ..." line

    rc = _run(["show", "default"])
    assert rc == 0
    out = capsys.readouterr().out
    assert "Cheers,\nBrian" in out


def test_set_from_stdin(monkeypatch, capsys):
    rc = _run(["set", "default", "--stdin"], monkeypatch=monkeypatch, stdin="Brian")
    assert rc == 0
    assert signatures.read_signature("default") == "Brian"


def test_set_editor_path_uses_editor_env(monkeypatch, capsys, tmp_path):
    # Use `tee` as a stand-in editor: it copies stdin into the temp file when
    # invoked as `tee <path>`. Combined with a piped stdin, we can simulate
    # an editor writing content.
    fake_editor = tmp_path / "fake_editor.sh"
    fake_editor.write_text(
        '#!/usr/bin/env bash\nprintf "Edited content\\n" > "$1"\n',
        encoding="utf-8",
    )
    fake_editor.chmod(0o755)

    monkeypatch.setenv("EDITOR", str(fake_editor))
    monkeypatch.delenv("VISUAL", raising=False)

    rc = _run(["set", "default", "--editor"])
    assert rc == 0
    assert signatures.read_signature("default") == "Edited content\n"


def test_list_table_and_json(capsys):
    signatures.write_signature("default", "Brian")
    signatures.write_signature("replies", "Brian")

    rc = _run(["list"])
    assert rc == 0
    out = capsys.readouterr().out
    assert "default" in out and "replies" in out

    rc = _run(["list", "--json"])
    assert rc == 0
    data = json.loads(capsys.readouterr().out)
    names = {row["name"] for row in data}
    assert names == {"default", "replies"}


def test_rm_with_yes_flag(capsys):
    signatures.write_signature("default", "Brian")
    rc = _run(["rm", "default", "--yes"])
    assert rc == 0
    assert signatures.read_signature("default") is None


def test_rm_missing_returns_error(capsys):
    rc = _run(["rm", "default", "--yes"])
    assert rc == 1
    err = capsys.readouterr().err
    assert "not found" in err.lower()


def test_show_missing_returns_error(capsys):
    rc = _run(["show", "default"])
    assert rc == 1
    err = capsys.readouterr().err
    assert "not found" in err.lower()


def test_account_override_writes_to_other_slug(capsys):
    rc = _run(["set", "default", "--account", "jp-work", "--from-file", "/dev/null"])
    assert rc == 0
    # Stored under jp-work, not brian-work.
    p = signatures.signature_path("default", account="jp-work")
    assert p.exists()
    assert not signatures.signature_path("default", account="brian-work").exists()


def test_invalid_name_returns_error(capsys):
    rc = _run(["path", "bad name!"])
    assert rc == 1
    assert "invalid signature name" in capsys.readouterr().err.lower()
