import shutil
import subprocess
import sys
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent
SERVER_PY = REPO_ROOT / "src" / "microsoft_mcp" / "server.py"

# Resolve the console script once at module load so the subprocess env can be
# stripped down without needing a PATH that reaches the venv's bin directory.
CONSOLE_SCRIPT = shutil.which("microsoft-mcp")


def _run_server_form(
    argv: list[str], cwd: str | None = None
) -> subprocess.CompletedProcess[str]:
    """Invoke a server entry-point form and return the completed process."""
    env = {"PATH": "/usr/bin:/bin", "HOME": "/tmp"}
    return subprocess.run(
        argv,
        env=env,
        cwd=cwd,
        capture_output=True,
        text=True,
        timeout=15,
    )


def _assert_clean_import(proc: subprocess.CompletedProcess[str], label: str) -> None:
    """Assert that the server process reached the CLIENT_ID guard without ImportError."""
    assert "ImportError" not in proc.stderr, (
        f"{label} raised ImportError; invocation broken.\nstderr: {proc.stderr}"
    )
    assert "MICROSOFT_MCP_CLIENT_ID" in proc.stderr, (
        f"{label} did not reach the CLIENT_ID guard; unexpected output: {proc.stderr}"
    )


def test_server_py_script_form_imports_cleanly():
    """server.py must work when run as `python src/microsoft_mcp/server.py`.

    Guards against regression of the relative-import footgun. We invoke
    with MICROSOFT_MCP_CLIENT_ID unset so main() exits with code 1 AFTER
    the imports succeed — the test asserts we got past import time.
    """
    proc = _run_server_form([sys.executable, str(SERVER_PY)])
    # Expected: exit 1 with "MICROSOFT_MCP_CLIENT_ID ... required" on stderr.
    # Failure mode we're guarding against: ImportError on `from .tools import`.
    _assert_clean_import(proc, "script form (python src/microsoft_mcp/server.py)")


def test_server_module_form_imports_cleanly():
    """server.py must work when run as `python -m microsoft_mcp.server`.

    The venv interpreter already has microsoft_mcp on sys.path via site-packages,
    so no PYTHONPATH manipulation is needed.
    """
    proc = _run_server_form([sys.executable, "-m", "microsoft_mcp.server"])
    _assert_clean_import(proc, "module form (python -m microsoft_mcp.server)")


@pytest.mark.skipif(
    CONSOLE_SCRIPT is None, reason="microsoft-mcp console script not on PATH"
)
def test_server_console_script_imports_cleanly():
    """microsoft-mcp console script must reach the CLIENT_ID guard without ImportError."""
    proc = _run_server_form([CONSOLE_SCRIPT])  # type: ignore[list-item]
    _assert_clean_import(proc, "console script (microsoft-mcp)")


def test_auth_subcommand_dispatches_to_auth_cli(monkeypatch):
    from unittest import mock

    from microsoft_mcp import server

    called = {}

    def fake_main(argv):
        called["argv"] = argv
        return 0

    monkeypatch.setattr("microsoft_mcp.auth_cli.main", fake_main)
    monkeypatch.setattr(sys, "argv", ["microsoft-mcp", "auth", "refresh", "--json"])
    # Real sys.exit raises SystemExit and halts; mirror that so the dispatch
    # actually short-circuits before reaching load_dotenv()/mcp.run().
    with mock.patch.object(sys, "exit", side_effect=SystemExit) as fake_exit:
        with pytest.raises(SystemExit):
            server.main()
    assert called["argv"] == ["refresh", "--json"]
    fake_exit.assert_called_once_with(0)
