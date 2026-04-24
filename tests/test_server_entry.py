import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SERVER_PY = REPO_ROOT / "src" / "microsoft_mcp" / "server.py"


def test_server_py_script_form_imports_cleanly(monkeypatch):
    """server.py must work when run as `python src/microsoft_mcp/server.py`.

    Guards against regression of the relative-import footgun. We invoke
    with MICROSOFT_MCP_CLIENT_ID unset so main() exits with code 1 AFTER
    the imports succeed — the test asserts we got past import time.
    """
    env = {"PATH": "/usr/bin:/bin", "HOME": "/tmp"}
    proc = subprocess.run(
        [sys.executable, str(SERVER_PY)],
        env=env,
        capture_output=True,
        text=True,
        timeout=15,
    )
    # Expected: exit 1 with "MICROSOFT_MCP_CLIENT_ID ... required" on stderr.
    # Failure mode we're guarding against: ImportError on `from .tools import`.
    assert "ImportError" not in proc.stderr, (
        f"server.py raised ImportError; script-form invocation broken.\n"
        f"stderr: {proc.stderr}"
    )
    # Sanity: reached the CLIENT_ID check.
    assert "MICROSOFT_MCP_CLIENT_ID" in proc.stderr, (
        f"server.py did not reach the CLIENT_ID guard; unexpected output: {proc.stderr}"
    )
