import sys
import pytest


def test_shim_maps_verify_to_auth_cli(monkeypatch):
    import auth_refresh

    captured = {}

    def fake_main(argv):
        captured["argv"] = argv
        return 0

    monkeypatch.setattr("microsoft_mcp.auth_cli.main", fake_main)
    monkeypatch.setattr(sys, "argv", ["auth_refresh.py", "--verify", "--live"])
    with pytest.raises(SystemExit) as exc:
        auth_refresh.main()
    assert captured["argv"] == ["verify", "--live"]
    assert exc.value.code == 0


def test_shim_maps_force_email(monkeypatch):
    import auth_refresh

    captured = {}
    monkeypatch.setattr(
        "microsoft_mcp.auth_cli.main",
        lambda argv: captured.setdefault("argv", argv) or 0,
    )
    monkeypatch.setattr(sys, "argv", ["auth_refresh.py", "--force", "broach@cresa.com"])
    with pytest.raises(SystemExit):
        auth_refresh.main()
    assert captured["argv"] == ["refresh", "broach@cresa.com", "--force"]


def test_shim_maps_bare_refresh_all(monkeypatch):
    import auth_refresh

    captured = {}
    monkeypatch.setattr(
        "microsoft_mcp.auth_cli.main",
        lambda argv: captured.setdefault("argv", argv) or 0,
    )
    monkeypatch.setattr(sys, "argv", ["auth_refresh.py"])
    with pytest.raises(SystemExit):
        auth_refresh.main()
    assert captured["argv"] == ["refresh"]
