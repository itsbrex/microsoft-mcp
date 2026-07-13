import json
import datetime as dt
from microsoft_mcp import auth_msal

TEST_EMAIL = "broach@cresa.email"


def _seed_graph_account(tmp_path, identifier, valid=True):
    now = dt.datetime.now(dt.timezone.utc)
    delta = 3600 if valid else -10
    exp = now + dt.timedelta(seconds=delta)
    (tmp_path / f"{identifier}_access_token.json").write_text(
        json.dumps(
            {
                "email": identifier,
                "access_token": "a.b.c",
                "token_type": "Bearer",
                "expires_in": delta,
                "expires_at": exp.strftime("%Y-%m-%dT%H:%M:%SZ"),
                "scopes": auth_msal.GRAPH_SCOPE,
                "api_type": "graph",
            }
        )
    )
    (tmp_path / f"{identifier}_refresh_only.txt").write_text("shared-refresh")


def test_refresh_account_outlook_uses_outlook_probe(tmp_path, monkeypatch):
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)

    captured = {}

    def fake_refresh(self, refresh_token):
        captured["api_type"] = self.api_type
        captured["scope_default"] = self._default_scope()
        return {
            "access_token": "new",
            "refresh_token": "shared-refresh",
            "expires_in": 3600,
            "scope": auth_msal.OUTLOOK_SCOPE,
        }

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth, "_refresh_access_token", fake_refresh
    )
    result = auth_msal.refresh_account(
        TEST_EMAIL, tokens_dir=tmp_path, api_type="outlook"
    )
    assert result["status"] == "refreshed"
    assert captured["api_type"] == "outlook"
    assert (tmp_path / f"{TEST_EMAIL}_outlook_access_token.json").exists()


def test_refresh_all_accounts_both_returns_two_entries(tmp_path, monkeypatch):
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)
    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth,
        "_refresh_access_token",
        lambda self, rt: {
            "access_token": "n",
            "refresh_token": "shared-refresh",
            "expires_in": 3600,
            "scope": self._default_scope(),
        },
    )
    results = auth_msal.refresh_all_accounts(tokens_dir=tmp_path, api_type="both")
    api_types = sorted(r["api_type"] for r in results)
    assert api_types == ["graph", "outlook"]


def test_refresh_all_accounts_default_is_graph_only(tmp_path):
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=True)
    results = auth_msal.refresh_all_accounts(tokens_dir=tmp_path)
    assert len(results) == 1
    assert results[0].get("api_type", "graph") == "graph"


def test_refresh_all_accounts_skips_outlook_sibling_files(tmp_path):
    # an outlook sibling file must not be enumerated as its own account
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=True)
    (tmp_path / f"{TEST_EMAIL}_outlook_access_token.json").write_text("{}")
    results = auth_msal.refresh_all_accounts(tokens_dir=tmp_path)
    idents = [r["identifier"] for r in results]
    assert idents == [TEST_EMAIL]  # not TEST_EMAIL_outlook


def test_interactive_auth_allowed_default_and_opt_out(monkeypatch):
    monkeypatch.delenv("MICROSOFT_MCP_NONINTERACTIVE", raising=False)
    assert auth_msal._interactive_auth_allowed() is True
    for val in ("1", "true", "YES", "on"):
        monkeypatch.setenv("MICROSOFT_MCP_NONINTERACTIVE", val)
        assert auth_msal._interactive_auth_allowed() is False
    monkeypatch.setenv("MICROSOFT_MCP_NONINTERACTIVE", "0")
    assert auth_msal._interactive_auth_allowed() is True


def test_cresa_com_auth_is_blocked_before_token_access(tmp_path, monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS", "cresa.com")
    metadata_loads = 0

    def track_metadata_load(identifier):
        nonlocal metadata_loads
        metadata_loads += 1
        return None

    monkeypatch.setattr(
        auth_msal,
        "_load_outlook_creds_account_metadata",
        track_metadata_load,
    )

    import pytest

    with pytest.raises(RuntimeError, match="blocked for account domain"):
        auth_msal.MSALRefreshTokenAuth(
            tokens_dir=tmp_path,
            account_identifier="protected@cresa.com",
        )

    assert metadata_loads == 0


def test_cresa_com_force_reauth_is_blocked_before_cache_clear(tmp_path, monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS", "cresa.com")
    constructed = False

    class UnexpectedAuth:
        def __init__(self, **kwargs):
            nonlocal constructed
            constructed = True

    monkeypatch.setattr(auth_msal, "MSALRefreshTokenAuth", UnexpectedAuth)

    import pytest

    with pytest.raises(RuntimeError, match="blocked for account domain"):
        auth_msal.force_reauthenticate(
            "protected@cresa.com",
            tokens_dir=tmp_path,
        )

    assert constructed is False


def test_cresa_email_auth_domain_is_allowed(monkeypatch):
    monkeypatch.setenv("MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS", "cresa.com")
    auth_msal._require_account_auth_allowed("allowed@cresa.email")


def test_cresa_com_live_verification_is_blocked_before_token_read(
    tmp_path, monkeypatch
):
    monkeypatch.setenv("MICROSOFT_MCP_BLOCKED_AUTH_DOMAINS", "cresa.com")
    token_file = tmp_path / "protected@cresa.com_access_token.json"
    token_file.write_text("not-json-and-must-not-be-read")

    import pytest

    with pytest.raises(RuntimeError, match="blocked for account domain"):
        auth_msal.verify_account_tokens(tokens_dir=tmp_path, live=True)


def test_acquire_token_data_raises_when_noninteractive_and_no_refresh(
    tmp_path, monkeypatch
):
    # No refresh token on disk + NONINTERACTIVE set → raise instead of hanging
    # on an interactive device-code flow. authenticate() must NOT be called.
    monkeypatch.setenv("MICROSOFT_MCP_NONINTERACTIVE", "1")
    auth = auth_msal.MSALRefreshTokenAuth(
        tokens_dir=tmp_path, account_identifier=TEST_EMAIL
    )
    called = {"auth": False}

    def boom_authenticate(self):
        called["auth"] = True
        return {}

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth, "authenticate", boom_authenticate
    )
    import pytest

    with pytest.raises(RuntimeError, match="MICROSOFT_MCP_NONINTERACTIVE"):
        auth.get_token()
    assert called["auth"] is False


def test_force_refresh_raises_when_noninteractive_and_refresh_fails(
    tmp_path, monkeypatch
):
    monkeypatch.setenv("MICROSOFT_MCP_NONINTERACTIVE", "1")
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)
    auth = auth_msal.MSALRefreshTokenAuth(
        tokens_dir=tmp_path, account_identifier=TEST_EMAIL
    )
    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth,
        "_refresh_access_token",
        lambda self, rt: (_ for _ in ()).throw(RuntimeError("AADSTS65002")),
    )
    called = {"auth": False}
    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth,
        "authenticate",
        lambda self: called.__setitem__("auth", True) or {},
    )
    import pytest

    with pytest.raises(RuntimeError, match="MICROSOFT_MCP_NONINTERACTIVE"):
        auth.force_refresh()
    assert called["auth"] is False


def test_acquire_token_data_preserves_cache_when_noninteractive_refresh_fails(
    tmp_path, monkeypatch
):
    monkeypatch.setenv("MICROSOFT_MCP_NONINTERACTIVE", "1")
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)
    auth = auth_msal.MSALRefreshTokenAuth(
        tokens_dir=tmp_path, account_identifier=TEST_EMAIL
    )
    original_files = {path.name: path.read_bytes() for path in tmp_path.iterdir()}
    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth,
        "_refresh_access_token",
        lambda self, rt: (_ for _ in ()).throw(RuntimeError("network unavailable")),
    )

    import pytest

    with pytest.raises(RuntimeError, match="MICROSOFT_MCP_NONINTERACTIVE"):
        auth.get_token()

    assert {
        path.name: path.read_bytes() for path in tmp_path.iterdir()
    } == original_files


def test_force_reauthenticate_also_outlook_mints_outlook(tmp_path, monkeypatch):
    # After the Graph device-flow re-auth, also_outlook=True must mint an
    # Outlook token off the fresh shared refresh token (silent, no 2nd prompt).
    def fake_authenticate(self):
        self._save_tokens(
            access_token="a.b.c",
            refresh_token="fresh-shared-rt",
            expires_in=3600,
            scopes=auth_msal.GRAPH_SCOPE,
            email=TEST_EMAIL,
        )
        return {"access_token": "a.b.c", "username": TEST_EMAIL}

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth, "authenticate", fake_authenticate
    )
    # The outlook leg does a silent refresh off the shared token.
    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth,
        "_refresh_access_token",
        lambda self, rt: {
            "access_token": "o.o.o",
            "refresh_token": "outlook-rotated-token",
            "expires_in": 3600,
            "scope": auth_msal.OUTLOOK_SCOPE,
        },
    )
    result = auth_msal.force_reauthenticate(
        TEST_EMAIL, tokens_dir=tmp_path, also_outlook=True
    )
    assert result["status"] == "reauthenticated"
    assert "outlook" in result
    assert result["outlook"]["api_type"] == "outlook"
    assert result["outlook"]["status"] in ("refreshed", "valid")
    # Outlook token file written and latest replacement refresh token retained.
    assert (tmp_path / f"{TEST_EMAIL}_outlook_access_token.json").exists()
    assert (
        tmp_path / f"{TEST_EMAIL}_refresh_only.txt"
    ).read_text() == "outlook-rotated-token"


def test_force_reauthenticate_rejects_wrong_account_before_outlook_mint(
    tmp_path, monkeypatch
):
    wrong_email = "other@example.com"

    def fake_authenticate(self):
        self._save_tokens(
            access_token="a.b.c",
            refresh_token="wrong-account-refresh",
            expires_in=3600,
            scopes=auth_msal.GRAPH_SCOPE,
            email=wrong_email,
        )
        return {"access_token": "a.b.c", "username": wrong_email}

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth, "authenticate", fake_authenticate
    )
    refresh_called = False

    def fake_refresh(*args, **kwargs):
        nonlocal refresh_called
        refresh_called = True
        return {"status": "refreshed"}

    monkeypatch.setattr(auth_msal, "_refresh_one", fake_refresh)

    import pytest

    with pytest.raises(RuntimeError, match="does not match requested account"):
        auth_msal.force_reauthenticate(
            TEST_EMAIL, tokens_dir=tmp_path, also_outlook=True
        )

    assert refresh_called is False
    assert list(tmp_path.iterdir()) == []


def test_force_reauthenticate_reports_partial_when_outlook_mint_fails(
    tmp_path, monkeypatch
):
    def fake_authenticate(self):
        self._save_tokens(
            access_token="a.b.c",
            refresh_token="fresh-shared-rt",
            expires_in=3600,
            scopes=auth_msal.GRAPH_SCOPE,
            email=TEST_EMAIL,
        )
        return {"access_token": "a.b.c", "username": TEST_EMAIL}

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth, "authenticate", fake_authenticate
    )
    monkeypatch.setattr(
        auth_msal,
        "_refresh_one",
        lambda *args, **kwargs: {
            "identifier": TEST_EMAIL,
            "status": "failed",
            "expires_at": None,
            "error": "outlook unavailable",
        },
    )

    result = auth_msal.force_reauthenticate(
        TEST_EMAIL, tokens_dir=tmp_path, also_outlook=True
    )

    assert result["status"] == "partial"
    assert result["error"] == "Outlook token mint failed: outlook unavailable"


def test_force_reauthenticate_without_outlook_has_no_outlook_key(tmp_path, monkeypatch):
    def fake_authenticate(self):
        self._save_tokens(
            access_token="a.b.c",
            refresh_token="fresh-shared-rt",
            expires_in=3600,
            scopes=auth_msal.GRAPH_SCOPE,
            email=TEST_EMAIL,
        )
        return {"access_token": "a.b.c", "username": TEST_EMAIL}

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth, "authenticate", fake_authenticate
    )
    result = auth_msal.force_reauthenticate(TEST_EMAIL, tokens_dir=tmp_path)
    assert "outlook" not in result


def test_force_reauthenticate_rejects_unverifiable_identity(tmp_path, monkeypatch):
    def fake_authenticate(self):
        self._save_tokens(
            access_token="not-a-jwt",
            refresh_token="unverified-refresh",
            expires_in=3600,
            scopes=auth_msal.GRAPH_SCOPE,
        )
        return {"access_token": "not-a-jwt"}

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth, "authenticate", fake_authenticate
    )

    import pytest

    with pytest.raises(RuntimeError, match="could not verify"):
        auth_msal.force_reauthenticate(TEST_EMAIL, tokens_dir=tmp_path)

    assert list(tmp_path.iterdir()) == []


def test_classify_refresh_error_recognizes_65002():
    hint = auth_msal.classify_refresh_error(
        "AADSTS65002: Consent between first party application ...",
        identifier=TEST_EMAIL,
    )
    assert hint is not None
    assert hint["code"] == "AADSTS65002"
    assert "preauthorized" in hint["summary"]
    assert "MICROSOFT_MCP_CLIENT_ID" in hint["remedy"]
    assert "app registration" in hint["remedy"]


def test_classify_refresh_error_expired_and_password_change():
    expired = auth_msal.classify_refresh_error("... AADSTS70008 expired ...")
    assert expired is not None and expired["code"] == "AADSTS70008"
    pw = auth_msal.classify_refresh_error("AADSTS50173 credential changed")
    assert pw is not None and pw["code"] == "AADSTS50173"


def test_classify_refresh_error_unknown_and_empty_return_none():
    assert auth_msal.classify_refresh_error("some unrelated error") is None
    assert auth_msal.classify_refresh_error("") is None
    assert auth_msal.classify_refresh_error(None) is None


def test_classify_refresh_error_falls_back_to_placeholder_email():
    hint = auth_msal.classify_refresh_error("AADSTS65002 ...")
    assert hint is not None
    assert "MICROSOFT_MCP_CLIENT_ID" in hint["remedy"]


def test_refresh_failure_attaches_65002_hint(tmp_path, monkeypatch):
    # A failed refresh whose error carries AADSTS65002 must surface a
    # structured `hint` on the result so the CLI / MCP tool can guide recovery.
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)

    def boom(self, refresh_token):
        raise RuntimeError(
            "Token refresh failed: AADSTS65002: Consent between first party ..."
        )

    monkeypatch.setattr(auth_msal.MSALRefreshTokenAuth, "_refresh_access_token", boom)
    result = auth_msal.refresh_account(TEST_EMAIL, tokens_dir=tmp_path)
    assert result["status"] == "failed"
    assert "AADSTS65002" in result["error"]
    assert result["hint"]["code"] == "AADSTS65002"
    assert TEST_EMAIL in result["hint"]["remedy"]


def test_refresh_failure_without_known_code_has_no_hint(tmp_path, monkeypatch):
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)

    def boom(self, refresh_token):
        raise RuntimeError("network unreachable")

    monkeypatch.setattr(auth_msal.MSALRefreshTokenAuth, "_refresh_access_token", boom)
    result = auth_msal.refresh_account(TEST_EMAIL, tokens_dir=tmp_path)
    assert result["status"] == "failed"
    assert "hint" not in result


def test_outlook_refresh_persists_latest_replacement_refresh_token(
    tmp_path, monkeypatch
):
    # Microsoft refresh tokens are user/client-bound, not resource-bound. Each
    # successful refresh returns the preferred replacement token to persist.
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)
    refresh_path = tmp_path / f"{TEST_EMAIL}_refresh_only.txt"
    assert refresh_path.read_text() == "shared-refresh"

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth,
        "_refresh_access_token",
        lambda self, rt: {
            "access_token": "n",
            "refresh_token": "outlook-rotated-token",
            "expires_in": 3600,
            "scope": auth_msal.OUTLOOK_SCOPE,
        },
    )
    result = auth_msal.refresh_account(
        TEST_EMAIL, tokens_dir=tmp_path, api_type="outlook"
    )
    assert result["status"] == "refreshed"
    assert refresh_path.read_text() == "outlook-rotated-token"


def test_graph_refresh_does_persist_rotated_refresh_token(tmp_path, monkeypatch):
    # Graph refreshes also persist Microsoft's latest replacement token.
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)
    refresh_path = tmp_path / f"{TEST_EMAIL}_refresh_only.txt"

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth,
        "_refresh_access_token",
        lambda self, rt: {
            "access_token": "n",
            "refresh_token": "graph-rotated-token",
            "expires_in": 3600,
            "scope": auth_msal.GRAPH_SCOPE,
        },
    )
    result = auth_msal.refresh_account(
        TEST_EMAIL, tokens_dir=tmp_path, api_type="graph"
    )
    assert result["status"] == "refreshed"
    assert refresh_path.read_text() == "graph-rotated-token"


def test_outlook_refresh_persists_outlook_scope_when_response_omits_scope(
    tmp_path, monkeypatch
):
    # Azure normally returns `scope`, but if it doesn't, an outlook instance
    # must still persist the OUTLOOK scope (not a graph fallback).
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)
    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth,
        "_refresh_access_token",
        lambda self, rt: {
            "access_token": "n",
            "refresh_token": "shared-refresh",
            "expires_in": 3600,
            # deliberately NO "scope" key
        },
    )
    result = auth_msal.refresh_account(
        TEST_EMAIL, tokens_dir=tmp_path, api_type="outlook"
    )
    assert result["status"] == "refreshed"
    import json as _json

    data = _json.loads(
        (tmp_path / f"{TEST_EMAIL}_outlook_access_token.json").read_text()
    )
    assert data["scopes"] == auth_msal.OUTLOOK_SCOPE
