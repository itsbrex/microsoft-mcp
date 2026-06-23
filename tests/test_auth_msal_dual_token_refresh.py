import json
import datetime as dt
from microsoft_mcp import auth_msal

TEST_EMAIL = "broach@cresa.com"


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
    api_types = sorted(r.get("api_type") for r in results)
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


def test_classify_refresh_error_recognizes_65002():
    hint = auth_msal.classify_refresh_error(
        "AADSTS65002: Consent between first party application ...",
        identifier=TEST_EMAIL,
    )
    assert hint is not None
    assert hint["code"] == "AADSTS65002"
    assert "Outlook grant" in hint["summary"]
    # Remedy interpolates the identifier and points at the force/both recovery.
    assert TEST_EMAIL in hint["remedy"]
    assert "--force --api both" in hint["remedy"]


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
    assert "<email>" in hint["remedy"]


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


def test_outlook_refresh_does_not_clobber_shared_refresh_token(tmp_path, monkeypatch):
    # The shared {id}_refresh_only.txt must stay Graph-consented. An Outlook
    # refresh response carries a rotated refresh token scoped to the Outlook
    # grant; persisting it would make the next Graph `.default` refresh fail
    # with AADSTS65002. The outlook leg must leave the shared token untouched.
    _seed_graph_account(tmp_path, TEST_EMAIL, valid=False)
    refresh_path = tmp_path / f"{TEST_EMAIL}_refresh_only.txt"
    assert refresh_path.read_text() == "shared-refresh"

    monkeypatch.setattr(
        auth_msal.MSALRefreshTokenAuth,
        "_refresh_access_token",
        lambda self, rt: {
            "access_token": "n",
            "refresh_token": "outlook-rotated-token",  # MUST NOT be persisted
            "expires_in": 3600,
            "scope": auth_msal.OUTLOOK_SCOPE,
        },
    )
    result = auth_msal.refresh_account(
        TEST_EMAIL, tokens_dir=tmp_path, api_type="outlook"
    )
    assert result["status"] == "refreshed"
    # Shared refresh token unchanged — still the Graph-consented one.
    assert refresh_path.read_text() == "shared-refresh"


def test_graph_refresh_does_persist_rotated_refresh_token(tmp_path, monkeypatch):
    # The Graph leg is the canonical writer: a rotated Graph refresh token
    # SHOULD overwrite the shared {id}_refresh_only.txt.
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
