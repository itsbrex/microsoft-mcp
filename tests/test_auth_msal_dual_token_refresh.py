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
