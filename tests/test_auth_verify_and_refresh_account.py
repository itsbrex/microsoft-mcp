"""Tests for verify_account_tokens, refresh_account, and force_reauthenticate.

Test fixtures use a single email (``broach@cresa.com``) intentionally.
``refresh_all_accounts`` and related single-account flows must work under
the supported single-account configuration. The mismatch/drift tests still
write only one token file on disk; they vary the JWT claim inside that file
to simulate misconfiguration.
"""

import base64
import json
from datetime import datetime, timedelta, timezone
from unittest.mock import patch

import pytest

from microsoft_mcp.auth_msal import (
    MSALRefreshTokenAuth,
    force_reauthenticate,
    refresh_account,
    verify_account_tokens,
)

# Canonical single-account email used across all test fixtures in this file.
TEST_EMAIL = "broach@cresa.com"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode()


def _fake_jwt(claims: dict) -> str:
    """Build a JWT with arbitrary payload claims.

    The header and signature are not validated by our code path, so we use
    placeholders. Only the middle (payload) segment is decoded.
    """
    header = _b64url(json.dumps({"alg": "RS256", "typ": "JWT"}).encode())
    payload = _b64url(json.dumps(claims).encode())
    signature = _b64url(b"sig-placeholder")
    return f"{header}.{payload}.{signature}"


def _write_token_file(
    tmp_path,
    identifier: str,
    *,
    jwt_upn: str | None = None,
    jwt_oid: str = "00000000-0000-0000-0000-000000000000",
    jwt_tid: str = "11111111-1111-1111-1111-111111111111",
    expires_delta: timedelta = timedelta(hours=1),
    use_garbage_token: bool = False,
) -> dict:
    """Write a token file whose JWT payload claims the given identity."""
    expires_at = datetime.now(timezone.utc) + expires_delta
    upn = jwt_upn if jwt_upn is not None else identifier
    if use_garbage_token:
        access_token = "not-a-jwt"
    else:
        access_token = _fake_jwt(
            {
                "upn": upn,
                "oid": jwt_oid,
                "tid": jwt_tid,
                "aud": "https://graph.microsoft.com",
            }
        )
    data = {
        "email": identifier,
        "access_token": access_token,
        "token_type": "Bearer",
        "expires_in": max(int(expires_delta.total_seconds()), 1),
        "expires_at": expires_at.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "refreshed_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "scopes": "https://graph.microsoft.com/.default",
        "api_type": "graph",
    }
    (tmp_path / f"{identifier}_access_token.json").write_text(json.dumps(data))
    return data


def _write_refresh_token(tmp_path, identifier: str, value: str = "fake-rt") -> None:
    (tmp_path / f"{identifier}_refresh_only.txt").write_text(value)


# ---------------------------------------------------------------------------
# verify_account_tokens
# ---------------------------------------------------------------------------


class TestVerifyAccountTokens:
    def test_empty_dir_returns_empty(self, tmp_path):
        assert verify_account_tokens(tokens_dir=tmp_path) == []

    def test_nonexistent_dir_returns_empty(self, tmp_path):
        missing = tmp_path / "nope"
        assert verify_account_tokens(tokens_dir=missing) == []

    def test_jwt_match_reports_match_true(self, tmp_path):
        _write_token_file(tmp_path, TEST_EMAIL)
        result = verify_account_tokens(tokens_dir=tmp_path)
        assert len(result) == 1
        entry = result[0]
        assert entry["identifier"] == TEST_EMAIL
        assert entry["jwt_upn"] == TEST_EMAIL
        assert entry["match"] is True

    def test_jwt_mismatch_reports_match_false(self, tmp_path):
        """Filename labeled with TEST_EMAIL but JWT claims a different upn.

        Only one token file is on disk; the test simulates a wrong-account
        save by varying the JWT claim, not by adding a second file.
        """
        _write_token_file(tmp_path, TEST_EMAIL, jwt_upn="someone-else@cresa.com")
        result = verify_account_tokens(tokens_dir=tmp_path)
        assert len(result) == 1
        entry = result[0]
        assert entry["identifier"] == TEST_EMAIL
        assert entry["jwt_upn"] == "someone-else@cresa.com"
        assert entry["match"] is False

    def test_case_insensitive_match(self, tmp_path):
        """Filename casing differs from JWT casing; match must be case-insensitive."""
        upper = TEST_EMAIL.upper()  # "BROACH@CRESA.COM"
        _write_token_file(tmp_path, upper, jwt_upn=TEST_EMAIL)
        result = verify_account_tokens(tokens_dir=tmp_path)
        assert result[0]["match"] is True

    def test_garbage_token_reports_decode_error(self, tmp_path):
        _write_token_file(tmp_path, TEST_EMAIL, use_garbage_token=True)
        result = verify_account_tokens(tokens_dir=tmp_path)
        assert result[0]["match"] is False
        assert "jwt_decode_error" in result[0]

    def test_live_mode_calls_graph_me(self, tmp_path):
        """When live=True, hit Graph /me and incorporate its identifiers."""
        _write_token_file(tmp_path, TEST_EMAIL, jwt_upn=TEST_EMAIL)

        class FakeResponse:
            status_code = 200
            text = ""

            def json(self):
                return {
                    "userPrincipalName": TEST_EMAIL,
                    "mail": TEST_EMAIL,
                    "id": "graph-id-1",
                }

        with patch("httpx.get", return_value=FakeResponse()):
            result = verify_account_tokens(tokens_dir=tmp_path, live=True)

        entry = result[0]
        assert entry["graph_userPrincipalName"] == TEST_EMAIL
        assert entry["graph_id"] == "graph-id-1"
        assert entry["graph_error"] is None
        assert entry["match"] is True

    def test_live_mode_handles_graph_error(self, tmp_path):
        """Graph errors don't crash; they end up in graph_error."""
        _write_token_file(tmp_path, TEST_EMAIL)

        class FakeResponse:
            status_code = 429
            text = "Throttled"

            def json(self):
                return {}

        with patch("httpx.get", return_value=FakeResponse()):
            result = verify_account_tokens(tokens_dir=tmp_path, live=True)

        entry = result[0]
        assert entry["graph_userPrincipalName"] is None
        assert "HTTP 429" in entry["graph_error"]
        # JWT still matches, so overall match is True regardless of Graph
        assert entry["match"] is True


# ---------------------------------------------------------------------------
# refresh_account
# ---------------------------------------------------------------------------


class TestRefreshAccount:
    def test_empty_identifier_raises(self, tmp_path):
        with pytest.raises(ValueError, match="non-empty"):
            refresh_account(identifier="", tokens_dir=tmp_path)

    def test_missing_token_file_returns_missing(self, tmp_path):
        result = refresh_account(identifier=TEST_EMAIL, tokens_dir=tmp_path)
        assert result["status"] == "missing"
        assert result["identifier"] == TEST_EMAIL
        assert result["expires_at"] is None
        assert "no token file" in result["error"]

    def test_valid_token_reports_valid(self, tmp_path):
        data = _write_token_file(tmp_path, TEST_EMAIL)
        result = refresh_account(identifier=TEST_EMAIL, tokens_dir=tmp_path)
        assert result["status"] == "valid"
        assert result["expires_at"] == data["expires_at"]
        assert result["error"] is None

    def test_expired_token_triggers_refresh(self, tmp_path):
        _write_token_file(tmp_path, TEST_EMAIL, expires_delta=timedelta(seconds=-10))
        _write_refresh_token(tmp_path, TEST_EMAIL)

        mock_response = {
            "access_token": _fake_jwt({"upn": TEST_EMAIL}),
            "refresh_token": "new-rt",
            "expires_in": 3600,
            "scope": "https://graph.microsoft.com/.default",
        }
        with patch.object(
            MSALRefreshTokenAuth, "_refresh_access_token", return_value=mock_response
        ):
            result = refresh_account(identifier=TEST_EMAIL, tokens_dir=tmp_path)

        assert result["status"] == "refreshed"
        assert result["expires_at"] is not None
        assert result["error"] is None

    def test_refresh_failure_reports_failed(self, tmp_path):
        _write_token_file(tmp_path, TEST_EMAIL, expires_delta=timedelta(seconds=-10))
        _write_refresh_token(tmp_path, TEST_EMAIL)

        with patch.object(
            MSALRefreshTokenAuth,
            "_refresh_access_token",
            side_effect=RuntimeError("AADSTS70008"),
        ):
            result = refresh_account(identifier=TEST_EMAIL, tokens_dir=tmp_path)

        assert result["status"] == "failed"
        assert "AADSTS70008" in result["error"]
        # Failure must NOT evict the refresh token (per the contract).
        assert (tmp_path / f"{TEST_EMAIL}_refresh_only.txt").exists()


# ---------------------------------------------------------------------------
# force_reauthenticate
# ---------------------------------------------------------------------------


class TestForceReauthenticate:
    def test_empty_identifier_raises(self, tmp_path):
        with pytest.raises(ValueError, match="non-empty"):
            force_reauthenticate(identifier="", tokens_dir=tmp_path)

    def test_clears_then_authenticates(self, tmp_path):
        # Pre-existing tokens that must be wiped by clear_cache.
        _write_token_file(tmp_path, TEST_EMAIL)
        _write_refresh_token(tmp_path, TEST_EMAIL, "old-rt")
        (tmp_path / f"{TEST_EMAIL}_access_only.txt").write_text("old-at")

        new_token = _fake_jwt({"upn": TEST_EMAIL, "oid": "abc", "tid": "t1"})

        def fake_authenticate(self):
            # Simulate what real authenticate() does: write fresh token files.
            self._save_tokens(
                access_token=new_token,
                refresh_token="brand-new-rt",
                expires_in=3600,
                scopes="https://graph.microsoft.com/.default",
                email=TEST_EMAIL,
            )
            return {"access_token": new_token, "username": TEST_EMAIL}

        with patch.object(MSALRefreshTokenAuth, "authenticate", fake_authenticate):
            result = force_reauthenticate(identifier=TEST_EMAIL, tokens_dir=tmp_path)

        assert result["status"] == "reauthenticated"
        assert result["identifier"] == TEST_EMAIL
        assert result["signed_in_as"] == TEST_EMAIL
        assert result["expires_at"] is not None

        # New refresh token replaced the old one.
        assert (
            tmp_path / f"{TEST_EMAIL}_refresh_only.txt"
        ).read_text() == "brand-new-rt"

    def test_detects_drift_when_user_signs_into_wrong_account(self, tmp_path):
        """signed_in_as should reflect the JWT upn, exposing wrong-account drift.

        Only one file is on disk (labeled ``TEST_EMAIL``); the JWT inside it
        claims a different upn, simulating a user signing in as the wrong
        identity during the device-code flow.
        """
        wrong_upn = "someone-else@cresa.com"
        wrong_token = _fake_jwt({"upn": wrong_upn})

        def fake_authenticate(self):
            self._save_tokens(
                access_token=wrong_token,
                refresh_token="rt",
                expires_in=3600,
                scopes="https://graph.microsoft.com/.default",
                email=wrong_upn,
            )
            return {"access_token": wrong_token}

        with patch.object(MSALRefreshTokenAuth, "authenticate", fake_authenticate):
            result = force_reauthenticate(identifier=TEST_EMAIL, tokens_dir=tmp_path)

        assert result["signed_in_as"] == wrong_upn
        assert result["identifier"] == TEST_EMAIL


# ---------------------------------------------------------------------------
# MCP tool wrappers
# ---------------------------------------------------------------------------


class TestMcpToolWrappers:
    def test_verify_tool_refuses_under_azure(self):
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import verify_account_tokens as tool_fn

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "azure"
            with pytest.raises(ValueError, match="MSAL"):
                tool_fn.fn()
        finally:
            tools_mod.auth_method = original

    def test_refresh_account_tool_refuses_under_azure(self):
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import refresh_account as tool_fn

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "azure"
            with pytest.raises(ValueError, match="MSAL"):
                tool_fn.fn(email=TEST_EMAIL)
        finally:
            tools_mod.auth_method = original

    def test_force_reauthenticate_tool_refuses_under_azure(self):
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import force_reauthenticate_account as tool_fn

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "azure"
            with pytest.raises(ValueError, match="MSAL"):
                tool_fn.fn(email=TEST_EMAIL)
        finally:
            tools_mod.auth_method = original

    def test_refresh_account_tool_validates_empty_email(self):
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import refresh_account as tool_fn

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "msal"
            with pytest.raises(ValueError, match="non-empty"):
                tool_fn.fn(email="")
        finally:
            tools_mod.auth_method = original

    def test_verify_tool_delegates_to_library(self, tmp_path, monkeypatch):
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import verify_account_tokens as tool_fn

        monkeypatch.setenv("MICROSOFT_MCP_TOKENS_DIR", str(tmp_path))
        fake = [{"identifier": TEST_EMAIL, "match": True}]
        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "msal"
            with patch(
                "microsoft_mcp.auth_msal.verify_account_tokens", return_value=fake
            ) as mock_lib:
                returned = tool_fn.fn()
            assert returned == fake
            assert mock_lib.call_args.kwargs.get("live") is False
        finally:
            tools_mod.auth_method = original

    def test_refresh_account_tool_delegates_to_library(self, tmp_path, monkeypatch):
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import refresh_account as tool_fn

        monkeypatch.setenv("MICROSOFT_MCP_TOKENS_DIR", str(tmp_path))
        monkeypatch.setenv("MICROSOFT_MCP_CLIENT_ID", "id")
        fake = {
            "identifier": TEST_EMAIL,
            "status": "valid",
            "expires_at": None,
            "error": None,
            "api_type": "graph",
        }
        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "msal"
            with patch(
                "microsoft_mcp.auth_msal.refresh_account", return_value=fake
            ) as mock_lib:
                returned = tool_fn.fn(email=TEST_EMAIL)
            assert returned == fake
            assert mock_lib.call_args.kwargs.get("identifier") == TEST_EMAIL
            # Default api_type is graph.
            assert mock_lib.call_args.kwargs.get("api_type") == "graph"
        finally:
            tools_mod.auth_method = original

    def test_refresh_account_tool_both_loops_graph_then_outlook(
        self, tmp_path, monkeypatch
    ):
        """api_type='both' must expand into two library calls (graph, outlook)
        and return a list of two result dicts — mirroring the auth CLI."""
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import refresh_account as tool_fn

        monkeypatch.setenv("MICROSOFT_MCP_TOKENS_DIR", str(tmp_path))
        monkeypatch.setenv("MICROSOFT_MCP_CLIENT_ID", "id")

        def fake_lib(identifier, api_type="graph", **_):
            return {
                "identifier": identifier,
                "status": "refreshed",
                "expires_at": None,
                "error": None,
                "api_type": api_type,
            }

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "msal"
            with patch(
                "microsoft_mcp.auth_msal.refresh_account", side_effect=fake_lib
            ) as mock_lib:
                returned = tool_fn.fn(email=TEST_EMAIL, api_type="both")
            assert isinstance(returned, list)
            assert [r["api_type"] for r in returned] == ["graph", "outlook"]
            assert mock_lib.call_count == 2
            # Graph leg first, then outlook.
            assert [c.kwargs.get("api_type") for c in mock_lib.call_args_list] == [
                "graph",
                "outlook",
            ]
        finally:
            tools_mod.auth_method = original

    def test_refresh_account_tool_rejects_invalid_api_type(self):
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import refresh_account as tool_fn

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "msal"
            with pytest.raises(ValueError, match="api_type"):
                tool_fn.fn(email=TEST_EMAIL, api_type="nonsense")
        finally:
            tools_mod.auth_method = original

    def test_refresh_all_accounts_tool_passes_api_type(self, tmp_path, monkeypatch):
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import refresh_all_accounts as tool_fn

        monkeypatch.setenv("MICROSOFT_MCP_TOKENS_DIR", str(tmp_path))
        fake = [
            {"identifier": TEST_EMAIL, "status": "valid", "api_type": "graph"},
            {"identifier": TEST_EMAIL, "status": "valid", "api_type": "outlook"},
        ]
        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "msal"
            with patch(
                "microsoft_mcp.auth_msal.refresh_all_accounts", return_value=fake
            ) as mock_lib:
                returned = tool_fn.fn(api_type="both")
            assert returned == fake
            assert mock_lib.call_args.kwargs.get("api_type") == "both"
        finally:
            tools_mod.auth_method = original

    def test_refresh_all_accounts_tool_rejects_invalid_api_type(self):
        import microsoft_mcp.tools as tools_mod
        from microsoft_mcp.tools import refresh_all_accounts as tool_fn

        original = tools_mod.auth_method
        try:
            tools_mod.auth_method = "msal"
            with pytest.raises(ValueError, match="api_type"):
                tool_fn.fn(api_type="nonsense")
        finally:
            tools_mod.auth_method = original
