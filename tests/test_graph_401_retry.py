"""Tests for 401 auto-refresh and replay in graph.request / graph.download_raw."""

import json
import pytest
import httpx
from pathlib import Path
from unittest.mock import MagicMock, patch

import src.microsoft_mcp.graph as graph_module
from src.microsoft_mcp.graph import request, download_raw, set_auth_instance
from src.microsoft_mcp.auth_msal import MSALRefreshTokenAuth


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _make_response(status_code: int, body: bytes = b'{"ok": true}') -> MagicMock:
    """Build a fake httpx.Response MagicMock."""
    resp = MagicMock(spec=httpx.Response)
    resp.status_code = status_code
    resp.headers = {}
    resp.content = body
    resp.json = lambda: json.loads(body)

    if status_code >= 400:
        resp.raise_for_status.side_effect = httpx.HTTPStatusError(
            message=f"HTTP {status_code}",
            request=MagicMock(),
            response=resp,
        )
    else:
        resp.raise_for_status.return_value = None

    return resp


# ---------------------------------------------------------------------------
# C.1 — MSALRefreshTokenAuth.force_refresh() unit tests
# ---------------------------------------------------------------------------


class TestForceRefresh:
    def _make_auth(self, tmp_path: Path) -> MSALRefreshTokenAuth:
        """Construct an MSALRefreshTokenAuth pointing at tmp_path."""
        return MSALRefreshTokenAuth(
            tokens_dir=tmp_path,
            client_id="d3590ed6-52b3-4102-aeff-aad2292ab01c",
            account_identifier="test@example.com",
        )

    def test_force_refresh_clears_and_replays(self, tmp_path: Path) -> None:
        """force_refresh() acquires token via _refresh_access_token and persists."""
        auth = self._make_auth(tmp_path)

        # Write a fake refresh token to disk
        refresh_token_path = tmp_path / "test@example.com_refresh_only.txt"
        refresh_token_path.write_text("old-refresh-token")

        new_token_dict = {
            "access_token": "new-access-token",
            "refresh_token": "new-refresh-token",
            "expires_in": 3600,
            "scope": "https://graph.microsoft.com/.default",
        }

        with patch.object(
            auth, "_refresh_access_token", return_value=new_token_dict
        ) as mock_refresh:
            auth.force_refresh()

        mock_refresh.assert_called_once_with("old-refresh-token")

        # The new access token should be persisted
        access_token_path = tmp_path / "test@example.com_access_token.json"
        assert access_token_path.exists()
        saved = json.loads(access_token_path.read_text())
        assert saved["access_token"] == "new-access-token"

    def test_force_refresh_raises_when_no_refresh_token(self, tmp_path: Path) -> None:
        """force_refresh() raises RuntimeError when no refresh token file exists."""
        auth = self._make_auth(tmp_path)
        # No refresh token file present

        with pytest.raises(RuntimeError, match="refresh token"):
            auth.force_refresh()


# ---------------------------------------------------------------------------
# C.2 — graph.request() 401 retry tests
# ---------------------------------------------------------------------------


class TestRequest401Retry:
    def setup_method(self) -> None:
        """Install a fresh mock auth before each test."""
        self.mock_auth = MagicMock()
        self.mock_auth.get_token.side_effect = ["token-old", "token-new"]
        self.mock_auth.force_refresh.return_value = None
        set_auth_instance(self.mock_auth)

    def teardown_method(self) -> None:
        graph_module._global_auth = None

    @patch("src.microsoft_mcp.graph._client")
    def test_request_401_then_200_replays_with_new_token(
        self, mock_client: MagicMock
    ) -> None:
        """On 401, force_refresh is called and the request is replayed successfully."""
        r401 = _make_response(401)
        r200 = _make_response(200, b'{"ok": true}')
        mock_client.request.side_effect = [r401, r200]

        result = request("GET", "/me")

        assert mock_client.request.call_count == 2
        self.mock_auth.force_refresh.assert_called_once()

        # Second call should carry the refreshed token
        second_call_headers = mock_client.request.call_args_list[1][1]["headers"]
        assert second_call_headers["Authorization"] == "Bearer token-new"

        assert result == {"ok": True}

    @patch("src.microsoft_mcp.graph._client")
    def test_request_401_then_401_raises_httpstatuserror(
        self, mock_client: MagicMock
    ) -> None:
        """Two consecutive 401s surface HTTPStatusError; force_refresh called once."""
        r401a = _make_response(401)
        r401b = _make_response(401)
        mock_client.request.side_effect = [r401a, r401b]

        with pytest.raises(httpx.HTTPStatusError) as exc_info:
            request("GET", "/me")

        assert exc_info.value.response.status_code == 401
        # force_refresh called exactly once — the sentinel prevents a second attempt
        self.mock_auth.force_refresh.assert_called_once()
        assert mock_client.request.call_count == 2

    @patch("src.microsoft_mcp.graph._client")
    def test_request_401_with_azure_auth_does_not_attempt_refresh(
        self, mock_client: MagicMock
    ) -> None:
        """Azure auth (no force_refresh attribute) gets immediate HTTPStatusError on 401."""
        azure_auth = MagicMock(spec=["get_token"])
        azure_auth.get_token.return_value = "azure-token"
        set_auth_instance(azure_auth)

        r401 = _make_response(401)
        mock_client.request.return_value = r401

        with pytest.raises(httpx.HTTPStatusError) as exc_info:
            request("GET", "/me")

        assert exc_info.value.response.status_code == 401
        # Only one request — no replay
        assert mock_client.request.call_count == 1

    @patch("src.microsoft_mcp.graph._client")
    def test_request_401_with_failing_force_refresh_surfaces_original_401(
        self, mock_client: MagicMock, caplog: pytest.LogCaptureFixture
    ) -> None:
        """If force_refresh() raises, a warning is logged and the original 401 surfaces."""
        self.mock_auth.force_refresh.side_effect = RuntimeError("no refresh token")

        r401 = _make_response(401)
        mock_client.request.return_value = r401

        import logging

        with caplog.at_level(logging.WARNING, logger="microsoft_mcp.graph"):
            with pytest.raises(httpx.HTTPStatusError) as exc_info:
                request("GET", "/me")

        assert exc_info.value.response.status_code == 401
        assert any(
            "Force-refresh after 401 failed" in r.message for r in caplog.records
        )


# ---------------------------------------------------------------------------
# C.3 — graph.download_raw() 401 retry test
# ---------------------------------------------------------------------------


class TestDownloadRaw401Retry:
    def setup_method(self) -> None:
        self.mock_auth = MagicMock()
        self.mock_auth.get_token.side_effect = ["token-old", "token-new"]
        self.mock_auth.force_refresh.return_value = None
        set_auth_instance(self.mock_auth)

    def teardown_method(self) -> None:
        graph_module._global_auth = None

    @patch("src.microsoft_mcp.graph._client")
    def test_download_raw_401_retry(self, mock_client: MagicMock) -> None:
        """download_raw retries once after 401, returns bytes on second success."""
        r401 = _make_response(401)
        r200 = _make_response(200, b"file-bytes")
        r200.content = b"file-bytes"
        mock_client.get.side_effect = [r401, r200]

        result = download_raw("/drives/item/content")

        assert mock_client.get.call_count == 2
        self.mock_auth.force_refresh.assert_called_once()

        # Second call carries the refreshed token
        second_call_headers = mock_client.get.call_args_list[1][1]["headers"]
        assert second_call_headers["Authorization"] == "Bearer token-new"

        assert result == b"file-bytes"
