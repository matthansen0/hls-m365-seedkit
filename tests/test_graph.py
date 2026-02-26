"""Tests for m365seed.graph — retry logic and client behaviour."""

import time
from unittest.mock import MagicMock, patch

import httpx
import pytest

from m365seed.graph import GraphClient, MAX_RETRIES, DEFAULT_RETRY_AFTER


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

MINIMAL_CFG = {
    "tenant": {"tenant_id": "00000000-0000-0000-0000-000000000000"},
    "auth": {
        "mode": "client_secret",
        "client_id": "11111111-1111-1111-1111-111111111111",
        "client_secret_env": "M365SEED_CLIENT_SECRET",
    },
    "content": {"run_id": "test-001", "theme": "healthcare"},
    "targets": {"users": [{"upn": "user@test.com"}]},
}


def _make_client(dry_run=False):
    """Build a GraphClient with mocked credential."""
    with patch("m365seed.graph.build_credential") as mock_cred:
        mock_token = MagicMock()
        mock_token.token = "fake-token"
        mock_cred.return_value = MagicMock(get_token=MagicMock(return_value=mock_token))
        client = GraphClient(MINIMAL_CFG, dry_run=dry_run)
    return client


# ---------------------------------------------------------------------------
# Dry-run tests
# ---------------------------------------------------------------------------


class TestDryRun:
    def test_dry_run_returns_200(self):
        client = _make_client(dry_run=True)
        resp = client.get("/users")
        assert resp.status_code == 200
        data = resp.json()
        assert data.get("@dry_run") is True

    def test_dry_run_post(self):
        client = _make_client(dry_run=True)
        resp = client.post("/users/a@b.com/sendMail", json_body={"test": True})
        assert resp.status_code == 200

    def test_dry_run_does_not_send_request(self):
        client = _make_client(dry_run=True)
        # _http.request should never be called
        client._http = MagicMock()
        client.get("/anything")
        client._http.request.assert_not_called()


# ---------------------------------------------------------------------------
# Retry logic tests (mocked HTTP)
# ---------------------------------------------------------------------------


class TestRetryLogic:
    def test_successful_request(self):
        client = _make_client()
        mock_resp = httpx.Response(
            200,
            json={"value": []},
            request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/users"),
        )
        client._http = MagicMock()
        client._http.request.return_value = mock_resp

        resp = client.get("/users")
        assert resp.status_code == 200
        assert client._http.request.call_count == 1

    @patch("m365seed.graph.time.sleep")
    def test_retry_on_429(self, mock_sleep):
        client = _make_client()
        resp_429 = httpx.Response(
            429,
            headers={"Retry-After": "2"},
            request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/users"),
        )
        resp_200 = httpx.Response(
            200,
            json={"ok": True},
            request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/users"),
        )
        client._http = MagicMock()
        client._http.request.side_effect = [resp_429, resp_200]

        resp = client.get("/users")
        assert resp.status_code == 200
        assert client._http.request.call_count == 2
        mock_sleep.assert_called_once_with(2)

    @patch("m365seed.graph.time.sleep")
    def test_retry_on_503(self, mock_sleep):
        client = _make_client()
        resp_503 = httpx.Response(
            503,
            request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/users"),
        )
        resp_200 = httpx.Response(
            200,
            json={"ok": True},
            request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/users"),
        )
        client._http = MagicMock()
        client._http.request.side_effect = [resp_503, resp_200]

        resp = client.get("/users")
        assert resp.status_code == 200
        assert client._http.request.call_count == 2

    @patch("m365seed.graph.time.sleep")
    def test_max_retries_exhausted(self, mock_sleep):
        client = _make_client()
        resp_429 = httpx.Response(
            429,
            headers={"Retry-After": "1"},
            request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/users"),
        )
        client._http = MagicMock()
        client._http.request.return_value = resp_429

        with pytest.raises(RuntimeError, match="Max retries exhausted"):
            client.get("/users")

        assert client._http.request.call_count == MAX_RETRIES

    @patch("m365seed.graph.time.sleep")
    def test_transport_error_retry(self, mock_sleep):
        client = _make_client()
        resp_200 = httpx.Response(
            200,
            json={"ok": True},
            request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/users"),
        )
        client._http = MagicMock()
        client._http.request.side_effect = [
            httpx.ConnectError("connection refused"),
            resp_200,
        ]

        resp = client.get("/users")
        assert resp.status_code == 200
        assert client._http.request.call_count == 2

    @patch("m365seed.graph.time.sleep")
    def test_transport_error_exhausted(self, mock_sleep):
        client = _make_client()
        client._http = MagicMock()
        client._http.request.side_effect = httpx.ConnectError("refused")

        with pytest.raises(httpx.ConnectError):
            client.get("/users")

        assert client._http.request.call_count == MAX_RETRIES

    def test_http_error_raised_for_400(self):
        client = _make_client()
        resp_400 = httpx.Response(
            400,
            json={"error": "bad request"},
            request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/users"),
        )
        client._http = MagicMock()
        client._http.request.return_value = resp_400

        with pytest.raises(httpx.HTTPStatusError):
            client.get("/users")


# ---------------------------------------------------------------------------
# Auth header tests
# ---------------------------------------------------------------------------


class TestAuthHeaders:
    def test_auth_header_present(self):
        client = _make_client()
        headers = client._auth_headers()
        assert headers["Authorization"] == "Bearer fake-token"
        assert headers["Content-Type"] == "application/json"
