"""Tests for m365seed.graph — retry logic and client behavior."""

import time
from unittest.mock import MagicMock, patch

import httpx
import pytest

from m365seed.graph import (
    AZURE_CLI_PUBLIC_CLIENT_ID,
    DEFAULT_RETRY_AFTER,
    MAX_RETRIES,
    CacheSafeAzureCliCredential,
    DelegatedGraphCredential,
    GraphClient,
)


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


# ---------------------------------------------------------------------------
# Delegated credential tests
# ---------------------------------------------------------------------------


class TestDelegatedCredentials:
    def test_cache_safe_azure_cli_credential_purges_cache(self):
        token = MagicMock()
        inner = MagicMock()
        inner.get_token.return_value = token

        with patch("m365seed.graph.AzureCliCredential", return_value=inner), patch(
            "m365seed.register._ensure_msal_cache_healthy"
        ) as ensure_healthy:
            credential = CacheSafeAzureCliCredential("tenant-id")
            result = credential.get_token("scope-a")

        assert result is token
        ensure_healthy.assert_called_once()
        inner.get_token.assert_called_once_with("scope-a")

    def test_delegated_graph_credential_uses_device_code(self):
        device_token = MagicMock()
        cli_credential = MagicMock()
        device_credential = MagicMock()
        device_credential.get_token.return_value = device_token

        with patch(
            "m365seed.graph.CacheSafeAzureCliCredential",
            return_value=cli_credential,
        ), patch(
            "m365seed.graph.DeviceCodeCredential",
            return_value=device_credential,
        ) as device_ctor, patch.object(
            DelegatedGraphCredential, "_resolve_cache_path", return_value=None,
        ):
            credential = DelegatedGraphCredential("tenant-id", "client-id")
            result = credential.get_token("scope-a", "scope-b")

        assert result is device_token
        # Azure CLI should NOT be called — its tokens lack delegated scopes
        cli_credential.get_token.assert_not_called()
        device_credential.get_token.assert_called_once_with(
            "https://graph.microsoft.com/.default",
        )
        device_ctor.assert_called_once_with(
            tenant_id="tenant-id",
            client_id="client-id",
            cache_persistence_options=device_ctor.call_args.kwargs["cache_persistence_options"],
            prompt_callback=device_ctor.call_args.kwargs["prompt_callback"],
        )

    def test_delegated_graph_credential_device_code_error_propagates(self):
        cli_credential = MagicMock()
        device_credential = MagicMock()
        device_credential.get_token.side_effect = RuntimeError("device code failed")

        with patch(
            "m365seed.graph.CacheSafeAzureCliCredential",
            return_value=cli_credential,
        ), patch(
            "m365seed.graph.DeviceCodeCredential",
            return_value=device_credential,
        ), patch.object(
            DelegatedGraphCredential, "_resolve_cache_path", return_value=None,
        ):
            credential = DelegatedGraphCredential("tenant-id", "client-id")
            with pytest.raises(RuntimeError, match="device code failed"):
                credential.get_token("scope-a", "scope-b")

    def test_build_delegated_client_uses_public_client_fallback(self):
        cfg = {
            **MINIMAL_CFG,
            "auth": {
                **MINIMAL_CFG["auth"],
                "mode": "client_secret",
            },
        }
        delegated_credential = MagicMock()

        with patch("m365seed.graph.DelegatedGraphCredential", return_value=delegated_credential) as delegated_ctor, patch(
            "m365seed.graph.GraphClient"
        ) as graph_client_ctor:
            from m365seed.graph import build_delegated_client

            build_delegated_client(cfg, dry_run=True)

        delegated_ctor.assert_called_once_with(
            tenant_id=cfg["tenant"]["tenant_id"],
            device_code_client_id=cfg["auth"]["client_id"],
        )
        delegated_cfg = graph_client_ctor.call_args.args[0]
        assert delegated_cfg["auth"]["mode"] == "device_code"
        assert delegated_cfg["auth"]["client_id"] == cfg["auth"]["client_id"]
