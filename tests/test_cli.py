"""Tests for CLI validation behavior."""

from unittest.mock import MagicMock, patch

import httpx
from typer.testing import CliRunner

from m365seed.cli import app

runner = CliRunner()


def _cfg() -> dict:
    return {
        "tenant": {"tenant_id": "00000000-0000-0000-0000-000000000000"},
        "auth": {
            "mode": "client_secret",
            "client_id": "11111111-1111-1111-1111-111111111111",
            "client_secret_env": "M365SEED_CLIENT_SECRET",
        },
        "content": {"run_id": "test-001", "theme": "healthcare"},
        "targets": {"users": [{"upn": "demo@contoso.com"}]},
    }


def _http_error(status_code: int, detail: str, url: str) -> httpx.HTTPStatusError:
    request = httpx.Request("GET", url)
    response = httpx.Response(
        status_code,
        json={"error": {"message": detail}},
        request=request,
    )
    return httpx.HTTPStatusError(detail, request=request, response=response)


def test_validate_allows_client_secret_org_lookup_403() -> None:
    client = MagicMock()
    client.check_auth.side_effect = _http_error(
        403,
        "Insufficient privileges to complete the operation.",
        "https://graph.microsoft.com/v1.0/organization",
    )
    client.ensure_token.return_value = "token"
    client.get.return_value = httpx.Response(
        200,
        json={"id": "user-id"},
        request=httpx.Request("GET", "https://graph.microsoft.com/v1.0/users/demo@contoso.com"),
    )

    with patch("m365seed.cli.load_config", return_value=_cfg()), patch(
        "m365seed.cli._build_client", return_value=client
    ), patch("m365seed.cli._setup_logging"), patch("m365seed.cli._print_log_path"):
        result = runner.invoke(app, ["validate", "-c", "seed-config.yaml"])

    assert result.exit_code == 0
    assert "Graph authentication succeeded" in result.stdout
    assert "Access token acquired, but Microsoft Graph rejected GET /organization" in result.stdout
    client.ensure_token.assert_called_once()


def test_validate_warns_when_user_lookup_is_forbidden() -> None:
    client = MagicMock()
    client.check_auth.return_value = {"value": [{"displayName": "Contoso"}]}
    client.get.side_effect = _http_error(
        403,
        "Insufficient privileges to complete the operation.",
        "https://graph.microsoft.com/v1.0/users/demo@contoso.com",
    )

    with patch("m365seed.cli.load_config", return_value=_cfg()), patch(
        "m365seed.cli._build_client", return_value=client
    ), patch("m365seed.cli._setup_logging"), patch("m365seed.cli._print_log_path"):
        result = runner.invoke(app, ["validate", "-c", "seed-config.yaml"])

    assert result.exit_code == 0
    assert "Tenant: Contoso" in result.stdout
    assert "User demo@contoso.com — could not verify" in result.stdout
    assert "Graph denied the directory lookup" in result.stdout