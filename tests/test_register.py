"""Tests for m365seed.register — Azure CLI registration helpers."""

from pathlib import Path
import subprocess
from unittest.mock import patch

from m365seed.register import (
    _clear_msal_http_cache,
    _is_msal_http_cache_error,
    _get_azure_config_dir,
    register_app,
)


def _cp(returncode: int, stdout: str = "", stderr: str = "") -> subprocess.CompletedProcess:
    """Build a CompletedProcess for test stubs."""
    return subprocess.CompletedProcess(args=["az"], returncode=returncode, stdout=stdout, stderr=stderr)


def test_is_msal_http_cache_error_detects_signature() -> None:
    output = (
        "Can't get attribute 'NormalizedResponse' on <module 'msal.throttled_http_client' "
        "from '/usr/lib/python3/dist-packages/msal/throttled_http_client.py'>"
    )
    assert _is_msal_http_cache_error(output) is True


def test_clear_msal_http_cache_removes_files(tmp_path: Path) -> None:
    azure_dir = tmp_path / ".azure"
    azure_dir.mkdir()
    (azure_dir / "msal_http_cache.bin").write_text("x", encoding="utf-8")
    (azure_dir / "msal_http_cache.bin.lockfile").write_text("x", encoding="utf-8")

    with patch("m365seed.register.Path.home", return_value=tmp_path):
        assert _clear_msal_http_cache() is True

    assert not (azure_dir / "msal_http_cache.bin").exists()
    assert not (azure_dir / "msal_http_cache.bin.lockfile").exists()


def test_get_azure_config_dir_honors_env(tmp_path: Path) -> None:
    with patch.dict("os.environ", {"AZURE_CONFIG_DIR": str(tmp_path)}):
        assert _get_azure_config_dir() == tmp_path


def test_register_retries_login_after_cache_error() -> None:
    cache_error = "Can't get attribute 'NormalizedResponse' ... msal.throttled_http_client"

    with patch("m365seed.register._check_az_cli", return_value=True), patch(
        "m365seed.register._is_logged_in", return_value=False
    ), patch("m365seed.register._clear_msal_http_cache", return_value=True) as clear_cache, patch(
        "m365seed.register._az_json", return_value=None
    ), patch(
        "m365seed.register._az",
        side_effect=[
            _cp(1),
            _cp(1, stderr=cache_error),
            _cp(1, stderr=cache_error),
            _cp(0),
        ],
    ) as az_cmd:
        result = register_app("2c627739-3b65-451a-ac0d-d3ecea353a55")

    assert result is None
    assert az_cmd.call_count == 4
    clear_cache.assert_called_once()


def test_register_clears_cache_before_login_on_probe_error() -> None:
    cache_error = "Can't get attribute 'NormalizedResponse' ... msal.throttled_http_client"

    with patch("m365seed.register._check_az_cli", return_value=True), patch(
        "m365seed.register._is_logged_in", return_value=False
    ), patch(
        "m365seed.register._clear_msal_http_cache", return_value=True
    ) as clear_cache, patch("m365seed.register._az_json") as az_json, patch(
        "m365seed.register._az",
        side_effect=[
            _cp(1, stderr=cache_error),
            _cp(0),
            _cp(0),
            _cp(0),
        ],
    ):
        az_json.side_effect = [
            {"appId": "a", "id": "b"},
            {"id": "sp"},
            {"password": "s"},
        ]
        result = register_app("2c627739-3b65-451a-ac0d-d3ecea353a55")

    assert result is not None
    clear_cache.assert_called_once()
