"""Tests for m365seed.config — schema validation and helpers."""

import os
import pytest
import jsonschema

from m365seed.config import (
    validate_config,
    load_config,
    resolve_secret,
    get_run_id,
    get_theme,
    get_users,
)

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

MINIMAL_VALID_CONFIG = {
    "tenant": {"tenant_id": "00000000-0000-0000-0000-000000000000"},
    "auth": {
        "mode": "client_secret",
        "client_id": "11111111-1111-1111-1111-111111111111",
        "client_secret_env": "M365SEED_CLIENT_SECRET",
    },
    "targets": {
        "users": [
            {"upn": "user@example.onmicrosoft.com", "role": "Tester"},
        ],
    },
    "content": {
        "theme": "healthcare",
        "run_id": "test-run-001",
    },
}


def _cfg(**overrides):
    """Return a deep copy of the minimal config with overrides applied."""
    import copy

    cfg = copy.deepcopy(MINIMAL_VALID_CONFIG)
    for key, val in overrides.items():
        parts = key.split(".")
        d = cfg
        for p in parts[:-1]:
            d = d[p]
        d[parts[-1]] = val
    return cfg


# ---------------------------------------------------------------------------
# Schema validation tests
# ---------------------------------------------------------------------------


class TestValidateConfig:
    def test_minimal_valid_config(self):
        validate_config(MINIMAL_VALID_CONFIG)  # should not raise

    def test_missing_tenant(self):
        cfg = _cfg()
        del cfg["tenant"]
        with pytest.raises(jsonschema.ValidationError, match="tenant"):
            validate_config(cfg)

    def test_missing_auth(self):
        cfg = _cfg()
        del cfg["auth"]
        with pytest.raises(jsonschema.ValidationError, match="auth"):
            validate_config(cfg)

    def test_missing_targets(self):
        cfg = _cfg()
        del cfg["targets"]
        with pytest.raises(jsonschema.ValidationError, match="targets"):
            validate_config(cfg)

    def test_missing_content(self):
        cfg = _cfg()
        del cfg["content"]
        with pytest.raises(jsonschema.ValidationError, match="content"):
            validate_config(cfg)

    def test_empty_tenant_id_fails(self):
        cfg = _cfg(**{"tenant.tenant_id": ""})
        with pytest.raises(jsonschema.ValidationError):
            validate_config(cfg)

    def test_invalid_auth_mode(self):
        cfg = _cfg(**{"auth.mode": "password"})
        with pytest.raises(jsonschema.ValidationError):
            validate_config(cfg)

    def test_invalid_theme(self):
        cfg = _cfg(**{"content.theme": "dental"})
        with pytest.raises(jsonschema.ValidationError):
            validate_config(cfg)

    def test_valid_themes(self):
        for theme in ("healthcare", "pharma", "medtech", "payor"):
            cfg = _cfg(**{"content.theme": theme})
            validate_config(cfg)  # should not raise

    def test_empty_users_array_fails(self):
        cfg = _cfg()
        cfg["targets"]["users"] = []
        with pytest.raises(jsonschema.ValidationError):
            validate_config(cfg)

    def test_user_missing_upn_fails(self):
        cfg = _cfg()
        cfg["targets"]["users"] = [{"role": "Doctor"}]
        with pytest.raises(jsonschema.ValidationError):
            validate_config(cfg)

    def test_mail_threads_validation(self):
        cfg = _cfg()
        cfg["mail"] = {
            "threads": [
                {
                    "thread_id": "t-001",
                    "subject": "Test Thread",
                    "participants": ["a@test.com"],
                    "messages": 3,
                    "include_attachments": False,
                }
            ]
        }
        validate_config(cfg)  # should not raise

    def test_mail_thread_missing_subject_fails(self):
        cfg = _cfg()
        cfg["mail"] = {
            "threads": [
                {
                    "thread_id": "t-001",
                    "participants": ["a@test.com"],
                    "messages": 3,
                }
            ]
        }
        with pytest.raises(jsonschema.ValidationError):
            validate_config(cfg)

    def test_calendar_events_valid(self):
        cfg = _cfg()
        cfg["calendar"] = {
            "enabled": True,
            "events": [
                {
                    "event_id": "e-001",
                    "subject": "Test Meeting",
                    "organizer": "user@test.com",
                    "duration_minutes": 30,
                }
            ],
        }
        validate_config(cfg)  # should not raise


# ---------------------------------------------------------------------------
# Load config tests
# ---------------------------------------------------------------------------


class TestLoadConfig:
    def test_file_not_found(self):
        with pytest.raises(FileNotFoundError):
            load_config("/nonexistent/path.yaml")

    def test_loads_valid_yaml(self, tmp_path):
        import yaml

        cfg_path = tmp_path / "test-config.yaml"
        cfg_path.write_text(yaml.dump(MINIMAL_VALID_CONFIG))
        result = load_config(str(cfg_path))
        assert result["content"]["run_id"] == "test-run-001"


# ---------------------------------------------------------------------------
# Helper function tests
# ---------------------------------------------------------------------------


class TestHelpers:
    def test_get_run_id(self):
        assert get_run_id(MINIMAL_VALID_CONFIG) == "test-run-001"

    def test_get_theme(self):
        assert get_theme(MINIMAL_VALID_CONFIG) == "healthcare"

    def test_get_theme_default(self):
        cfg = _cfg()
        del cfg["content"]["theme"]
        # Without theme key, get_theme returns default "healthcare"
        assert get_theme(cfg) == "healthcare"

    def test_get_users(self):
        users = get_users(MINIMAL_VALID_CONFIG)
        assert len(users) == 1
        assert users[0]["upn"] == "user@example.onmicrosoft.com"

    def test_resolve_secret_success(self, monkeypatch):
        monkeypatch.setenv("M365SEED_CLIENT_SECRET", "my-secret-value")
        result = resolve_secret(MINIMAL_VALID_CONFIG)
        assert result == "my-secret-value"

    def test_resolve_secret_missing(self, monkeypatch):
        monkeypatch.delenv("M365SEED_CLIENT_SECRET", raising=False)
        with pytest.raises(RuntimeError, match="not set or empty"):
            resolve_secret(MINIMAL_VALID_CONFIG)

    def test_resolve_secret_custom_env(self, monkeypatch):
        cfg = _cfg(**{"auth.client_secret_env": "CUSTOM_SECRET_VAR"})
        monkeypatch.setenv("CUSTOM_SECRET_VAR", "custom-secret")
        assert resolve_secret(cfg) == "custom-secret"

    def test_resolve_secret_invalid_env_name(self):
        cfg = _cfg(**{"auth.client_secret_env": "not-an-env-var-value"})
        with pytest.raises(RuntimeError, match="must be an environment variable name"):
            resolve_secret(cfg)
