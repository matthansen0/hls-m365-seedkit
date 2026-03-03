"""Tests for m365seed.profiles — user profile branding."""

from unittest.mock import MagicMock, patch, call

import httpx
import pytest

from m365seed.profiles import (
    _build_profile_map,
    resolve_profile,
    seed_profiles,
    PROFILE_FIELDS,
)


# ---------------------------------------------------------------------------
# Profile map tests
# ---------------------------------------------------------------------------


class TestBuildProfileMap:
    def test_healthcare_profiles(self):
        pmap = _build_profile_map("healthcare")
        assert "Clinical Ops Manager" in pmap
        assert "Care Manager — Dr. Donald Wilson" in pmap
        assert "Care Manager — Dr. Mary Gonzalez" in pmap
        assert "Nurse Manager" in pmap
        assert "Compliance Officer" in pmap
        assert pmap["Clinical Ops Manager"]["companyName"] == "Contoso Health System"

    def test_pharma_profiles(self):
        pmap = _build_profile_map("pharma")
        assert "Research Director" in pmap
        assert pmap["Research Director"]["companyName"] == "Contoso Pharmaceuticals"
        assert "Research" in pmap["Research Director"]["department"]

    def test_medtech_profiles(self):
        pmap = _build_profile_map("medtech")
        assert "Product Manager" in pmap
        assert pmap["Product Manager"]["companyName"] == "Contoso Medical Devices"
        assert "Manufacturing Engineer" in pmap

    def test_payor_profiles(self):
        pmap = _build_profile_map("payor")
        assert "Claims Operations Manager" in pmap
        assert pmap["Claims Operations Manager"]["companyName"] == "Contoso Health Plans"
        assert "Actuarial Analyst" in pmap

    def test_all_themes_have_five_profiles(self):
        for theme in ("healthcare", "pharma", "medtech", "payor"):
            pmap = _build_profile_map(theme)
            assert len(pmap) == 5, f"Theme '{theme}' should have 5 profiles"

    def test_all_profiles_have_required_fields(self):
        for theme in ("healthcare", "pharma", "medtech", "payor"):
            pmap = _build_profile_map(theme)
            for role, profile in pmap.items():
                for field in PROFILE_FIELDS:
                    assert field in profile, (
                        f"Profile '{role}' in '{theme}' missing '{field}'"
                    )


# ---------------------------------------------------------------------------
# Profile resolution tests
# ---------------------------------------------------------------------------


class TestResolveProfile:
    def test_matching_role_returns_payload(self):
        pmap = _build_profile_map("healthcare")
        user = {"upn": "user@test.com", "role": "Care Manager — Dr. Donald Wilson"}
        payload = resolve_profile(user, pmap)
        assert payload is not None
        assert payload["jobTitle"] == "Care Manager"
        assert payload["department"] == "Care Management"
        assert payload["companyName"] == "Contoso Health System"
        assert "officeLocation" in payload

    def test_unknown_role_returns_none(self):
        pmap = _build_profile_map("healthcare")
        user = {"upn": "user@test.com", "role": "Unknown Role"}
        payload = resolve_profile(user, pmap)
        assert payload is None

    def test_missing_role_returns_none(self):
        pmap = _build_profile_map("healthcare")
        user = {"upn": "user@test.com"}
        payload = resolve_profile(user, pmap)
        assert payload is None

    def test_payload_has_only_graph_fields(self):
        """Ensure aboutMe and role are NOT in the PATCH payload."""
        pmap = _build_profile_map("healthcare")
        user = {"upn": "user@test.com", "role": "Compliance Officer"}
        payload = resolve_profile(user, pmap)
        assert payload is not None
        assert "aboutMe" not in payload
        assert "role" not in payload
        allowed = set(PROFILE_FIELDS)
        assert set(payload.keys()).issubset(allowed)


# ---------------------------------------------------------------------------
# Seed profiles integration tests (dry run)
# ---------------------------------------------------------------------------


class TestSeedProfilesDryRun:
    def _make_dry_client(self):
        with patch("m365seed.graph.build_credential") as mock_cred:
            mock_token = MagicMock()
            mock_token.token = "fake"
            mock_cred.return_value = MagicMock(
                get_token=MagicMock(return_value=mock_token)
            )
            from m365seed.graph import GraphClient

            client = GraphClient(
                {
                    "tenant": {"tenant_id": "t"},
                    "auth": {
                        "mode": "client_secret",
                        "client_id": "c",
                        "client_secret_env": "X",
                    },
                    "content": {"run_id": "r", "theme": "healthcare"},
                    "targets": {"users": [{"upn": "u@t.com"}]},
                },
                dry_run=True,
            )
        return client

    def test_seed_profiles_updates_all_users(self):
        client = self._make_dry_client()
        cfg = {
            "targets": {
                "users": [
                    {"upn": "allan@test.com", "role": "Clinical Ops Manager"},
                    {"upn": "megan@test.com", "role": "Care Manager — Dr. Donald Wilson"},
                    {"upn": "nestor@test.com", "role": "Care Manager — Dr. Mary Gonzalez"},
                ]
            }
        }
        actions = seed_profiles(client, cfg, "healthcare", "run-001")
        assert len(actions) == 3
        assert all(a["action"] == "update-profile" for a in actions)
        assert actions[0]["jobTitle"] == "Clinical Operations Manager"
        assert actions[1]["jobTitle"] == "Care Manager"
        assert actions[2]["department"] == "Care Management"

    def test_seed_profiles_skips_unknown_role(self):
        client = self._make_dry_client()
        cfg = {
            "targets": {
                "users": [
                    {"upn": "allan@test.com", "role": "Clinical Ops Manager"},
                    {"upn": "unknown@test.com", "role": "Marketing Manager"},
                ]
            }
        }
        actions = seed_profiles(client, cfg, "healthcare", "run-001")
        assert len(actions) == 2
        assert actions[0]["action"] == "update-profile"
        assert actions[1]["action"] == "skip-profile"
        assert "Marketing Manager" in actions[1]["reason"]

    def test_seed_profiles_empty_users(self):
        client = self._make_dry_client()
        cfg = {"targets": {"users": []}}
        actions = seed_profiles(client, cfg, "healthcare", "run-001")
        assert actions == []

    def test_seed_profiles_pharma_theme(self):
        client = self._make_dry_client()
        cfg = {
            "targets": {
                "users": [
                    {"upn": "user@test.com", "role": "Research Director"},
                ]
            }
        }
        actions = seed_profiles(client, cfg, "pharma", "run-001")
        assert len(actions) == 1
        assert actions[0]["companyName"] == "Contoso Pharmaceuticals"
        assert "Research" in actions[0]["department"]

    def test_seed_profiles_calls_graph_patch(self):
        """Ensure the seeder calls client.patch with correct arguments."""
        client = MagicMock(spec=["patch"])
        client.patch.return_value = MagicMock(status_code=200)

        cfg = {
            "targets": {
                "users": [
                    {"upn": "megan@test.com", "role": "Nurse Manager"},
                ]
            }
        }
        actions = seed_profiles(client, cfg, "healthcare", "run-001")

        assert client.patch.call_count == 1
        call_args = client.patch.call_args
        assert call_args[0][0] == "/users/megan@test.com"
        payload = call_args[1]["json_body"]
        assert payload["jobTitle"] == "Nurse Manager — Med/Surg Unit"
        assert payload["department"] == "Nursing"
        assert payload["companyName"] == "Contoso Health System"

    def test_seed_profiles_all_themes(self):
        """All 4 themes should produce valid actions for their defined roles."""
        themes_and_roles = {
            "healthcare": "Clinical Ops Manager",
            "pharma": "Research Director",
            "medtech": "Product Manager",
            "payor": "Claims Operations Manager",
        }
        for theme, role in themes_and_roles.items():
            client = MagicMock(spec=["patch"])
            client.patch.return_value = MagicMock(status_code=200)
            cfg = {
                "targets": {
                    "users": [{"upn": "user@test.com", "role": role}]
                }
            }
            actions = seed_profiles(client, cfg, theme, "run-001")
            assert len(actions) == 1, f"Theme '{theme}' should produce 1 action"
            assert actions[0]["action"] == "update-profile"
            assert actions[0]["companyName"] != ""

    def test_seed_profiles_handles_404_gracefully(self):
        """A 404 should log an error action but not crash the pipeline."""
        request = httpx.Request("PATCH", "https://graph.microsoft.com/v1.0/users/gone@test.com")
        response = httpx.Response(404, request=request)
        error = httpx.HTTPStatusError("Not Found", request=request, response=response)

        client = MagicMock(spec=["patch"])
        client.patch.side_effect = error

        cfg = {
            "targets": {
                "users": [
                    {"upn": "gone@test.com", "role": "Clinical Ops Manager"},
                    {"upn": "also_gone@test.com", "role": "Care Manager — Dr. Donald Wilson"},
                ]
            }
        }
        actions = seed_profiles(client, cfg, "healthcare", "run-001")
        assert len(actions) == 2
        assert all(a["action"] == "error-profile" for a in actions)
        assert "Not Found" in actions[0]["error"]
