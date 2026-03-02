"""Tests for m365seed.files — file manifest and rendering."""

from unittest.mock import MagicMock, patch

import httpx
import pytest

from m365seed.files import (
    _render_file,
    seed_files,
)
from m365seed.theme_content import get_file_manifest


# ---------------------------------------------------------------------------
# Manifest tests
# ---------------------------------------------------------------------------


class TestFileManifest:
    def test_manifest_not_empty(self):
        manifest = get_file_manifest("healthcare")
        assert len(manifest) > 0

    def test_manifest_structure(self):
        manifest = get_file_manifest("healthcare")
        for folder, filename, template_name, desc in manifest:
            assert isinstance(folder, str) and folder
            assert isinstance(filename, str) and filename
            assert isinstance(template_name, str) and template_name.endswith(".j2")
            assert isinstance(desc, str) and desc

    def test_manifest_folders(self):
        manifest = get_file_manifest("healthcare")
        folders = {entry[0] for entry in manifest}
        # At a minimum, these folders should exist
        assert "Clinical Ops" in folders
        assert "Compliance" in folders

    def test_manifest_per_theme(self):
        """Each theme should return its own file manifest."""
        for theme in ("healthcare", "pharma", "medtech", "payor"):
            manifest = get_file_manifest(theme)
            assert len(manifest) > 0
            # Every entry should reference a valid template name
            for _, _, template_name, _ in manifest:
                assert template_name.endswith(".j2")


# ---------------------------------------------------------------------------
# Template rendering tests
# ---------------------------------------------------------------------------


class TestRenderFile:
    def test_render_sop(self):
        content = _render_file("healthcare", "sop.txt.j2", "test-run")
        assert "test-run" in content
        assert "synthetic" in content.lower()

    def test_render_compliance_checklist(self):
        content = _render_file("healthcare", "compliance_checklist.txt.j2", "test-run")
        assert "test-run" in content

    def test_render_unknown_template_fallback(self):
        content = _render_file("healthcare", "nonexistent.txt.j2", "test-run")
        assert "test-run" in content
        assert "synthetic" in content.lower()


# ---------------------------------------------------------------------------
# seed_files dry-run test
# ---------------------------------------------------------------------------


class TestSeedFilesDryRun:
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
                    "auth": {"mode": "client_secret", "client_id": "c", "client_secret_env": "X"},
                    "content": {"run_id": "r", "theme": "healthcare"},
                    "targets": {"users": [{"upn": "u@t.com"}]},
                },
                dry_run=True,
            )
        return client

    def test_seed_files_disabled(self):
        client = self._make_dry_client()
        cfg = {
            "files": {"oneDrive": {"enabled": False}},
            "targets": {"users": [{"upn": "u@t.com"}]},
        }
        actions = seed_files(client, cfg, "healthcare", "run-001")
        assert actions == []

    def test_seed_files_dry_run(self):
        client = self._make_dry_client()
        cfg = {
            "files": {
                "oneDrive": {
                    "enabled": True,
                    "target_user": "user@test.com",
                    "folders": ["Clinical Ops", "Compliance"],
                }
            },
            "targets": {"users": [{"upn": "user@test.com"}]},
        }
        actions = seed_files(client, cfg, "healthcare", "run-dry")
        # Should have actions for files in Clinical Ops and Compliance
        assert len(actions) > 0
        assert all(a["action"] == "upload" for a in actions)
        # All paths should start with one of the configured folders
        for a in actions:
            assert a["path"].startswith("Clinical Ops/") or a["path"].startswith(
                "Compliance/"
            )

    def test_seed_files_handles_upload_error_gracefully(self):
        """A 404 from OneDrive upload should log an error action, not crash."""
        request = httpx.Request("PUT", "https://graph.microsoft.com/v1.0/users/gone@test.com/drive/root:/test.txt:/content")
        response = httpx.Response(404, request=request)
        error = httpx.HTTPStatusError("Not Found", request=request, response=response)

        client = MagicMock(spec=["put", "post", "get", "dry_run"])
        client.dry_run = True
        client.put.side_effect = error
        client.post.return_value = MagicMock(status_code=201, json=lambda: {})
        client.get.return_value = MagicMock(status_code=200, json=lambda: {})

        cfg = {
            "files": {
                "oneDrive": {
                    "enabled": True,
                    "target_user": "gone@test.com",
                    "folders": ["Clinical Ops"],
                }
            },
            "targets": {"users": [{"upn": "gone@test.com"}]},
        }
        actions = seed_files(client, cfg, "healthcare", "run-err")
        error_actions = [a for a in actions if a["action"] == "error"]
        assert len(error_actions) > 0
        assert all("Not Found" in a["error"] for a in error_actions)
