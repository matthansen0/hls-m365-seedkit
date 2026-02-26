"""Tests for m365seed.sharepoint — SharePoint site/page seeding."""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock

from m365seed.sharepoint import seed_sharepoint


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _dry_client():
    client = MagicMock()
    client.dry_run = True

    mock_resp = MagicMock()
    mock_resp.json.return_value = {"id": "dry-run-id", "@dry_run": True}
    mock_resp.status_code = 200
    client.get.return_value = mock_resp
    client.post.return_value = mock_resp
    client.put.return_value = mock_resp
    return client


def _base_cfg():
    return {
        "targets": {
            "users": [{"upn": "admin@test.com", "role": "Admin"}]
        },
        "sharepoint": {
            "enabled": True,
            "owner": "admin@test.com",
            "sites": [
                {
                    "display_name": "Clinical Ops Hub",
                    "description": "Central hub for clinical ops",
                    "pages": [
                        {
                            "title": "Welcome",
                            "content": "<h2>Welcome to the hub</h2>",
                        }
                    ],
                    "documents": [
                        {
                            "filename": "SOP_Guide.txt",
                            "folder": "SOPs",
                            "content": "Standard operating procedures guide.",
                        }
                    ],
                }
            ],
        },
    }


# ---------------------------------------------------------------------------
# Tests — disabled / empty
# ---------------------------------------------------------------------------


class TestSharePointDisabled:
    def test_disabled_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["sharepoint"]["enabled"] = False
        actions = seed_sharepoint(client, cfg, "healthcare", "run001")
        assert actions == []

    def test_no_sites_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["sharepoint"]["sites"] = []
        actions = seed_sharepoint(client, cfg, "healthcare", "run001")
        assert actions == []

    def test_missing_sharepoint_section(self):
        client = _dry_client()
        cfg = {"targets": {"users": [{"upn": "a@b.com"}]}}
        actions = seed_sharepoint(client, cfg, "healthcare", "run001")
        assert actions == []


# ---------------------------------------------------------------------------
# Tests — dry-run seeding
# ---------------------------------------------------------------------------


class TestSharePointDryRun:
    def test_creates_site(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_sharepoint(client, cfg, "healthcare", "run001")
        site_actions = [a for a in actions if a["action"] == "create_site"]
        assert len(site_actions) == 1
        assert site_actions[0]["site"] == "Clinical Ops Hub"

    def test_creates_pages(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_sharepoint(client, cfg, "healthcare", "run001")
        page_actions = [a for a in actions if a["action"] == "create_page"]
        assert len(page_actions) == 1
        assert page_actions[0]["page"] == "Welcome"

    def test_uploads_documents(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_sharepoint(client, cfg, "healthcare", "run001")
        upload_actions = [a for a in actions if a["action"] == "upload_document"]
        assert len(upload_actions) == 1
        assert "run001_SOP_Guide.txt" in upload_actions[0]["filename"]

    def test_multiple_sites(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["sharepoint"]["sites"].append(
            {"display_name": "Compliance Portal", "pages": [], "documents": []}
        )
        actions = seed_sharepoint(client, cfg, "healthcare", "run001")
        site_actions = [a for a in actions if a["action"] == "create_site"]
        assert len(site_actions) == 2
