"""Tests for m365seed.cleanup — expanded cleanup for all content types."""

from __future__ import annotations

from unittest.mock import MagicMock, call

from m365seed.cleanup import cleanup


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _dry_client():
    client = MagicMock()
    client.dry_run = True

    mock_resp = MagicMock()
    mock_resp.json.return_value = {"value": []}
    mock_resp.status_code = 200
    mock_resp.headers = {"ETag": '"etag123"'}
    client.get.return_value = mock_resp
    client.delete.return_value = mock_resp
    client._auth_headers.return_value = {
        "Authorization": "Bearer fake",
        "Content-Type": "application/json",
    }
    return client


def _full_cfg():
    return {
        "targets": {
            "users": [
                {"upn": "user1@test.com", "role": "Nurse"},
            ]
        },
        "files": {
            "oneDrive": {
                "enabled": True,
                "target_user": "user1@test.com",
                "folders": ["Clinical Ops"],
            },
            "sharePoint": {"enabled": False},
        },
        "teams": {
            "enabled": True,
            "team_id": "team-123",
            "channels": [
                {"display_name": "Care Updates"},
            ],
        },
        "chats": {"enabled": True},
        "sharepoint": {"enabled": True, "sites": []},
        "planner": {
            "enabled": True,
            "group_id": "group-123",
        },
    }


# ---------------------------------------------------------------------------
# Tests — toggle flags
# ---------------------------------------------------------------------------


class TestCleanupFlags:
    def test_all_disabled(self):
        client = _dry_client()
        cfg = _full_cfg()
        actions = cleanup(
            client, cfg, "run001",
            mail=False,
            files=False,
            calendar=False,
            teams=False,
            chats=False,
            sharepoint=False,
            planner=False,
            team_group=False,
        )
        assert actions == []

    def test_mail_only(self):
        client = _dry_client()
        cfg = _full_cfg()
        # Should not raise even with other sections missing
        actions = cleanup(
            client, cfg, "run001",
            mail=True,
            files=False,
            calendar=False,
            teams=False,
            chats=False,
            sharepoint=False,
            planner=False,
            team_group=False,
        )
        # Will be empty because mock returns no messages
        assert isinstance(actions, list)

    def test_teams_only(self):
        client = _dry_client()
        cfg = _full_cfg()
        actions = cleanup(
            client, cfg, "run001",
            mail=False,
            files=False,
            calendar=False,
            teams=True,
            chats=False,
            sharepoint=False,
            planner=False,
            team_group=False,
        )
        assert isinstance(actions, list)


# ---------------------------------------------------------------------------
# Tests — cleanup with matching content
# ---------------------------------------------------------------------------


class TestCleanupWithContent:
    def test_cleanup_teams_channels(self):
        """Verify Teams cleanup queries channels and deletes matches."""
        client = _dry_client()
        # Return a channel that matches
        channel_resp = MagicMock()
        channel_resp.json.return_value = {
            "value": [{"id": "ch-123", "displayName": "Care Updates"}]
        }
        client.get.return_value = channel_resp

        cfg = _full_cfg()
        actions = cleanup(
            client, cfg, "run001",
            mail=False, files=False, calendar=False,
            teams=True, chats=False, sharepoint=False, planner=False,
            team_group=False,
        )
        delete_actions = [a for a in actions if a["action"] == "delete_channel"]
        assert len(delete_actions) == 1
        assert delete_actions[0]["channel"] == "Care Updates"

    def test_cleanup_planner_plans(self):
        """Verify Planner cleanup queries and deletes tagged plans."""
        client = _dry_client()

        plan_resp = MagicMock()
        plan_resp.json.return_value = {
            "value": [
                {"id": "plan-001", "title": "[DEMO-SEED:run001] Clinical Sprint"}
            ]
        }
        plan_resp.headers = {"ETag": '"etag-abc"'}
        client.get.return_value = plan_resp

        cfg = _full_cfg()
        actions = cleanup(
            client, cfg, "run001",
            mail=False, files=False, calendar=False,
            teams=False, chats=False, sharepoint=False, planner=True,
            team_group=False,
        )
        delete_actions = [a for a in actions if a["action"] == "delete_plan"]
        assert len(delete_actions) == 1
        assert "Clinical Sprint" in delete_actions[0]["title"]

    def test_cleanup_sharepoint_groups(self):
        """Verify SharePoint cleanup finds and deletes tagged groups."""
        client = _dry_client()

        group_resp = MagicMock()
        group_resp.json.return_value = {
            "value": [
                {
                    "id": "grp-001",
                    "displayName": "[DEMO-SEED:run001] Clinical Ops Hub",
                }
            ]
        }
        # For site lookup, return a site ID
        site_resp = MagicMock()
        site_resp.json.return_value = {"id": "site-001"}

        # First call gets groups, subsequent calls get site details
        client.get.side_effect = [group_resp, site_resp, MagicMock(json=MagicMock(return_value={"value": []})), MagicMock(json=MagicMock(return_value={"value": []}))]

        cfg = _full_cfg()
        actions = cleanup(
            client, cfg, "run001",
            mail=False, files=False, calendar=False,
            teams=False, chats=False, sharepoint=True, planner=False,
            team_group=False,
        )
        delete_actions = [a for a in actions if a["action"] == "delete_group_site"]
        assert len(delete_actions) == 1

    def test_cleanup_team_group_deletes_seeded_group(self):
        """Verify team group cleanup deletes groups created by the setup wizard."""
        client = _dry_client()

        # Return group details showing it was created by the tool
        group_resp = MagicMock()
        group_resp.json.return_value = {
            "id": "team-123",
            "displayName": "Contoso Health System",
            "description": "Demo team for Contoso Health System",
        }
        client.get.return_value = group_resp

        cfg = _full_cfg()
        actions = cleanup(
            client, cfg, "run001",
            mail=False, files=False, calendar=False,
            teams=False, chats=False, sharepoint=False, planner=False,
            team_group=True,
        )
        delete_actions = [a for a in actions if a["action"] == "delete_team_group"]
        # Both teams.team_id ("team-123") and planner.group_id ("group-123")
        # are candidates; both should be deleted since the mock returns
        # the "Demo team for" description for all lookups.
        deleted_ids = {a["group_id"] for a in delete_actions}
        assert "team-123" in deleted_ids
        assert "group-123" in deleted_ids
        assert all(a["display_name"] == "Contoso Health System" for a in delete_actions)

    def test_cleanup_team_group_skips_preexisting_group(self):
        """Verify team group cleanup does NOT delete pre-existing tenant groups."""
        client = _dry_client()

        # Return group details without "Demo team for" description
        group_resp = MagicMock()
        group_resp.json.return_value = {
            "id": "team-123",
            "displayName": "Contoso Health System",
            "description": "Original tenant team",
        }
        client.get.return_value = group_resp

        cfg = _full_cfg()
        actions = cleanup(
            client, cfg, "run001",
            mail=False, files=False, calendar=False,
            teams=False, chats=False, sharepoint=False, planner=False,
            team_group=True,
        )
        delete_actions = [a for a in actions if a["action"] == "delete_team_group"]
        assert len(delete_actions) == 0
