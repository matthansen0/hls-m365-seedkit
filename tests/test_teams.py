"""Tests for m365seed.teams — Teams channels & posts seeding."""

from __future__ import annotations

from unittest.mock import MagicMock

from m365seed.teams import seed_teams


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def _dry_client():
    """Return a mock GraphClient in dry-run mode."""
    client = MagicMock()
    client.dry_run = True

    mock_resp = MagicMock()
    mock_resp.json.return_value = {"id": "dry-run-channel-id"}
    mock_resp.status_code = 200
    client.get.return_value = mock_resp
    client.post.return_value = mock_resp
    return client


def _base_cfg(posts=None):
    if posts is None:
        posts = [{"message": "Welcome to the channel"}]
    return {
        "targets": {
            "users": [{"upn": "user1@test.com", "role": "Manager"}],
        },
        "teams": {
            "enabled": True,
            "team_id": "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
            "channels": [
                {
                    "channel_id": "updates-001",
                    "display_name": "Updates",
                    "description": "Team updates",
                    "posts": posts,
                }
            ],
        },
    }


# ---------------------------------------------------------------------------
# Tests — disabled / empty
# ---------------------------------------------------------------------------


class TestTeamsDisabled:
    def test_disabled_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["teams"]["enabled"] = False
        actions = seed_teams(client, cfg, "healthcare", "run001")
        assert actions == []

    def test_no_team_id_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["teams"]["team_id"] = ""
        actions = seed_teams(client, cfg, "healthcare", "run001")
        assert actions == []

    def test_no_channels_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["teams"]["channels"] = []
        actions = seed_teams(client, cfg, "healthcare", "run001")
        assert actions == []


# ---------------------------------------------------------------------------
# Tests — dry-run seeding
# ---------------------------------------------------------------------------


class TestTeamsDryRun:
    def test_creates_channel_and_posts(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_teams(client, cfg, "healthcare", "run001")
        create_actions = [a for a in actions if a["action"] == "create_channel"]
        assert len(create_actions) == 1
        assert create_actions[0]["channel"] == "Updates"
        post_actions = [a for a in actions if a["action"] == "post_message"]
        assert len(post_actions) == 1

    def test_posts_as_dicts(self):
        """Posts as dicts with 'message' key (from config generator)."""
        client = _dry_client()
        cfg = _base_cfg(posts=[
            {"message": "First post"},
            {"message": "Second post"},
        ])
        actions = seed_teams(client, cfg, "healthcare", "run001")
        post_actions = [a for a in actions if a["action"] == "post_message"]
        assert len(post_actions) == 2

    def test_posts_as_plain_strings(self):
        """Posts as plain strings (from theme enrichment) should not crash."""
        client = _dry_client()
        cfg = _base_cfg(posts=[
            "Plain string post one",
            "Plain string post two",
            "Plain string post three",
        ])
        actions = seed_teams(client, cfg, "healthcare", "run001")
        post_actions = [a for a in actions if a["action"] == "post_message"]
        assert len(post_actions) == 3

    def test_multiple_channels(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["teams"]["channels"].append({
            "channel_id": "ops-002",
            "display_name": "Operations",
            "description": "Ops channel",
            "posts": [{"message": "Ops update"}],
        })
        actions = seed_teams(client, cfg, "healthcare", "run001")
        create_actions = [a for a in actions if a["action"] == "create_channel"]
        assert len(create_actions) == 2
