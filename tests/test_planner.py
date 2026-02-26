"""Tests for m365seed.planner — Planner plan/task seeding."""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock

from m365seed.planner import seed_planner


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
    return client


def _base_cfg():
    return {
        "targets": {
            "users": [
                {"upn": "user1@test.com", "role": "Nurse"},
                {"upn": "user2@test.com", "role": "Doctor"},
            ]
        },
        "planner": {
            "enabled": True,
            "group_id": "group-id-123",
            "plans": [
                {
                    "title": "Clinical Sprint",
                    "buckets": [
                        {
                            "name": "To Do",
                            "tasks": [
                                {
                                    "title": "Review SOP",
                                    "priority": 3,
                                    "percent_complete": 0,
                                    "assignees": ["user1@test.com"],
                                },
                                {
                                    "title": "Update checklist",
                                    "priority": 5,
                                    "percent_complete": 50,
                                    "assignees": ["user2@test.com"],
                                },
                            ],
                        },
                        {
                            "name": "Done",
                            "tasks": [
                                {
                                    "title": "Complete audit",
                                    "priority": 3,
                                    "percent_complete": 100,
                                },
                            ],
                        },
                    ],
                }
            ],
        },
    }


# ---------------------------------------------------------------------------
# Tests — disabled / empty
# ---------------------------------------------------------------------------


class TestPlannerDisabled:
    def test_disabled_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["planner"]["enabled"] = False
        actions = seed_planner(client, cfg, "healthcare", "run001")
        assert actions == []

    def test_no_group_id_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["planner"]["group_id"] = ""
        actions = seed_planner(client, cfg, "healthcare", "run001")
        assert actions == []

    def test_no_plans_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["planner"]["plans"] = []
        actions = seed_planner(client, cfg, "healthcare", "run001")
        assert actions == []

    def test_missing_planner_section(self):
        client = _dry_client()
        cfg = {"targets": {"users": [{"upn": "a@b.com"}]}}
        actions = seed_planner(client, cfg, "healthcare", "run001")
        assert actions == []


# ---------------------------------------------------------------------------
# Tests — dry-run seeding
# ---------------------------------------------------------------------------


class TestPlannerDryRun:
    def test_creates_plan(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_planner(client, cfg, "healthcare", "run001")
        plan_actions = [a for a in actions if a["action"] == "create_plan"]
        assert len(plan_actions) == 1
        assert plan_actions[0]["plan"] == "Clinical Sprint"
        assert plan_actions[0]["group_id"] == "group-id-123"

    def test_creates_buckets(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_planner(client, cfg, "healthcare", "run001")
        bucket_actions = [a for a in actions if a["action"] == "create_bucket"]
        assert len(bucket_actions) == 2
        names = [a["bucket"] for a in bucket_actions]
        assert "To Do" in names
        assert "Done" in names

    def test_creates_tasks(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_planner(client, cfg, "healthcare", "run001")
        task_actions = [a for a in actions if a["action"] == "create_task"]
        assert len(task_actions) == 3
        titles = [a["task"] for a in task_actions]
        assert "Review SOP" in titles
        assert "Update checklist" in titles
        assert "Complete audit" in titles

    def test_task_in_correct_bucket(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_planner(client, cfg, "healthcare", "run001")
        task_actions = [a for a in actions if a["action"] == "create_task"]
        # All tasks should have bucket name
        for ta in task_actions:
            assert "bucket" in ta
