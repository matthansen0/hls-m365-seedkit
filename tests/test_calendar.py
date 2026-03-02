"""Tests for m365seed.calendar — event building and idempotency."""

from unittest.mock import MagicMock, patch

import httpx
import pytest

from m365seed.calendar import (
    _build_event_body,
    seed_calendar,
)


# ---------------------------------------------------------------------------
# Event body building
# ---------------------------------------------------------------------------


class TestBuildEventBody:
    def test_basic_event(self):
        event_cfg = {
            "event_id": "e-001",
            "subject": "Clinical Standup",
            "organizer": "org@test.com",
            "attendees": ["a@test.com", "b@test.com"],
            "duration_minutes": 15,
        }
        body = _build_event_body(event_cfg, "run-001")
        assert "[DEMO-SEED:run-001:e-001]" in body["subject"]
        assert body["attendees"][0]["emailAddress"]["address"] == "a@test.com"
        assert body["start"]["timeZone"] == "UTC"
        assert "synthetic" in body["body"]["content"].lower()

    def test_recurrence_daily(self):
        event_cfg = {
            "event_id": "e-002",
            "subject": "Daily Huddle",
            "organizer": "org@test.com",
            "recurrence": "daily",
            "duration_minutes": 10,
        }
        body = _build_event_body(event_cfg, "run-002")
        assert body["recurrence"]["pattern"]["type"] == "daily"
        assert body["recurrence"]["pattern"]["interval"] == 1

    def test_no_recurrence(self):
        event_cfg = {
            "event_id": "e-003",
            "subject": "One-off Meeting",
            "organizer": "org@test.com",
            "duration_minutes": 60,
        }
        body = _build_event_body(event_cfg, "run-003")
        assert "recurrence" not in body

    def test_online_meeting(self):
        event_cfg = {
            "event_id": "e-004",
            "subject": "Virtual Consult",
            "organizer": "doc@test.com",
            "attendees": ["nurse@test.com"],
            "duration_minutes": 20,
            "is_online_meeting": True,
        }
        body = _build_event_body(event_cfg, "run-004")
        assert body["isOnlineMeeting"] is True
        assert body["onlineMeetingProvider"] == "teamsForBusiness"

    def test_online_meeting_false(self):
        event_cfg = {
            "event_id": "e-005",
            "subject": "In-Person Huddle",
            "organizer": "doc@test.com",
            "duration_minutes": 15,
            "is_online_meeting": False,
        }
        body = _build_event_body(event_cfg, "run-005")
        assert "isOnlineMeeting" not in body

    def test_online_meeting_default(self):
        event_cfg = {
            "event_id": "e-006",
            "subject": "Default Meeting",
            "organizer": "doc@test.com",
            "duration_minutes": 30,
        }
        body = _build_event_body(event_cfg, "run-006")
        assert "isOnlineMeeting" not in body


# ---------------------------------------------------------------------------
# Seed calendar dry-run
# ---------------------------------------------------------------------------


class TestSeedCalendarDryRun:
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

    def test_calendar_disabled(self):
        client = self._make_dry_client()
        cfg = {"calendar": {"enabled": False}}
        actions = seed_calendar(client, cfg, "healthcare", "run-001")
        assert actions == []

    def test_calendar_no_events(self):
        client = self._make_dry_client()
        cfg = {"calendar": {"enabled": True, "events": []}}
        actions = seed_calendar(client, cfg, "healthcare", "run-001")
        assert actions == []

    def test_calendar_dry_run(self):
        client = self._make_dry_client()
        cfg = {
            "calendar": {
                "enabled": True,
                "events": [
                    {
                        "event_id": "e-001",
                        "subject": "Test Meeting",
                        "organizer": "org@test.com",
                        "attendees": ["a@test.com"],
                        "duration_minutes": 30,
                    }
                ],
            }
        }
        actions = seed_calendar(client, cfg, "healthcare", "run-dry")
        assert len(actions) == 1
        assert actions[0]["action"] == "create_event"
        assert actions[0]["event_id"] == "e-001"

    def test_seed_calendar_handles_404_gracefully(self):
        """A 404 from event creation should log an error action, not crash."""
        request = httpx.Request("POST", "https://graph.microsoft.com/v1.0/users/gone@test.com/events")
        response = httpx.Response(404, request=request)
        error = httpx.HTTPStatusError("Not Found", request=request, response=response)

        client = MagicMock(spec=["post", "dry_run"])
        client.dry_run = True
        client.post.side_effect = error

        cfg = {
            "calendar": {
                "enabled": True,
                "events": [
                    {
                        "event_id": "e-001",
                        "subject": "Test Meeting",
                        "organizer": "gone@test.com",
                        "attendees": ["a@test.com"],
                        "duration_minutes": 30,
                    }
                ],
            }
        }
        actions = seed_calendar(client, cfg, "healthcare", "run-err")
        assert len(actions) == 1
        assert actions[0]["action"] == "error"
        assert "Not Found" in actions[0]["error"]
