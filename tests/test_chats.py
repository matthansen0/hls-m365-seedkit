"""Tests for m365seed.chats — Teams chat seeding."""

from __future__ import annotations

import pytest
from unittest.mock import MagicMock, patch

from m365seed.chats import seed_chats, _create_chat, _send_chat_message


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _dry_client():
    """Return a mock GraphClient in dry-run mode."""
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
                {"upn": "user3@test.com", "role": "Coordinator"},
            ]
        },
        "chats": {
            "enabled": True,
            "conversations": [
                {
                    "conversation_id": "test-chat-001",
                    "type": "group",
                    "topic": "Test Chat",
                    "members": ["user1@test.com", "user2@test.com", "user3@test.com"],
                    "messages": [
                        {"sender": "user1@test.com", "text": "Hello team"},
                        {"sender": "user2@test.com", "text": "Hi there"},
                    ],
                }
            ],
        },
    }


# ---------------------------------------------------------------------------
# Tests — disabled / empty
# ---------------------------------------------------------------------------


class TestSeedChatsDisabled:
    def test_disabled_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["chats"]["enabled"] = False
        actions = seed_chats(client, cfg, "healthcare", "run001")
        assert actions == []

    def test_no_conversations_returns_empty(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["chats"]["conversations"] = []
        actions = seed_chats(client, cfg, "healthcare", "run001")
        assert actions == []

    def test_missing_chats_section(self):
        client = _dry_client()
        cfg = {"targets": {"users": [{"upn": "a@b.com"}]}}
        actions = seed_chats(client, cfg, "healthcare", "run001")
        assert actions == []


# ---------------------------------------------------------------------------
# Tests — dry-run seeding
# ---------------------------------------------------------------------------


class TestSeedChatsDryRun:
    def test_creates_chat_action(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_chats(client, cfg, "healthcare", "run001")
        create_actions = [a for a in actions if a["action"] == "create_chat"]
        assert len(create_actions) == 1
        assert create_actions[0]["conversation_id"] == "test-chat-001"
        assert create_actions[0]["api"] == "beta"

    def test_sends_messages(self):
        client = _dry_client()
        cfg = _base_cfg()
        actions = seed_chats(client, cfg, "healthcare", "run001")
        msg_actions = [a for a in actions if a["action"] == "send_chat_message"]
        assert len(msg_actions) == 2

    def test_one_on_one_chat(self):
        client = _dry_client()
        cfg = _base_cfg()
        cfg["chats"]["conversations"] = [
            {
                "conversation_id": "dm-001",
                "type": "oneOnOne",
                "members": ["user1@test.com", "user2@test.com"],
                "messages": [],
            }
        ]
        actions = seed_chats(client, cfg, "healthcare", "run001")
        create_actions = [a for a in actions if a["action"] == "create_chat"]
        assert len(create_actions) == 1
        assert create_actions[0]["type"] == "oneOnOne"
