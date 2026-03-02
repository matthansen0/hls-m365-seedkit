"""Tests for m365seed.mail — idempotency keys and template rendering."""

from unittest.mock import MagicMock, patch

import httpx
import pytest

from m365seed.mail import (
    _seed_subject,
    _fallback_body,
    _make_text_attachment,
    render_email_body,
    seed_mail,
)


# ---------------------------------------------------------------------------
# Idempotency key tests
# ---------------------------------------------------------------------------


class TestSeedSubject:
    def test_prefix_added(self):
        result = _seed_subject("Care Coordination Update", "thread-001", "run-001")
        assert result == "[DEMO-SEED:run-001:thread-001] Care Coordination Update"

    def test_no_double_prefix(self):
        already_tagged = "[DEMO-SEED:run-001:thread-001] Care Coordination Update"
        result = _seed_subject(already_tagged, "thread-001", "run-001")
        assert result == already_tagged  # unchanged

    def test_different_run_id_adds_new_prefix(self):
        tagged = "[DEMO-SEED:run-001:thread-001] Care Coordination Update"
        result = _seed_subject(tagged, "thread-001", "run-002")
        assert "[DEMO-SEED:run-002:thread-001]" in result


# ---------------------------------------------------------------------------
# Template rendering tests
# ---------------------------------------------------------------------------


class TestRenderEmailBody:
    def test_fallback_body(self):
        body = _fallback_body("t-001", 0, "Test Subject", "run-123")
        assert "t-001" in body
        assert "run-123" in body
        assert "synthetic" in body.lower()
        assert "no patient data" in body.lower()

    def test_render_email_body_healthcare(self):
        body = render_email_body(
            theme="healthcare",
            thread_id="t-001",
            message_index=0,
            participants=["a@test.com", "b@test.com"],
            subject="Test Thread",
            run_id="run-001",
        )
        # Should contain the disclaimer
        assert "synthetic" in body.lower()

    def test_render_email_body_pharma(self):
        body = render_email_body(
            theme="pharma",
            thread_id="t-002",
            message_index=0,
            participants=["x@test.com"],
            subject="Pharma Thread",
            run_id="run-002",
        )
        assert "run-002" in body

    def test_render_email_body_invalid_theme_uses_fallback(self):
        body = render_email_body(
            theme="nonexistent_theme",
            thread_id="t-003",
            message_index=0,
            participants=["y@test.com"],
            subject="Invalid Theme",
            run_id="run-003",
        )
        # Should still return content (fallback)
        assert "run-003" in body


# ---------------------------------------------------------------------------
# Attachment tests
# ---------------------------------------------------------------------------


class TestAttachments:
    def test_make_text_attachment(self):
        att = _make_text_attachment("test.txt", "Hello World")
        assert att["name"] == "test.txt"
        assert att["contentType"] == "text/plain"
        assert att["@odata.type"] == "#microsoft.graph.fileAttachment"
        # contentBytes should be base64 encoded
        import base64

        decoded = base64.b64decode(att["contentBytes"]).decode("utf-8")
        assert decoded == "Hello World"


# ---------------------------------------------------------------------------
# seed_mail integration test (dry run)
# ---------------------------------------------------------------------------


class TestSeedMailDryRun:
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

    def test_seed_mail_dry_run_no_threads(self):
        client = self._make_dry_client()
        actions = seed_mail(client, {}, "healthcare", "run-001")
        assert actions == []

    def test_seed_mail_dry_run_with_threads(self):
        client = self._make_dry_client()
        cfg = {
            "mail": {
                "threads": [
                    {
                        "thread_id": "t-001",
                        "subject": "Test Thread",
                        "participants": ["a@test.com", "b@test.com"],
                        "messages": 2,
                        "include_attachments": True,
                    }
                ]
            }
        }
        actions = seed_mail(client, cfg, "healthcare", "run-dry")
        assert len(actions) == 2
        assert all(a["action"] == "send_mail" for a in actions)
        # Check sender alternation
        assert actions[0]["sender"] == "a@test.com"
        assert actions[1]["sender"] == "b@test.com"

    def test_seed_mail_handles_404_gracefully(self):
        """A 404 from sendMail should log an error action, not crash."""
        request = httpx.Request("POST", "https://graph.microsoft.com/v1.0/users/gone@test.com/sendMail")
        response = httpx.Response(404, request=request)
        error = httpx.HTTPStatusError("Not Found", request=request, response=response)

        client = MagicMock(spec=["post", "dry_run"])
        client.dry_run = True
        client.post.side_effect = error

        cfg = {
            "mail": {
                "enabled": True,
                "threads": [
                    {
                        "thread_id": "t-001",
                        "subject": "Test Thread",
                        "participants": ["gone@test.com"],
                        "messages": 1,
                    }
                ]
            }
        }
        actions = seed_mail(client, cfg, "healthcare", "run-err")
        assert len(actions) == 1
        assert actions[0]["action"] == "error"
        assert "Not Found" in actions[0]["error"]
