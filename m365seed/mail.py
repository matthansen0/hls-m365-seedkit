"""Email seeding — send synthetic theme-specific email threads."""

from __future__ import annotations

import base64
import logging
from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader

from m365seed.graph import GraphClient, GRAPH_V1
from m365seed.theme_content import get_mail_threads

logger = logging.getLogger("m365seed.mail")

TEMPLATES_DIR = Path(__file__).parent / "templates"


# ---------------------------------------------------------------------------
# Template rendering
# ---------------------------------------------------------------------------


def _jinja_env(theme: str) -> Environment:
    """Return a Jinja2 environment rooted at the templates directory."""
    search_path = [
        str(TEMPLATES_DIR / theme),
        str(TEMPLATES_DIR / "healthcare"),  # fallback
    ]
    return Environment(loader=FileSystemLoader(search_path), autoescape=False)


def render_email_body(
    theme: str,
    thread_id: str,
    message_index: int,
    participants: list[str],
    subject: str,
    run_id: str,
) -> str:
    """Render a synthetic email body from a Jinja2 template."""
    env = _jinja_env(theme)
    try:
        tpl = env.get_template("email_body.html.j2")
    except Exception:
        # Fallback to a simple plaintext body
        return _fallback_body(thread_id, message_index, subject, run_id)

    return tpl.render(
        thread_id=thread_id,
        message_index=message_index,
        participants=participants,
        subject=subject,
        run_id=run_id,
    )


def _fallback_body(thread_id: str, idx: int, subject: str, run_id: str) -> str:
    return (
        f"<p>This is message #{idx + 1} in thread <b>{thread_id}</b>.</p>"
        f"<p>Subject: {subject}</p>"
        f"<p><em>Demo content — synthetic, no patient data. RunId: {run_id}</em></p>"
    )


# ---------------------------------------------------------------------------
# Attachment helpers
# ---------------------------------------------------------------------------


def _make_text_attachment(
    name: str, text_content: str
) -> dict[str, Any]:
    """Create a Graph-compatible file attachment dict."""
    encoded = base64.b64encode(text_content.encode("utf-8")).decode("ascii")
    return {
        "@odata.type": "#microsoft.graph.fileAttachment",
        "name": name,
        "contentType": "text/plain",
        "contentBytes": encoded,
    }


# ---------------------------------------------------------------------------
# Idempotency helpers
# ---------------------------------------------------------------------------


def _seed_subject(subject: str, thread_id: str, run_id: str) -> str:
    """Prefix the subject with a deterministic tag for idempotency."""
    tag = f"[DEMO-SEED:{run_id}:{thread_id}]"
    if tag in subject:
        return subject
    return f"{tag} {subject}"


def _thread_already_exists(
    client: GraphClient, sender_upn: str, thread_id: str, run_id: str
) -> bool:
    """Check if emails with this thread tag already exist in the sender's mailbox."""
    tag = f"DEMO-SEED:{run_id}:{thread_id}"
    filter_q = f"subject:'{tag}'"
    try:
        resp = client.get(
            f"/users/{sender_upn}/messages",
            params={"$search": f'"{filter_q}"', "$top": "1"},
            headers={**client._auth_headers(), "ConsistencyLevel": "eventual"},
        )
        data = resp.json()
        return len(data.get("value", [])) > 0
    except Exception as exc:
        logger.warning("Idempotency check failed for %s: %s", thread_id, exc)
        return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def seed_mail(
    client: GraphClient,
    cfg: dict[str, Any],
    theme: str,
    run_id: str,
) -> list[dict[str, Any]]:
    """Send all configured email threads. Returns a list of action records."""
    threads = cfg.get("mail", {}).get("threads", [])
    if not threads:
        logger.info("No mail threads configured — skipping.")
        return []

    # Build lookup for theme-specific attachment content
    theme_threads = {t["thread_id"]: t for t in get_mail_threads(theme)}

    actions: list[dict[str, Any]] = []

    for thread_cfg in threads:
        thread_id = thread_cfg["thread_id"]
        subject_raw = thread_cfg["subject"]
        participants = thread_cfg["participants"]
        num_messages = thread_cfg.get("messages", 1)
        include_attachments = thread_cfg.get("include_attachments", False)

        subject = _seed_subject(subject_raw, thread_id, run_id)

        # Use first participant as sender for the initial message,
        # then alternate for replies.
        sender = participants[0]

        # Idempotency: check if thread already exists
        if not client.dry_run and _thread_already_exists(
            client, sender, thread_id, run_id
        ):
            logger.info(
                "Thread '%s' already exists in %s's mailbox — skipping.",
                thread_id,
                sender,
            )
            actions.append(
                {"action": "skip", "thread_id": thread_id, "reason": "already_exists"}
            )
            continue

        for msg_idx in range(num_messages):
            # Alternate sender across participants
            current_sender = participants[msg_idx % len(participants)]
            recipients = [p for p in participants if p != current_sender]

            body_html = render_email_body(
                theme, thread_id, msg_idx, participants, subject_raw, run_id
            )

            to_recipients = [
                {"emailAddress": {"address": r}} for r in recipients
            ]

            mail_payload: dict[str, Any] = {
                "message": {
                    "subject": subject,
                    "body": {"contentType": "HTML", "content": body_html},
                    "toRecipients": to_recipients,
                    "internetMessageHeaders": [
                        {
                            "name": "X-DemoSeed-RunId",
                            "value": run_id,
                        },
                        {
                            "name": "X-DemoSeed-ThreadId",
                            "value": thread_id,
                        },
                    ],
                },
                "saveToSentItems": "true",
            }

            # Attachments — use theme-specific content when available
            if include_attachments and msg_idx == 0:
                tt = theme_threads.get(thread_id, {})
                attachment_name = tt.get(
                    "attachment_name", f"demo-{thread_id}-attachment.txt"
                )
                attachment_text = tt.get(
                    "attachment_content",
                    f"Synthetic attachment for thread {thread_id}.\n"
                    "Demo content — synthetic, no patient data.\n",
                )
                mail_payload["message"]["attachments"] = [
                    _make_text_attachment(attachment_name, attachment_text)
                ]

            logger.info(
                "Sending message %d/%d of thread '%s' from %s",
                msg_idx + 1,
                num_messages,
                thread_id,
                current_sender,
            )

            client.post(
                f"/users/{current_sender}/sendMail",
                json_body=mail_payload,
            )

            actions.append(
                {
                    "action": "send_mail",
                    "thread_id": thread_id,
                    "message_index": msg_idx,
                    "sender": current_sender,
                    "recipients": recipients,
                }
            )

    return actions
