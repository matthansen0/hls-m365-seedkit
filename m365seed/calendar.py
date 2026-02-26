"""Calendar seeding — create synthetic theme-specific calendar events."""

from __future__ import annotations

import logging
from datetime import datetime, timedelta, timezone
from typing import Any

from m365seed.graph import GraphClient
from m365seed.theme_content import get_calendar_events

logger = logging.getLogger("m365seed.calendar")

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

DISCLAIMER = "Demo content — synthetic, no patient data."


def _build_event_body(
    event_cfg: dict[str, Any],
    run_id: str,
) -> dict[str, Any]:
    """Build a Graph calendar-event payload."""
    subject = f"[DEMO-SEED:{run_id}:{event_cfg['event_id']}] {event_cfg['subject']}"
    organizer = event_cfg["organizer"]
    attendees = event_cfg.get("attendees", [])
    duration = event_cfg.get("duration_minutes", 30)
    recurrence_type = event_cfg.get("recurrence", None)

    # Start tomorrow at 09:00 UTC for determinism
    start_dt = datetime.now(timezone.utc).replace(
        hour=9, minute=0, second=0, microsecond=0
    ) + timedelta(days=1)
    end_dt = start_dt + timedelta(minutes=duration)

    body: dict[str, Any] = {
        "subject": subject,
        "body": {
            "contentType": "HTML",
            "content": (
                f"<p><strong>{event_cfg['subject']}</strong></p>"
                f"{event_cfg.get('body', '')}"
                f"<p>Organizer: {organizer}</p>"
                f"<p><em>{DISCLAIMER}</em></p>"
            ),
        },
        "start": {
            "dateTime": start_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": "UTC",
        },
        "end": {
            "dateTime": end_dt.strftime("%Y-%m-%dT%H:%M:%S"),
            "timeZone": "UTC",
        },
        "attendees": [
            {
                "emailAddress": {"address": a},
                "type": "required",
            }
            for a in attendees
        ],
    }

    # Recurrence (simplified — daily, weekly, or monthly)
    if recurrence_type in ("daily", "weekly", "monthly"):
        body["recurrence"] = {
            "pattern": {
                "type": recurrence_type,
                "interval": 1,
            },
            "range": {
                "type": "endDate",
                "startDate": start_dt.strftime("%Y-%m-%d"),
                "endDate": (start_dt + timedelta(days=30)).strftime("%Y-%m-%d"),
            },
        }

    # Teams online meeting — adds a join link to the event (v1.0 GA)
    if event_cfg.get("is_online_meeting", False):
        body["isOnlineMeeting"] = True
        body["onlineMeetingProvider"] = "teamsForBusiness"

    return body


def _event_exists(
    client: GraphClient,
    organizer_upn: str,
    event_id: str,
    run_id: str,
) -> bool:
    """Check if a seeded event already exists in the organizer's calendar."""
    tag = f"DEMO-SEED:{run_id}:{event_id}"
    try:
        resp = client.get(
            f"/users/{organizer_upn}/events",
            params={
                "$filter": f"startsWith(subject, '[{tag}]')",
                "$top": "1",
                "$select": "id,subject",
            },
        )
        data = resp.json()
        return len(data.get("value", [])) > 0
    except Exception as exc:
        logger.warning("Calendar idempotency check failed for %s: %s", event_id, exc)
        return False


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def seed_calendar(
    client: GraphClient,
    cfg: dict[str, Any],
    theme: str,
    run_id: str,
) -> list[dict[str, Any]]:
    """Create all configured calendar events. Returns a list of action records."""
    cal_cfg = cfg.get("calendar", {})
    if not cal_cfg.get("enabled"):
        logger.info("Calendar seeding is disabled — skipping.")
        return []

    events = cal_cfg.get("events", [])
    if not events:
        logger.info("No calendar events configured — skipping.")
        return []

    # Enrich config events with theme-specific body text
    theme_events = {e["event_id"]: e for e in get_calendar_events(theme)}
    for event_cfg in events:
        eid = event_cfg.get("event_id", "")
        if eid in theme_events:
            te = theme_events[eid]
            if "body" not in event_cfg and "body" in te:
                event_cfg["body"] = te["body"]
            # Allow theme to provide subject fallback
            if "subject" not in event_cfg and "subject" in te:
                event_cfg["subject"] = te["subject"]

    actions: list[dict[str, Any]] = []

    for event_cfg in events:
        event_id = event_cfg["event_id"]
        organizer = event_cfg["organizer"]

        # Idempotency
        if not client.dry_run and _event_exists(client, organizer, event_id, run_id):
            logger.info(
                "Event '%s' already exists for %s — skipping.",
                event_id,
                organizer,
            )
            actions.append(
                {"action": "skip", "event_id": event_id, "reason": "already_exists"}
            )
            continue

        payload = _build_event_body(event_cfg, run_id)

        logger.info(
            "Creating event '%s' for organizer %s",
            event_id,
            organizer,
        )

        client.post(
            f"/users/{organizer}/events",
            json_body=payload,
        )

        actions.append(
            {
                "action": "create_event",
                "event_id": event_id,
                "organizer": organizer,
                "subject": payload["subject"],
            }
        )

    return actions
