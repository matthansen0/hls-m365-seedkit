"""Teams seeding — create channels and posts via Microsoft Graph /beta.

⚠️  This module uses Microsoft Graph **beta** endpoints and is **off by
default**.  Enable with ``--enable-beta-teams``.  Beta endpoints may change
or break without notice.
"""

from __future__ import annotations

import logging
from typing import Any

from m365seed.graph import GraphClient, GRAPH_BETA
from m365seed.theme_content import get_teams_channels

logger = logging.getLogger("m365seed.teams")

DISCLAIMER = "Demo content — synthetic, no patient data."


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _channel_exists(
    client: GraphClient,
    team_id: str,
    display_name: str,
) -> str | None:
    """Return channel ID if a channel with *display_name* exists, else None.

    Uses the v1.0 endpoint (listing channels is GA).
    """
    try:
        resp = client.get(
            f"/teams/{team_id}/channels",
            params={"$filter": f"displayName eq '{display_name}'", "$top": "1"},
        )
        data = resp.json()
        channels = data.get("value", [])
        if channels:
            return channels[0]["id"]
    except Exception as exc:
        logger.warning("Channel existence check failed: %s", exc)
    return None


def _create_channel(
    client: GraphClient,
    team_id: str,
    display_name: str,
    description: str,
) -> str:
    """Create a channel and return its ID.  Uses **/beta**."""
    payload = {
        "displayName": display_name,
        "description": description,
        "membershipType": "standard",
    }
    resp = client.post(
        f"/teams/{team_id}/channels",
        json_body=payload,
        base=GRAPH_BETA,
    )
    return resp.json().get("id", "")


def _post_message(
    client: GraphClient,
    team_id: str,
    channel_id: str,
    message: str,
    run_id: str,
) -> None:
    """Post a message to a Teams channel.  Uses **/beta**."""
    payload = {
        "body": {
            "contentType": "html",
            "content": (
                f"<p>{message}</p>"
                f"<p><em>{DISCLAIMER} | RunId: {run_id}</em></p>"
            ),
        },
    }
    client.post(
        f"/teams/{team_id}/channels/{channel_id}/messages",
        json_body=payload,
        base=GRAPH_BETA,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def seed_teams(
    client: GraphClient,
    cfg: dict[str, Any],
    theme: str,
    run_id: str,
) -> list[dict[str, Any]]:
    """Seed Teams channels and posts.  Returns action records.

    This function is gated behind the ``teams.enabled`` config flag
    **and** the CLI ``--enable-beta-teams`` argument.
    """
    teams_cfg = cfg.get("teams", {})
    if not teams_cfg.get("enabled"):
        logger.info("Teams seeding is disabled — skipping.")
        return []

    team_id = teams_cfg.get("team_id", "")
    if not team_id:
        logger.warning("No team_id configured — skipping Teams seeding.")
        return []

    channels = teams_cfg.get("channels", [])

    # Enrich config channels with theme-specific descriptions and posts
    theme_channels = {
        c["channel_id"]: c for c in get_teams_channels(theme) if "channel_id" in c
    }
    for ch_cfg in channels:
        cid = ch_cfg.get("channel_id", "")
        if cid in theme_channels:
            tc = theme_channels[cid]
            if not ch_cfg.get("description") and tc.get("description"):
                ch_cfg["description"] = tc["description"]
            if not ch_cfg.get("posts") and tc.get("posts"):
                ch_cfg["posts"] = tc["posts"]
    actions: list[dict[str, Any]] = []

    for ch_cfg in channels:
        display_name = ch_cfg["display_name"]
        description = ch_cfg.get("description", "")

        # Idempotency: check if channel already exists
        existing_id = (
            None if client.dry_run else _channel_exists(client, team_id, display_name)
        )

        if existing_id:
            channel_id = existing_id
            logger.info(
                "Channel '%s' already exists (id=%s) — reusing.",
                display_name,
                channel_id,
            )
            actions.append(
                {
                    "action": "skip_channel",
                    "channel": display_name,
                    "reason": "already_exists",
                }
            )
        else:
            logger.info(
                "[BETA] Creating channel '%s' in team %s",
                display_name,
                team_id,
            )
            channel_id = _create_channel(client, team_id, display_name, description)
            actions.append(
                {
                    "action": "create_channel",
                    "channel": display_name,
                    "channel_id": channel_id,
                    "api": "beta",
                }
            )

        # Post messages
        for post_cfg in ch_cfg.get("posts", []):
            message = post_cfg["message"]
            logger.info(
                "[BETA] Posting message to channel '%s'",
                display_name,
            )
            _post_message(client, team_id, channel_id, message, run_id)
            actions.append(
                {
                    "action": "post_message",
                    "channel": display_name,
                    "api": "beta",
                }
            )

    return actions
