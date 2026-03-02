"""Teams seeding — create channels and posts via Microsoft Graph /beta.

⚠️  This module uses Microsoft Graph **beta** endpoints and is **off by
default**.  Enable with ``--enable-beta-teams``.  Beta endpoints may change
or break without notice.
"""

from __future__ import annotations

import logging
from typing import Any

import httpx

from m365seed.graph import GraphClient, GRAPH_BASE, GRAPH_BETA, build_delegated_client
from m365seed.theme_content import get_teams_channels

logger = logging.getLogger("m365seed.teams")

DISCLAIMER = "Demo content — synthetic, no patient data."


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _ensure_team_members(
    client: GraphClient,
    team_id: str,
    user_upns: list[str],
) -> None:
    """Add users as team members (idempotent — skips existing members).

    Uses the app-only *client* (needs ``Group.ReadWrite.All``).
    """
    # Fetch current members
    existing_ids: set[str] = set()
    try:
        resp = client.get(
            f"/groups/{team_id}/members",
            params={"$select": "id"},
        )
        for m in resp.json().get("value", []):
            existing_ids.add(m.get("id", ""))
    except Exception as exc:
        logger.warning("Could not list team members: %s", exc)

    for upn in user_upns:
        # Resolve UPN → user id
        try:
            user_resp = client.get(
                f"/users/{upn}",
                params={"$select": "id"},
            )
            uid = user_resp.json().get("id", "")
        except Exception as exc:
            logger.warning("Could not resolve %s for team membership: %s", upn, exc)
            continue

        if uid in existing_ids:
            logger.debug("User %s already a team member — skipping.", upn)
            continue

        try:
            client.post(
                f"/groups/{team_id}/members/$ref",
                json_body={
                    "@odata.id": f"{GRAPH_BASE}/v1.0/directoryObjects/{uid}",
                },
            )
            logger.info("Added %s as team member.", upn)
        except Exception as exc:
            logger.warning("Failed to add %s as team member: %s", upn, exc)


def _channel_exists(
    client: GraphClient,
    team_id: str,
    display_name: str,
) -> str | None:
    """Return channel ID if a channel with *display_name* exists, else None.

    Lists all channels and filters client-side because the Graph
    ``/teams/{id}/channels`` endpoint does not support ``$filter``
    on ``displayName``.
    """
    try:
        next_url = f"{GRAPH_BETA}/teams/{team_id}/channels"
        params: dict[str, str] | None = {"$select": "id,displayName"}

        while next_url:
            resp = client.request("GET", next_url, params=params)
            data = resp.json()
            for ch in data.get("value", []):
                if ch.get("displayName", "").lower() == display_name.lower():
                    return ch["id"]

            next_url = data.get("@odata.nextLink")
            params = None
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

    # Detect app-only auth — channel message posting requires delegated auth.
    # Automatically fall back to device_code for message posting.
    auth_mode = cfg.get("auth", {}).get("mode", "")
    app_only = auth_mode == "client_secret"
    msg_client: GraphClient | None = None  # lazy-init delegated client

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

    # Ensure all configured users are team members before posting.
    # Uses the app-only client (Group.ReadWrite.All) — works regardless
    # of auth mode.
    all_upns = [
        u["upn"] for u in cfg.get("targets", {}).get("users", [])
    ]
    if all_upns and not client.dry_run:
        _ensure_team_members(client, team_id, all_upns)

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
            try:
                channel_id = _create_channel(client, team_id, display_name, description)
            except httpx.HTTPStatusError as exc:
                detail = ""
                if exc.response is not None:
                    try:
                        detail = (
                            exc.response.json().get("error", {}).get("message", "")
                        )
                    except Exception:
                        detail = exc.response.text or ""
                duplicate_name = (
                    exc.response is not None
                    and exc.response.status_code == 400
                    and "already existed" in detail
                )
                if duplicate_name:
                    logger.info(
                        "Channel '%s' name already exists — skipping create.",
                        display_name,
                    )
                    actions.append(
                        {
                            "action": "skip_channel",
                            "channel": display_name,
                            "reason": "already_exists",
                        }
                    )
                    continue
                logger.warning(
                    "Failed to create channel '%s': %s", display_name, exc,
                )
                actions.append({
                    "action": "error",
                    "channel": display_name,
                    "error": str(exc),
                })
                continue
            actions.append(
                {
                    "action": "create_channel",
                    "channel": display_name,
                    "channel_id": channel_id,
                    "api": "beta",
                }
            )

        # Post messages — requires delegated auth when running app-only
        posts = ch_cfg.get("posts", [])
        if posts:
            # Determine which client to use for message posting
            post_client = client
            if app_only:
                if msg_client is None:
                    logger.info(
                        "Channel message posting requires delegated auth — "
                        "initiating device-code sign-in …"
                    )
                    try:
                        msg_client = build_delegated_client(cfg, dry_run=client.dry_run)
                        # Warm the token so the device-code prompt appears now
                        msg_client._get_token()
                    except Exception as exc:
                        logger.warning(
                            "Could not obtain delegated credentials — "
                            "skipping message posting: %s", exc,
                        )
                        msg_client = False  # type: ignore[assignment]
                if msg_client is False:
                    actions.append({
                        "action": "skip_messages",
                        "channel": display_name,
                        "count": len(posts),
                        "reason": "delegated_auth_failed",
                    })
                else:
                    post_client = msg_client  # type: ignore[assignment]

            if post_client is not client and msg_client is False:
                pass  # already appended skip action above
            else:
                for post_cfg in posts:
                    # Accept both dict {"message": "..."} and plain string formats
                    message = (
                        post_cfg["message"] if isinstance(post_cfg, dict) else str(post_cfg)
                    )
                    logger.info(
                        "[BETA] Posting message to channel '%s'",
                        display_name,
                    )
                    try:
                        _post_message(post_client, team_id, channel_id, message, run_id)
                    except httpx.HTTPStatusError as exc:
                        logger.warning(
                            "Failed to post message to channel '%s': %s",
                            display_name,
                            exc,
                        )
                        actions.append({
                            "action": "error",
                            "channel": display_name,
                            "error": str(exc),
                        })
                        continue
                    actions.append(
                        {
                            "action": "post_message",
                            "channel": display_name,
                            "api": "beta",
                        }
                    )

    return actions
