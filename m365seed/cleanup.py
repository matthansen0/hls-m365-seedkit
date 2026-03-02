"""Cleanup mode — remove all seeded content tagged with a specific run_id."""

from __future__ import annotations

import logging
from typing import Any

from m365seed.graph import GraphClient, GRAPH_BETA

logger = logging.getLogger("m365seed.cleanup")


# ---------------------------------------------------------------------------
# Mail cleanup
# ---------------------------------------------------------------------------


def _cleanup_mail(
    client: GraphClient,
    cfg: dict[str, Any],
    run_id: str,
) -> list[dict[str, Any]]:
    """Delete all emails matching the DEMO-SEED run_id tag."""
    actions: list[dict[str, Any]] = []
    users = cfg["targets"]["users"]

    for user in users:
        upn = user["upn"]
        tag = f"DEMO-SEED:{run_id}"
        try:
            resp = client.get(
                f"/users/{upn}/messages",
                params={
                    "$search": f'"subject:{tag}"',
                    "$select": "id,subject",
                    "$top": "100",
                },
                headers={**client._auth_headers(), "ConsistencyLevel": "eventual"},
            )
            messages = resp.json().get("value", [])
        except Exception as exc:
            logger.warning("Failed to search mail for %s: %s", upn, exc)
            continue

        for msg in messages:
            msg_id = msg["id"]
            logger.info(
                "Deleting mail '%s' from %s",
                msg.get("subject", msg_id)[:60],
                upn,
            )
            try:
                client.delete(f"/users/{upn}/messages/{msg_id}")
                actions.append(
                    {
                        "action": "delete_mail",
                        "user": upn,
                        "message_id": msg_id,
                    }
                )
            except Exception as exc:
                logger.error("Failed to delete message %s: %s", msg_id, exc)

    return actions


# ---------------------------------------------------------------------------
# Files cleanup
# ---------------------------------------------------------------------------


def _cleanup_files(
    client: GraphClient,
    cfg: dict[str, Any],
    run_id: str,
) -> list[dict[str, Any]]:
    """Delete all files whose names start with the run_id prefix."""
    actions: list[dict[str, Any]] = []

    od_cfg = cfg.get("files", {}).get("oneDrive", {})
    if not od_cfg.get("enabled"):
        return actions

    user_upn = od_cfg.get("target_user", cfg["targets"]["users"][0]["upn"])
    folders = od_cfg.get("folders", [])

    for folder in folders:
        try:
            resp = client.get(
                f"/users/{user_upn}/drive/root:/{folder}:/children",
                params={"$select": "id,name"},
            )
            items = resp.json().get("value", [])
        except Exception as exc:
            logger.warning("Failed to list folder '%s': %s", folder, exc)
            continue

        for item in items:
            if item["name"].startswith(f"{run_id}_"):
                logger.info(
                    "Deleting file '%s/%s' from %s",
                    folder,
                    item["name"],
                    user_upn,
                )
                try:
                    client.delete(
                        f"/users/{user_upn}/drive/items/{item['id']}"
                    )
                    actions.append(
                        {
                            "action": "delete_file",
                            "user": user_upn,
                            "path": f"{folder}/{item['name']}",
                        }
                    )
                except Exception as exc:
                    logger.error("Failed to delete file %s: %s", item["id"], exc)

    return actions


# ---------------------------------------------------------------------------
# Calendar cleanup
# ---------------------------------------------------------------------------


def _cleanup_calendar(
    client: GraphClient,
    cfg: dict[str, Any],
    run_id: str,
) -> list[dict[str, Any]]:
    """Delete all calendar events matching the DEMO-SEED run_id tag."""
    actions: list[dict[str, Any]] = []
    users = cfg["targets"]["users"]

    for user in users:
        upn = user["upn"]
        tag = f"[DEMO-SEED:{run_id}:"
        try:
            resp = client.get(
                f"/users/{upn}/events",
                params={
                    "$filter": f"startsWith(subject, '{tag}')",
                    "$select": "id,subject",
                    "$top": "100",
                },
            )
            events = resp.json().get("value", [])
        except Exception as exc:
            logger.warning("Failed to list events for %s: %s", upn, exc)
            continue

        for event in events:
            event_id = event["id"]
            logger.info(
                "Deleting event '%s' from %s",
                event.get("subject", event_id)[:60],
                upn,
            )
            try:
                client.delete(f"/users/{upn}/events/{event_id}")
                actions.append(
                    {
                        "action": "delete_event",
                        "user": upn,
                        "event_id": event_id,
                    }
                )
            except Exception as exc:
                logger.error("Failed to delete event %s: %s", event_id, exc)

    return actions


# ---------------------------------------------------------------------------
# Teams channels cleanup
# ---------------------------------------------------------------------------


def _cleanup_teams(
    client: GraphClient,
    cfg: dict[str, Any],
    run_id: str,
) -> list[dict[str, Any]]:
    """Delete Teams channels created during seeding.

    Channels are identified by matching configured display_name values.
    Channel messages are deleted when the channel is deleted.
    Uses /beta for channel deletion.
    """
    actions: list[dict[str, Any]] = []
    teams_cfg = cfg.get("teams", {})

    if not teams_cfg.get("enabled"):
        return actions

    team_id = teams_cfg.get("team_id", "")
    if not team_id:
        return actions

    channels = teams_cfg.get("channels", [])
    if not channels:
        return actions

    # List all channels once and filter client-side
    # (Graph /teams/{id}/channels does not support $filter on displayName)
    try:
        resp = client.get(
            f"/teams/{team_id}/channels",
            params={"$select": "id,displayName"},
        )
        all_channels = resp.json().get("value", [])
    except Exception as exc:
        logger.warning("Failed to list channels for team %s: %s", team_id, exc)
        return actions

    # Build a lookup: lowercase display_name -> channel data
    channel_map: dict[str, dict[str, Any]] = {
        ch["displayName"].lower(): ch for ch in all_channels if ch.get("displayName")
    }

    for ch_cfg in channels:
        display_name = ch_cfg["display_name"]
        matches = [channel_map[display_name.lower()]] if display_name.lower() in channel_map else []

        for ch in matches:
            ch_id = ch["id"]
            logger.info("Deleting Teams channel '%s' …", display_name)
            try:
                client.delete(
                    f"/teams/{team_id}/channels/{ch_id}",
                    base=GRAPH_BETA,
                )
                actions.append(
                    {
                        "action": "delete_channel",
                        "channel": display_name,
                        "channel_id": ch_id,
                        "api": "beta",
                    }
                )
            except Exception as exc:
                logger.error("Failed to delete channel '%s': %s", display_name, exc)

    return actions


# ---------------------------------------------------------------------------
# Teams chats cleanup
# ---------------------------------------------------------------------------


def _cleanup_chats(
    client: GraphClient,
    cfg: dict[str, Any],
    run_id: str,
) -> list[dict[str, Any]]:
    """Delete Teams chats created during seeding.

    Group chats are identified by topic containing the DEMO-SEED tag.
    1:1 chats cannot be reliably filtered, so we only clean group chats
    with run_id in the topic.  Uses /beta.
    """
    actions: list[dict[str, Any]] = []
    chats_cfg = cfg.get("chats", {})

    if not chats_cfg.get("enabled"):
        return actions

    # Search for chats with our run_id tag in the topic
    tag = f"DEMO-SEED:{run_id}"
    try:
        # List chats and filter client-side (Graph /chats doesn't support $filter on topic)
        resp = client.get(
            "/chats",
            params={"$top": "50", "$select": "id,topic,chatType"},
            base=GRAPH_BETA,
        )
        all_chats = resp.json().get("value", [])
    except Exception as exc:
        exc_str = str(exc)
        if "403" in exc_str or "Forbidden" in exc_str:
            logger.warning(
                "Cannot list chats (403 Forbidden). "
                "Chat cleanup requires delegated (user) auth — "
                "it is not supported with app-only client_credentials. "
                "Skipping chat cleanup."
            )
        else:
            logger.warning("Failed to list chats: %s", exc)
        return actions

    for chat in all_chats:
        topic = chat.get("topic", "") or ""
        if tag in topic:
            chat_id = chat["id"]
            logger.info("Deleting chat '%s' …", topic[:60])
            try:
                client.delete(f"/chats/{chat_id}", base=GRAPH_BETA)
                actions.append(
                    {
                        "action": "delete_chat",
                        "chat_id": chat_id,
                        "topic": topic,
                        "api": "beta",
                    }
                )
            except Exception as exc:
                logger.error("Failed to delete chat %s: %s", chat_id, exc)

    return actions


# ---------------------------------------------------------------------------
# SharePoint cleanup (sites via group deletion, pages, documents)
# ---------------------------------------------------------------------------


def _cleanup_sharepoint(
    client: GraphClient,
    cfg: dict[str, Any],
    run_id: str,
) -> list[dict[str, Any]]:
    """Delete SharePoint sites (via M365 Group deletion), pages, and documents.

    Sites created by this tool are backed by M365 Groups whose displayName
    starts with ``[DEMO-SEED:<run_id>]``.  Deleting the group also deletes
    the associated SharePoint site, Planner plans, and other group resources.
    """
    actions: list[dict[str, Any]] = []

    tag = f"[DEMO-SEED:{run_id}]"

    try:
        resp = client.get(
            "/groups",
            params={
                "$filter": f"startsWith(displayName, '{tag}')",
                "$select": "id,displayName",
                "$top": "100",
            },
        )
        groups = resp.json().get("value", [])
    except Exception as exc:
        logger.warning("Failed to search for seeded groups: %s", exc)
        return actions

    for grp in groups:
        group_id = grp["id"]
        display_name = grp.get("displayName", group_id)

        # Before deleting group, clean up site pages tagged with run_id
        try:
            site_resp = client.get(
                f"/groups/{group_id}/sites/root",
                params={"$select": "id"},
            )
            site_id = site_resp.json().get("id", "")
            if site_id:
                _cleanup_site_pages(client, site_id, run_id, actions)
                _cleanup_site_documents(client, site_id, run_id, actions)
        except Exception:
            pass  # site may not exist yet

        logger.info("Deleting group/site '%s' …", display_name)
        try:
            client.delete(f"/groups/{group_id}")
            # Permanently delete from recycle bin to free the mailNickname
            try:
                client.delete(f"/directory/deletedItems/{group_id}")
            except Exception:
                pass  # may fail if not yet in deleted items; non-critical
            actions.append(
                {
                    "action": "delete_group_site",
                    "group_id": group_id,
                    "display_name": display_name,
                }
            )
        except Exception as exc:
            logger.error("Failed to delete group %s: %s", group_id, exc)

    # Also clean up documents on explicitly configured SharePoint sites
    sp_file_cfg = cfg.get("files", {}).get("sharePoint", {})
    if sp_file_cfg.get("enabled") and sp_file_cfg.get("site_id"):
        site_id = sp_file_cfg["site_id"]
        _cleanup_site_documents(client, site_id, run_id, actions)

    return actions


def _cleanup_site_pages(
    client: GraphClient,
    site_id: str,
    run_id: str,
    actions: list[dict[str, Any]],
) -> None:
    """Delete site pages whose title contains the DEMO-SEED tag."""
    tag = f"[DEMO-SEED:{run_id}]"
    try:
        resp = client.get(
            f"/sites/{site_id}/pages",
            params={
                "$filter": f"startsWith(title, '{tag}')",
                "$select": "id,title",
                "$top": "100",
            },
        )
        pages = resp.json().get("value", [])
    except Exception as exc:
        logger.warning("Failed to list site pages: %s", exc)
        return

    for page in pages:
        page_id = page["id"]
        logger.info("Deleting page '%s' …", page.get("title", page_id)[:60])
        try:
            client.delete(f"/sites/{site_id}/pages/{page_id}")
            actions.append(
                {
                    "action": "delete_page",
                    "page_id": page_id,
                    "title": page.get("title", ""),
                }
            )
        except Exception as exc:
            logger.error("Failed to delete page %s: %s", page_id, exc)


def _cleanup_site_documents(
    client: GraphClient,
    site_id: str,
    run_id: str,
    actions: list[dict[str, Any]],
) -> None:
    """Delete documents from a site's drive that are prefixed with run_id."""
    try:
        resp = client.get(
            f"/sites/{site_id}/drive/root/children",
            params={"$select": "id,name,folder"},
        )
        items = resp.json().get("value", [])
    except Exception as exc:
        logger.warning("Failed to list site drive: %s", exc)
        return

    for item in items:
        name = item.get("name", "")
        if name.startswith(f"{run_id}_"):
            item_id = item["id"]
            logger.info("Deleting site document '%s' …", name)
            try:
                client.delete(f"/sites/{site_id}/drive/items/{item_id}")
                actions.append(
                    {
                        "action": "delete_site_document",
                        "item_id": item_id,
                        "filename": name,
                    }
                )
            except Exception as exc:
                logger.error("Failed to delete document %s: %s", item_id, exc)

        # Check subfolders
        if "folder" in item:
            try:
                sub_resp = client.get(
                    f"/sites/{site_id}/drive/items/{item['id']}/children",
                    params={"$select": "id,name"},
                )
                for sub in sub_resp.json().get("value", []):
                    if sub["name"].startswith(f"{run_id}_"):
                        logger.info("Deleting site document '%s/%s' …", name, sub["name"])
                        try:
                            client.delete(f"/sites/{site_id}/drive/items/{sub['id']}")
                            actions.append(
                                {
                                    "action": "delete_site_document",
                                    "item_id": sub["id"],
                                    "filename": f"{name}/{sub['name']}",
                                }
                            )
                        except Exception as exc:
                            logger.error("Failed to delete %s: %s", sub["id"], exc)
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Teams / Planner group cleanup
# ---------------------------------------------------------------------------


def _cleanup_team_group(
    client: GraphClient,
    cfg: dict[str, Any],
    run_id: str,
) -> list[dict[str, Any]]:
    """Delete the M365 Group backing Teams and/or Planner if it was created
    by the seeding tool.

    The group is identified from ``teams.team_id`` or ``planner.group_id``
    in the config.  We only delete groups whose description starts with
    ``"Demo team for"`` (the marker set by ``_create_team_group`` in setup)
    to avoid accidentally deleting pre-existing tenant groups.
    """
    actions: list[dict[str, Any]] = []

    # Collect candidate group IDs from config
    group_ids: set[str] = set()
    teams_cfg = cfg.get("teams", {})
    planner_cfg = cfg.get("planner", {})
    if teams_cfg.get("enabled") and teams_cfg.get("team_id"):
        group_ids.add(teams_cfg["team_id"])
    if planner_cfg.get("enabled") and planner_cfg.get("group_id"):
        group_ids.add(planner_cfg["group_id"])

    if not group_ids:
        return actions

    for gid in group_ids:
        # Fetch group details to verify it's one we created
        try:
            resp = client.get(
                f"/groups/{gid}",
                params={"$select": "id,displayName,description"},
            )
            group = resp.json()
        except Exception as exc:
            logger.warning("Failed to fetch group %s: %s", gid, exc)
            continue

        description = group.get("description", "") or ""
        display_name = group.get("displayName", gid)

        if not description.startswith("Demo team for"):
            logger.info(
                "Skipping group '%s' (%s) — not created by seeding tool.",
                display_name,
                gid[:8],
            )
            continue

        logger.info("Deleting group '%s' (%s) …", display_name, gid[:8])
        try:
            client.delete(f"/groups/{gid}")
            # Permanently delete from recycle bin to free the mailNickname
            try:
                client.delete(f"/directory/deletedItems/{gid}")
            except Exception:
                pass  # non-critical
            actions.append(
                {
                    "action": "delete_team_group",
                    "group_id": gid,
                    "display_name": display_name,
                }
            )
        except Exception as exc:
            logger.error("Failed to delete group %s: %s", gid, exc)

    return actions


# ---------------------------------------------------------------------------
# Planner cleanup
# ---------------------------------------------------------------------------


def _cleanup_planner(
    client: GraphClient,
    cfg: dict[str, Any],
    run_id: str,
) -> list[dict[str, Any]]:
    """Delete Planner plans (and their tasks/buckets) tagged with run_id.

    Plans are identified by title starting with ``[DEMO-SEED:<run_id>]``.
    Deleting a plan cascades to its buckets and tasks.
    """
    actions: list[dict[str, Any]] = []
    planner_cfg = cfg.get("planner", {})

    if not planner_cfg.get("enabled"):
        return actions

    group_id = planner_cfg.get("group_id", "")
    if not group_id:
        return actions

    tag = f"[DEMO-SEED:{run_id}]"

    try:
        resp = client.get(
            f"/groups/{group_id}/planner/plans",
            params={"$select": "id,title", "$top": "50"},
        )
        plans = resp.json().get("value", [])
    except Exception as exc:
        logger.warning("Failed to list plans for group %s: %s", group_id, exc)
        return actions

    for plan in plans:
        title = plan.get("title", "")
        if title.startswith(tag):
            plan_id = plan["id"]
            logger.info("Deleting Planner plan '%s' …", title[:60])
            try:
                # Planner DELETE requires If-Match with etag
                detail_resp = client.get(f"/planner/plans/{plan_id}")
                etag = detail_resp.headers.get("ETag", "*")

                client.delete(
                    f"/planner/plans/{plan_id}",
                    headers={"If-Match": etag},
                )
                actions.append(
                    {
                        "action": "delete_plan",
                        "plan_id": plan_id,
                        "title": title,
                    }
                )
            except Exception as exc:
                logger.error("Failed to delete plan %s: %s", plan_id, exc)

    return actions


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def cleanup(
    client: GraphClient,
    cfg: dict[str, Any],
    run_id: str,
    *,
    mail: bool = True,
    files: bool = True,
    calendar: bool = True,
    teams: bool = True,
    chats: bool = True,
    sharepoint: bool = True,
    planner: bool = True,
    team_group: bool = True,
) -> list[dict[str, Any]]:
    """Remove all seeded content tagged with *run_id*.

    Returns a combined list of action records across all content types.
    Each content type can be individually toggled.
    """
    actions: list[dict[str, Any]] = []

    if mail:
        logger.info("Cleaning up seeded mail for run_id=%s …", run_id)
        actions.extend(_cleanup_mail(client, cfg, run_id))

    if files:
        logger.info("Cleaning up seeded files for run_id=%s …", run_id)
        actions.extend(_cleanup_files(client, cfg, run_id))

    if calendar:
        logger.info("Cleaning up seeded calendar events for run_id=%s …", run_id)
        actions.extend(_cleanup_calendar(client, cfg, run_id))

    if teams:
        logger.info("Cleaning up seeded Teams channels for run_id=%s …", run_id)
        actions.extend(_cleanup_teams(client, cfg, run_id))

    if chats:
        logger.info("Cleaning up seeded Teams chats for run_id=%s …", run_id)
        actions.extend(_cleanup_chats(client, cfg, run_id))

    if sharepoint:
        logger.info("Cleaning up seeded SharePoint sites for run_id=%s …", run_id)
        actions.extend(_cleanup_sharepoint(client, cfg, run_id))

    if planner:
        logger.info("Cleaning up seeded Planner plans for run_id=%s …", run_id)
        actions.extend(_cleanup_planner(client, cfg, run_id))

    if team_group:
        logger.info("Cleaning up Teams/Planner group for run_id=%s …", run_id)
        actions.extend(_cleanup_team_group(client, cfg, run_id))

    logger.info("Cleanup complete: %d actions performed.", len(actions))
    return actions
