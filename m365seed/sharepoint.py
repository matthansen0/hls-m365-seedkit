"""SharePoint seeding — create sites, pages, and document libraries.

Creates SharePoint communication sites (via M365 Groups), seeds site pages,
and uploads documents.  All content is tagged with the run_id for cleanup.
"""

from __future__ import annotations

import logging
from typing import Any

from m365seed.graph import GraphClient
from m365seed.theme_content import get_sharepoint_sites

logger = logging.getLogger("m365seed.sharepoint")

DISCLAIMER = "Demo content — synthetic, no patient data."


# ---------------------------------------------------------------------------
# Site creation (via Microsoft 365 Group)
# ---------------------------------------------------------------------------


def _group_exists(
    client: GraphClient,
    mail_nickname: str,
) -> dict[str, Any] | None:
    """Check if a group with the given mailNickname exists."""
    try:
        resp = client.get(
            "/groups",
            params={
                "$filter": f"mailNickname eq '{mail_nickname}'",
                "$select": "id,displayName,mailNickname",
                "$top": "1",
            },
        )
        groups = resp.json().get("value", [])
        return groups[0] if groups else None
    except Exception as exc:
        logger.warning("Group existence check failed for '%s': %s", mail_nickname, exc)
        return None


def _create_group_site(
    client: GraphClient,
    site_cfg: dict[str, Any],
    run_id: str,
    owner_id: str,
) -> dict[str, Any]:
    """Create a Microsoft 365 Group (which auto-provisions a SharePoint site).

    Returns the group object (contains id, displayName, etc.).
    """
    display_name = f"[DEMO-SEED:{run_id}] {site_cfg['display_name']}"
    mail_nickname = site_cfg.get(
        "mail_nickname",
        site_cfg["display_name"].replace(" ", "").lower()[:50],
    )
    # Prefix mail_nickname with run_id fragment for uniqueness
    mail_nickname = f"seed{run_id[:8]}{mail_nickname}"[:64]

    payload = {
        "displayName": display_name,
        "description": site_cfg.get(
            "description", f"Demo site for {site_cfg['display_name']}"
        ),
        "mailNickname": mail_nickname,
        "mailEnabled": False,
        "securityEnabled": False,
        "groupTypes": ["Unified"],
        "owners@odata.bind": [
            f"https://graph.microsoft.com/v1.0/users/{owner_id}"
        ],
        "members@odata.bind": [
            f"https://graph.microsoft.com/v1.0/users/{owner_id}"
        ],
        "visibility": "Private",
    }

    resp = client.post("/groups", json_body=payload)
    return resp.json()


def _get_group_site_id(client: GraphClient, group_id: str) -> str:
    """Get the SharePoint site ID associated with a group."""
    resp = client.get(f"/groups/{group_id}/sites/root", params={"$select": "id"})
    return resp.json().get("id", "")


# ---------------------------------------------------------------------------
# Site pages
# ---------------------------------------------------------------------------


def _page_exists(
    client: GraphClient,
    site_id: str,
    title: str,
) -> bool:
    """Check if a site page with the given title exists."""
    try:
        resp = client.get(
            f"/sites/{site_id}/pages",
            params={
                "$filter": f"title eq '{title}'",
                "$select": "id,title",
                "$top": "1",
            },
        )
        return len(resp.json().get("value", [])) > 0
    except Exception:
        return False


def _create_site_page(
    client: GraphClient,
    site_id: str,
    page_cfg: dict[str, Any],
    run_id: str,
) -> dict[str, Any]:
    """Create a SharePoint site page with HTML content."""
    title = f"[DEMO-SEED:{run_id}] {page_cfg['title']}"
    content = page_cfg.get("content", f"<p>{DISCLAIMER}</p>")

    payload = {
        "name": f"{page_cfg['title'].replace(' ', '_')}.aspx",
        "title": title,
        "pageLayout": "article",
        "showComments": True,
        "showRecommendedPages": False,
        "titleArea": {
            "enableGradientEffect": True,
            "layout": "plain",
            "showAuthor": True,
        },
        "canvasLayout": {
            "horizontalSections": [
                {
                    "layout": "fullWidth",
                    "columns": [
                        {
                            "width": 12,
                            "webparts": [
                                {
                                    "innerHtml": (
                                        f"{content}"
                                        f"<p><em>{DISCLAIMER} | RunId: {run_id}</em></p>"
                                    ),
                                }
                            ],
                        }
                    ],
                }
            ]
        },
    }

    resp = client.post(
        f"/sites/{site_id}/pages",
        json_body=payload,
    )
    return resp.json()


# ---------------------------------------------------------------------------
# Document upload to site drive
# ---------------------------------------------------------------------------


def _upload_to_site_drive(
    client: GraphClient,
    site_id: str,
    folder: str,
    filename: str,
    content: bytes,
    run_id: str,
) -> dict[str, Any]:
    """Upload a file to a SharePoint site's default document library."""
    prefixed = f"{run_id}_{filename}"
    path = f"{folder}/{prefixed}" if folder else prefixed

    resp = client.put(
        f"/sites/{site_id}/drive/root:/{path}:/content",
        content=content,
        headers={"Content-Type": "application/octet-stream"},
    )
    return resp.json()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def seed_sharepoint(
    client: GraphClient,
    cfg: dict[str, Any],
    theme: str,
    run_id: str,
) -> list[dict[str, Any]]:
    """Create SharePoint sites, pages, and upload documents.

    Returns a list of action records with site/group IDs for cleanup.
    """
    sp_cfg = cfg.get("sharepoint", {})
    if not sp_cfg.get("enabled"):
        logger.info("SharePoint seeding is disabled — skipping.")
        return []

    sites = sp_cfg.get("sites", [])
    if not sites:
        logger.info("No SharePoint sites configured — skipping.")
        return []

    # Enrich config sites with theme-specific pages and document content
    theme_sites = {
        s["display_name"]: s for s in get_sharepoint_sites(theme) if "display_name" in s
    }
    for site_cfg in sites:
        sname = site_cfg.get("display_name", "")
        if sname in theme_sites:
            ts = theme_sites[sname]
            if not site_cfg.get("pages") and ts.get("pages"):
                site_cfg["pages"] = ts["pages"]
            if not site_cfg.get("documents") and ts.get("documents"):
                site_cfg["documents"] = ts["documents"]
            if not site_cfg.get("description") and ts.get("description"):
                site_cfg["description"] = ts["description"]

    actions: list[dict[str, Any]] = []

    # Resolve owner ID
    owner_upn = sp_cfg.get(
        "owner", cfg["targets"]["users"][0]["upn"]
    )
    try:
        owner_resp = client.get(f"/users/{owner_upn}", params={"$select": "id"})
        owner_id = owner_resp.json().get("id", owner_upn)
    except Exception:
        owner_id = owner_upn

    for site_cfg in sites:
        site_name = site_cfg["display_name"]
        mail_nickname = site_cfg.get(
            "mail_nickname",
            site_name.replace(" ", "").lower()[:50],
        )
        prefixed_nickname = f"seed{run_id[:8]}{mail_nickname}"[:64]

        # Idempotency — check if group already exists
        existing = (
            None if client.dry_run else _group_exists(client, prefixed_nickname)
        )

        if existing:
            group_id = existing["id"]
            logger.info(
                "Group/site '%s' already exists (id=%s) — reusing.",
                site_name,
                group_id,
            )
            actions.append(
                {
                    "action": "skip_site",
                    "site": site_name,
                    "group_id": group_id,
                    "reason": "already_exists",
                }
            )
        else:
            logger.info("Creating SharePoint site '%s' via M365 Group …", site_name)
            try:
                group_data = _create_group_site(client, site_cfg, run_id, owner_id)
                group_id = group_data.get("id", "dry-run-id")
                actions.append(
                    {
                        "action": "create_site",
                        "site": site_name,
                        "group_id": group_id,
                    }
                )
            except Exception as exc:
                logger.error("Failed to create site '%s': %s", site_name, exc)
                actions.append(
                    {"action": "error", "site": site_name, "error": str(exc)}
                )
                continue

        # Get site ID for pages and document uploads
        site_id = ""
        if not client.dry_run:
            try:
                site_id = _get_group_site_id(client, group_id)
            except Exception as exc:
                logger.warning(
                    "Could not get site ID for group %s: %s", group_id, exc
                )

        # Create pages
        for page_cfg in site_cfg.get("pages", []):
            title = f"[DEMO-SEED:{run_id}] {page_cfg['title']}"
            if not client.dry_run and site_id and _page_exists(client, site_id, title):
                logger.info("Page '%s' already exists — skipping.", page_cfg["title"])
                actions.append(
                    {
                        "action": "skip_page",
                        "page": page_cfg["title"],
                        "reason": "already_exists",
                    }
                )
                continue

            logger.info("Creating page '%s' on site '%s'", page_cfg["title"], site_name)
            try:
                page_data = _create_site_page(
                    client, site_id or "dry-run-site", page_cfg, run_id
                )
                actions.append(
                    {
                        "action": "create_page",
                        "page": page_cfg["title"],
                        "page_id": page_data.get("id", ""),
                        "site": site_name,
                    }
                )
            except Exception as exc:
                logger.error("Failed to create page: %s", exc)

        # Upload documents to site drive
        for doc_cfg in site_cfg.get("documents", []):
            folder = doc_cfg.get("folder", "")
            filename = doc_cfg["filename"]
            content_text = doc_cfg.get(
                "content", f"Synthetic document — {filename}\n{DISCLAIMER}"
            )

            logger.info(
                "Uploading '%s' to site '%s' …", filename, site_name
            )
            try:
                upload_data = _upload_to_site_drive(
                    client,
                    site_id or "dry-run-site",
                    folder,
                    filename,
                    content_text.encode("utf-8"),
                    run_id,
                )
                actions.append(
                    {
                        "action": "upload_document",
                        "site": site_name,
                        "filename": f"{run_id}_{filename}",
                        "item_id": upload_data.get("id", ""),
                    }
                )
            except Exception as exc:
                logger.error("Failed to upload '%s': %s", filename, exc)

    return actions
