"""File seeding — upload synthetic theme-specific documents."""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import httpx
from jinja2 import Environment, FileSystemLoader

from m365seed.graph import GraphClient
from m365seed.theme_content import get_file_manifest

logger = logging.getLogger("m365seed.files")

TEMPLATES_DIR = Path(__file__).parent / "templates"
DATA_DIR = Path(__file__).parent / "data"

# ---------------------------------------------------------------------------
# File manifest — loaded from themes.json via theme_content module
# ---------------------------------------------------------------------------


def _jinja_env(theme: str) -> Environment:
    search_path = [
        str(TEMPLATES_DIR / theme),
        str(TEMPLATES_DIR / "healthcare"),
    ]
    return Environment(loader=FileSystemLoader(search_path), autoescape=False)


def _render_file(theme: str, template_name: str, run_id: str) -> str:
    """Render a document template to string."""
    env = _jinja_env(theme)
    try:
        tpl = env.get_template(template_name)
        return tpl.render(run_id=run_id, theme=theme)
    except Exception:
        return (
            f"[Synthetic document — {template_name}]\n"
            f"Run ID: {run_id}\n"
            "Demo content — synthetic, no patient data.\n"
        )


# ---------------------------------------------------------------------------
# OneDrive helpers
# ---------------------------------------------------------------------------


def _ensure_folder(client: GraphClient, user_upn: str, folder_path: str) -> None:
    """Create a folder in the user's OneDrive if it does not exist."""
    # Graph: PUT /users/{upn}/drive/root:/{path}  — creates if absent
    # We use a PATCH with folder facet; simpler to just try creating
    try:
        client.get(f"/users/{user_upn}/drive/root:/{folder_path}")
        logger.debug("Folder '%s' already exists for %s", folder_path, user_upn)
    except Exception:
        payload = {
            "name": folder_path.split("/")[-1],
            "folder": {},
            "@microsoft.graph.conflictBehavior": "fail",
        }
        parent = "/".join(folder_path.split("/")[:-1]) or ""
        path = (
            f"/users/{user_upn}/drive/root/children"
            if not parent
            else f"/users/{user_upn}/drive/root:/{parent}:/children"
        )
        try:
            client.post(path, json_body=payload)
            logger.info("Created folder '%s' for %s", folder_path, user_upn)
        except Exception as exc:
            logger.debug("Folder creation skipped (may exist): %s", exc)


def _file_exists(
    client: GraphClient, user_upn: str, folder: str, filename: str
) -> bool:
    """Check if a file already exists in the user's OneDrive."""
    file_path = f"{folder}/{filename}"
    try:
        resp = client.get(
            f"/users/{user_upn}/drive/root:/{file_path}",
            params={"$select": "id,name"},
        )
        return resp.status_code == 200
    except Exception:
        return False


def _upload_file(
    client: GraphClient,
    user_upn: str,
    folder: str,
    filename: str,
    content: str,
    run_id: str,
) -> dict[str, Any]:
    """Upload (or replace) a small text file to OneDrive."""
    file_path = f"{folder}/{filename}"
    resp = client.put(
        f"/users/{user_upn}/drive/root:/{file_path}:/content",
        content=content.encode("utf-8"),
        headers={
            "Content-Type": "text/plain",
        },
        params={
            "@microsoft.graph.conflictBehavior": "replace",
        },
    )
    return resp.json()


# ---------------------------------------------------------------------------
# SharePoint helpers
# ---------------------------------------------------------------------------


def _sp_file_exists(
    client: GraphClient, site_id: str, drive_id: str, folder: str, filename: str
) -> bool:
    """Check if a file already exists on a SharePoint drive."""
    file_path = f"{folder}/{filename}"
    try:
        resp = client.get(
            f"/sites/{site_id}/drives/{drive_id}/root:/{file_path}",
            params={"$select": "id,name"},
        )
        return resp.status_code == 200
    except Exception:
        return False


def _upload_sp_file(
    client: GraphClient,
    site_id: str,
    drive_id: str,
    folder: str,
    filename: str,
    content: str,
) -> dict[str, Any]:
    """Upload a small text file to a SharePoint document library."""
    file_path = f"{folder}/{filename}"
    resp = client.put(
        f"/sites/{site_id}/drives/{drive_id}/root:/{file_path}:/content",
        content=content.encode("utf-8"),
        headers={"Content-Type": "text/plain"},
    )
    return resp.json()


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def seed_files(
    client: GraphClient,
    cfg: dict[str, Any],
    theme: str,
    run_id: str,
) -> list[dict[str, Any]]:
    """Upload all configured files. Returns a list of action records."""
    files_cfg = cfg.get("files", {})
    actions: list[dict[str, Any]] = []

    # ── OneDrive ──────────────────────────────────────────────
    od_cfg = files_cfg.get("oneDrive", {})
    if od_cfg.get("enabled"):
        user_upn = od_cfg.get("target_user", cfg["targets"]["users"][0]["upn"])
        configured_folders = set(od_cfg.get("folders", []))

        file_manifest = get_file_manifest(theme)
        for folder, filename, template_name, desc in file_manifest:
            if configured_folders and folder not in configured_folders:
                continue

            # Deterministic filename with run_id
            det_filename = f"{run_id}_{filename}"

            # Idempotency: skip if exists
            if not client.dry_run and _file_exists(
                client, user_upn, folder, det_filename
            ):
                logger.info(
                    "File '%s/%s' already exists — skipping.",
                    folder,
                    det_filename,
                )
                actions.append(
                    {
                        "action": "skip",
                        "target": "onedrive",
                        "user": user_upn,
                        "path": f"{folder}/{det_filename}",
                        "reason": "already_exists",
                    }
                )
                continue

            _ensure_folder(client, user_upn, folder)

            content = _render_file(theme, template_name, run_id)
            try:
                _upload_file(client, user_upn, folder, det_filename, content, run_id)
            except httpx.HTTPStatusError as exc:
                logger.warning(
                    "Failed to upload '%s/%s' to %s: %s",
                    folder, det_filename, user_upn, exc,
                )
                actions.append({
                    "action": "error",
                    "target": "onedrive",
                    "user": user_upn,
                    "path": f"{folder}/{det_filename}",
                    "error": str(exc),
                })
                continue

            logger.info("Uploaded '%s/%s' to %s OneDrive", folder, det_filename, user_upn)
            actions.append(
                {
                    "action": "upload",
                    "target": "onedrive",
                    "user": user_upn,
                    "path": f"{folder}/{det_filename}",
                    "description": desc,
                }
            )

    # ── SharePoint ────────────────────────────────────────────
    sp_cfg = files_cfg.get("sharePoint", {})
    if sp_cfg.get("enabled"):
        site_id = sp_cfg["site_id"]
        drive_id = sp_cfg["drive_id"]

        file_manifest = get_file_manifest(theme)
        for folder, filename, template_name, desc in file_manifest:
            det_filename = f"{run_id}_{filename}"

            if not client.dry_run and _sp_file_exists(
                client, site_id, drive_id, folder, det_filename
            ):
                logger.info(
                    "SP file '%s/%s' already exists — skipping.",
                    folder,
                    det_filename,
                )
                actions.append(
                    {
                        "action": "skip",
                        "target": "sharepoint",
                        "path": f"{folder}/{det_filename}",
                        "reason": "already_exists",
                    }
                )
                continue

            content = _render_file(theme, template_name, run_id)
            try:
                _upload_sp_file(client, site_id, drive_id, folder, det_filename, content)
            except httpx.HTTPStatusError as exc:
                logger.warning(
                    "Failed to upload SP file '%s/%s': %s",
                    folder, det_filename, exc,
                )
                actions.append({
                    "action": "error",
                    "target": "sharepoint",
                    "path": f"{folder}/{det_filename}",
                    "error": str(exc),
                })
                continue

            logger.info(
                "Uploaded '%s/%s' to SharePoint site %s",
                folder,
                det_filename,
                site_id,
            )
            actions.append(
                {
                    "action": "upload",
                    "target": "sharepoint",
                    "path": f"{folder}/{det_filename}",
                    "description": desc,
                }
            )

    return actions
