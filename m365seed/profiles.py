"""User-profile branding — update M365 user attributes to match the HLS theme.

Uses ``PATCH /users/{upn}`` to set jobTitle, department, companyName,
officeLocation, and (if possible) aboutMe so that the demo tenant users
look like they actually belong to the themed organization.

All content is synthetic — no PHI, no PII.

Required Graph permissions: ``User.ReadWrite.All`` (application).
"""

from __future__ import annotations

import logging
from typing import Any

import httpx

from m365seed.config import get_users, get_theme
from m365seed.graph import GraphClient
from m365seed.theme_content import get_user_profiles

logger = logging.getLogger("m365seed.profiles")

# Graph-patchable user profile fields (v1.0)
PROFILE_FIELDS = ("jobTitle", "department", "companyName", "officeLocation")


# ---------------------------------------------------------------------------
# Profile resolution
# ---------------------------------------------------------------------------


def _build_profile_map(theme: str) -> dict[str, dict[str, str]]:
    """Return a mapping of role → profile attributes from theme data."""
    profiles = get_user_profiles(theme)
    return {p["role"]: p for p in profiles}


def resolve_profile(
    user: dict[str, str],
    profile_map: dict[str, dict[str, str]],
) -> dict[str, str] | None:
    """Resolve a user's profile patch payload from their configured role.

    Returns ``None`` if no matching profile is found for the user's role.
    """
    role = user.get("role", "")
    profile = profile_map.get(role)
    if not profile:
        logger.warning(
            "No profile found for role '%s' (user %s) — skipping.",
            role,
            user.get("upn", "?"),
        )
        return None

    # Build the PATCH payload with only the fields Graph v1.0 accepts
    payload: dict[str, str] = {}
    for field in PROFILE_FIELDS:
        if field in profile:
            payload[field] = profile[field]
    return payload


# ---------------------------------------------------------------------------
# Seeder entry point
# ---------------------------------------------------------------------------


def seed_profiles(
    client: GraphClient,
    cfg: dict[str, Any],
    theme: str,
    run_id: str,
) -> list[dict[str, Any]]:
    """Update target users' M365 profiles to match the theme.

    Returns a list of action dicts for the CLI actions table.
    """
    users = get_users(cfg)
    profile_map = _build_profile_map(theme)
    actions: list[dict[str, Any]] = []

    for user in users:
        upn = user["upn"]
        payload = resolve_profile(user, profile_map)

        if payload is None:
            actions.append({
                "action": "skip-profile",
                "upn": upn,
                "reason": f"No profile for role '{user.get('role', '')}'",
            })
            continue

        logger.info(
            "Updating profile for %s → %s (%s)",
            upn,
            payload.get("jobTitle", "?"),
            payload.get("department", "?"),
        )

        try:
            client.patch(f"/users/{upn}", json_body=payload)
        except httpx.HTTPStatusError as exc:
            logger.warning(
                "Failed to update profile for %s: %s", upn, exc,
            )
            actions.append({
                "action": "error-profile",
                "upn": upn,
                "error": str(exc),
            })
            continue

        actions.append({
            "action": "update-profile",
            "upn": upn,
            "jobTitle": payload.get("jobTitle", ""),
            "department": payload.get("department", ""),
            "companyName": payload.get("companyName", ""),
            "officeLocation": payload.get("officeLocation", ""),
        })

    logger.info(
        "Profile branding complete: %d users updated.",
        sum(1 for a in actions if a["action"] == "update-profile"),
    )
    return actions
