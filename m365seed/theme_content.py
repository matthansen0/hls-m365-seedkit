"""Theme content provider — typed access to per-theme content from themes.json.

Centralises access to theme-specific content for all seeding modules.
Each module can call ``get_*`` functions to retrieve contextually rich,
industry-specific content that falls back to sensible defaults.

All content is synthetic — no PHI, no PII.
"""

from __future__ import annotations

import json
import logging
from functools import lru_cache
from pathlib import Path
from typing import Any

logger = logging.getLogger("m365seed.theme_content")

DATA_DIR = Path(__file__).parent / "data"
THEMES_FILE = DATA_DIR / "themes.json"

VALID_THEMES = frozenset({"healthcare", "pharma", "medtech", "payor"})
_DEFAULT_THEME = "healthcare"


# ---------------------------------------------------------------------------
# Core loader
# ---------------------------------------------------------------------------


@lru_cache(maxsize=1)
def _load_all_themes() -> dict[str, Any]:
    """Load and cache the full themes.json file."""
    with open(THEMES_FILE, encoding="utf-8") as fh:
        data = json.load(fh)
    logger.debug("Loaded themes.json with keys: %s", list(data.keys()))
    return data


def load_theme(theme: str) -> dict[str, Any]:
    """Return the full content dict for the given theme.

    Falls back to ``healthcare`` if the requested theme is not found.
    """
    themes = _load_all_themes()
    if theme not in themes:
        logger.warning(
            "Theme '%s' not found in themes.json — falling back to '%s'.",
            theme,
            _DEFAULT_THEME,
        )
        theme = _DEFAULT_THEME
    return themes[theme]


def _get_section(theme: str, section: str) -> list[dict[str, Any]] | list[Any]:
    """Retrieve a list-valued section from theme content."""
    data = load_theme(theme)
    result = data.get(section, [])
    if not result:
        logger.debug(
            "Section '%s' empty for theme '%s' — trying fallback.",
            section,
            theme,
        )
        result = load_theme(_DEFAULT_THEME).get(section, [])
    return result


# ---------------------------------------------------------------------------
# File manifest
# ---------------------------------------------------------------------------


def get_file_manifest(
    theme: str,
) -> list[tuple[str, str, str, str]]:
    """Return the file manifest as a list of (folder, filename, template, desc).

    Each module entry in themes.json has keys:
    ``folder``, ``filename``, ``template``, ``description``.
    """
    raw = _get_section(theme, "file_manifest")
    return [
        (
            entry["folder"],
            entry["filename"],
            entry["template"],
            entry["description"],
        )
        for entry in raw
    ]


# ---------------------------------------------------------------------------
# Mail threads
# ---------------------------------------------------------------------------


def get_mail_threads(theme: str) -> list[dict[str, Any]]:
    """Return mail thread definitions for the given theme.

    Each entry has: ``thread_id``, ``subject``, ``attachment_name``,
    ``attachment_content``.
    """
    return _get_section(theme, "mail_threads")


# ---------------------------------------------------------------------------
# Calendar events
# ---------------------------------------------------------------------------


def get_calendar_events(theme: str) -> list[dict[str, Any]]:
    """Return calendar event definitions for the given theme.

    Each entry has: ``subject``, ``body``, ``duration_minutes``,
    ``recurrence`` (optional).
    """
    return _get_section(theme, "calendar_events")


# ---------------------------------------------------------------------------
# Teams channels
# ---------------------------------------------------------------------------


def get_teams_channels(theme: str) -> list[dict[str, Any]]:
    """Return Teams channel definitions for the given theme.

    Each entry has: ``display_name``, ``description``, ``posts``
    (list of dicts with ``message``).
    """
    return _get_section(theme, "teams_channels")


# ---------------------------------------------------------------------------
# Chat conversations
# ---------------------------------------------------------------------------


def get_chat_conversations(theme: str) -> list[dict[str, Any]]:
    """Return chat conversation definitions for the given theme.

    Each entry has: ``chat_type``, ``topic``, ``messages``
    (list of dicts with ``text``).
    """
    return _get_section(theme, "chat_conversations")


# ---------------------------------------------------------------------------
# SharePoint sites
# ---------------------------------------------------------------------------


def get_sharepoint_sites(theme: str) -> list[dict[str, Any]]:
    """Return SharePoint site definitions for the given theme.

    Each entry has: ``display_name``, ``description``, ``mail_nickname``,
    ``pages`` (list), ``documents`` (list).
    """
    return _get_section(theme, "sharepoint_sites")


# ---------------------------------------------------------------------------
# Planner plans
# ---------------------------------------------------------------------------


def get_planner_plans(theme: str) -> list[dict[str, Any]]:
    """Return Planner plan definitions for the given theme.

    Each entry has: ``title``, ``buckets`` (list of dicts with ``name``
    and ``tasks`` list).
    """
    return _get_section(theme, "planner_plans")


# ---------------------------------------------------------------------------
# Misc theme metadata
# ---------------------------------------------------------------------------


def get_organization(theme: str) -> str:
    """Return the synthetic organization name for the theme."""
    return load_theme(theme).get("organization", "Contoso Health")


def get_roles(theme: str) -> list[str]:
    """Return the list of role titles for the theme."""
    return load_theme(theme).get("roles", [])


def get_folders(theme: str) -> list[str]:
    """Return the list of OneDrive folder names for the theme."""
    return load_theme(theme).get("folders", [])


def get_industry_context(theme: str) -> str:
    """Return the industry context blurb for the theme."""
    return load_theme(theme).get("industry_context", "Healthcare operations")
