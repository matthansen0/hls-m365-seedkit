"""Configuration loading, validation, and schema definition."""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Any

import jsonschema
import yaml

# ---------------------------------------------------------------------------
# JSON-Schema for seed-config.yaml
# ---------------------------------------------------------------------------

CONFIG_SCHEMA: dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "type": "object",
    "required": ["tenant", "auth", "targets", "content"],
    "properties": {
        "tenant": {
            "type": "object",
            "required": ["tenant_id"],
            "properties": {
                "tenant_id": {"type": "string", "minLength": 1},
                "authority": {"type": "string"},
            },
        },
        "auth": {
            "type": "object",
            "required": ["mode", "client_id"],
            "properties": {
                "mode": {"type": "string", "enum": ["client_secret", "device_code"]},
                "client_id": {"type": "string", "minLength": 1},
                "client_secret_env": {"type": "string"},
            },
        },
        "targets": {
            "type": "object",
            "required": ["users"],
            "properties": {
                "users": {
                    "type": "array",
                    "minItems": 1,
                    "items": {
                        "type": "object",
                        "required": ["upn"],
                        "properties": {
                            "upn": {"type": "string", "minLength": 1},
                            "role": {"type": "string"},
                        },
                    },
                },
            },
        },
        "content": {
            "type": "object",
            "required": ["theme", "run_id"],
            "properties": {
                "theme": {
                    "type": "string",
                    "enum": ["healthcare", "pharma", "medtech", "payor"],
                },
                "run_id": {"type": "string", "minLength": 1},
            },
        },
        "mail": {
            "type": "object",
            "properties": {
                "threads": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["thread_id", "subject", "participants", "messages"],
                        "properties": {
                            "thread_id": {"type": "string", "minLength": 1},
                            "subject": {"type": "string"},
                            "participants": {
                                "type": "array",
                                "items": {"type": "string"},
                                "minItems": 1,
                            },
                            "messages": {"type": "integer", "minimum": 1},
                            "include_attachments": {"type": "boolean"},
                        },
                    },
                },
            },
        },
        "files": {
            "type": "object",
            "properties": {
                "oneDrive": {
                    "type": "object",
                    "properties": {
                        "enabled": {"type": "boolean"},
                        "target_user": {"type": "string"},
                        "folders": {
                            "type": "array",
                            "items": {"type": "string"},
                        },
                    },
                },
                "sharePoint": {
                    "type": "object",
                    "properties": {
                        "enabled": {"type": "boolean"},
                        "site_id": {"type": "string"},
                        "drive_id": {"type": "string"},
                    },
                },
            },
        },
        "calendar": {
            "type": "object",
            "properties": {
                "enabled": {"type": "boolean"},
                "events": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["event_id", "subject", "organizer"],
                        "properties": {
                            "event_id": {"type": "string", "minLength": 1},
                            "subject": {"type": "string"},
                            "organizer": {"type": "string"},
                            "attendees": {
                                "type": "array",
                                "items": {"type": "string"},
                            },
                            "recurrence": {"type": "string"},
                            "duration_minutes": {"type": "integer", "minimum": 5},
                            "is_online_meeting": {"type": "boolean"},
                        },
                    },
                },
            },
        },
        "teams": {
            "type": "object",
            "properties": {
                "enabled": {"type": "boolean"},
                "team_id": {"type": "string"},
                "channels": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "channel_id": {"type": "string"},
                            "display_name": {"type": "string"},
                            "description": {"type": "string"},
                            "posts": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "properties": {
                                        "message": {"type": "string"},
                                    },
                                },
                            },
                        },
                    },
                },
            },
        },
        "chats": {
            "type": "object",
            "properties": {
                "enabled": {"type": "boolean"},
                "conversations": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["conversation_id", "members"],
                        "properties": {
                            "conversation_id": {"type": "string", "minLength": 1},
                            "type": {
                                "type": "string",
                                "enum": ["oneOnOne", "group"],
                            },
                            "topic": {"type": "string"},
                            "members": {
                                "type": "array",
                                "items": {"type": "string"},
                                "minItems": 2,
                            },
                            "messages": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "required": ["sender", "text"],
                                    "properties": {
                                        "sender": {"type": "string"},
                                        "text": {"type": "string"},
                                    },
                                },
                            },
                        },
                    },
                },
            },
        },
        "sharepoint": {
            "type": "object",
            "properties": {
                "enabled": {"type": "boolean"},
                "owner": {"type": "string"},
                "sites": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["display_name"],
                        "properties": {
                            "display_name": {"type": "string", "minLength": 1},
                            "mail_nickname": {"type": "string"},
                            "description": {"type": "string"},
                            "pages": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "required": ["title"],
                                    "properties": {
                                        "title": {"type": "string"},
                                        "content": {"type": "string"},
                                    },
                                },
                            },
                            "documents": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "required": ["filename"],
                                    "properties": {
                                        "filename": {"type": "string"},
                                        "folder": {"type": "string"},
                                        "content": {"type": "string"},
                                    },
                                },
                            },
                        },
                    },
                },
            },
        },
        "planner": {
            "type": "object",
            "properties": {
                "enabled": {"type": "boolean"},
                "group_id": {"type": "string"},
                "plans": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "required": ["title"],
                        "properties": {
                            "title": {"type": "string", "minLength": 1},
                            "buckets": {
                                "type": "array",
                                "items": {
                                    "type": "object",
                                    "required": ["name"],
                                    "properties": {
                                        "name": {"type": "string"},
                                        "tasks": {
                                            "type": "array",
                                            "items": {
                                                "type": "object",
                                                "required": ["title"],
                                                "properties": {
                                                    "title": {"type": "string"},
                                                    "priority": {"type": "integer"},
                                                    "percent_complete": {
                                                        "type": "integer",
                                                        "enum": [0, 50, 100],
                                                    },
                                                    "assignees": {
                                                        "type": "array",
                                                        "items": {"type": "string"},
                                                    },
                                                },
                                            },
                                        },
                                    },
                                },
                            },
                        },
                    },
                },
            },
        },
    },
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def load_config(path: str | Path) -> dict[str, Any]:
    """Load and validate a seed-config YAML file.

    Returns the validated config dict.
    Raises ``jsonschema.ValidationError`` or ``FileNotFoundError``.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"Config file not found: {path}")

    with open(path, encoding="utf-8") as fh:
        raw = yaml.safe_load(fh)

    if not isinstance(raw, dict):
        raise ValueError("Config file must contain a YAML mapping at the top level.")

    validate_config(raw)
    return raw


def validate_config(cfg: dict[str, Any]) -> None:
    """Validate *cfg* against ``CONFIG_SCHEMA``.

    Raises ``jsonschema.ValidationError`` on the first error found.
    """
    jsonschema.validate(instance=cfg, schema=CONFIG_SCHEMA)


def resolve_secret(cfg: dict[str, Any]) -> str:
    """Return the client secret by reading the env var named in config.

    Raises ``RuntimeError`` if the env var is unset or empty.
    """
    env_name = cfg["auth"].get("client_secret_env", "M365SEED_CLIENT_SECRET")
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", env_name):
        raise RuntimeError(
            "auth.client_secret_env must be an environment variable name "
            "(for example: M365SEED_CLIENT_SECRET), not the secret value itself. "
            f"Got: '{env_name}'"
        )
    value = os.environ.get(env_name)
    if not value:
        raise RuntimeError(
            f"Environment variable '{env_name}' is not set or empty. "
            "Set it to the Entra app client secret."
        )
    return value


def get_run_id(cfg: dict[str, Any]) -> str:
    """Return the deterministic run identifier from config."""
    return cfg["content"]["run_id"]


def get_theme(cfg: dict[str, Any]) -> str:
    """Return the content theme (healthcare | pharma | medtech | payor)."""
    return cfg["content"].get("theme", "healthcare")


def get_users(cfg: dict[str, Any]) -> list[dict[str, str]]:
    """Return the list of target user dicts (upn, role)."""
    return cfg["targets"]["users"]
