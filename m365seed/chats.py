"""Teams chat seeding — create 1:1 and group chats with messages.

⚠️  This module uses Microsoft Graph **beta** endpoints and is **off by
default**.  Enable with ``--enable-beta-teams``.  Chat creation with
application permissions requires ``Chat.Create`` and message posting
requires ``ChatMessage.Send`` (delegate) or migration endpoints.

For demo tenants this typically uses application permissions with
``Chat.ReadWrite.All`` on /beta.
"""

from __future__ import annotations

import logging
from typing import Any

from m365seed.graph import GraphClient, GRAPH_BETA
from m365seed.theme_content import get_chat_conversations

logger = logging.getLogger("m365seed.chats")

DISCLAIMER = "Demo content — synthetic, no patient data."


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _resolve_user_id(client: GraphClient, upn: str) -> str:
    """Resolve a UPN to a directory object id (needed for chat members)."""
    resp = client.get(f"/users/{upn}", params={"$select": "id"})
    return resp.json().get("id", upn)


def _create_chat(
    client: GraphClient,
    chat_cfg: dict[str, Any],
    run_id: str,
    user_ids: dict[str, str],
) -> dict[str, Any]:
    """Create a 1:1 or group chat.  Uses **/beta** with app permissions."""
    members = chat_cfg["members"]
    chat_type = chat_cfg.get("type", "group" if len(members) > 2 else "oneOnOne")
    topic = chat_cfg.get("topic", "")

    member_payload = []
    for upn in members:
        uid = user_ids.get(upn, upn)
        member_payload.append(
            {
                "@odata.type": "#microsoft.graph.aadUserConversationMember",
                "roles": ["owner"],
                "user@odata.bind": f"https://graph.microsoft.com/v1.0/users('{uid}')",
            }
        )

    payload: dict[str, Any] = {
        "chatType": chat_type,
        "members": member_payload,
    }
    if topic and chat_type == "group":
        payload["topic"] = f"[DEMO-SEED:{run_id}] {topic}"

    resp = client.post("/chats", json_body=payload, base=GRAPH_BETA)
    return resp.json()


def _send_chat_message(
    client: GraphClient,
    chat_id: str,
    message_text: str,
    run_id: str,
) -> None:
    """Post a message into a chat.  Uses **/beta**."""
    payload = {
        "body": {
            "contentType": "html",
            "content": (
                f"<p>{message_text}</p>"
                f"<p><em>{DISCLAIMER} | RunId: {run_id}</em></p>"
            ),
        },
    }
    client.post(
        f"/chats/{chat_id}/messages",
        json_body=payload,
        base=GRAPH_BETA,
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def seed_chats(
    client: GraphClient,
    cfg: dict[str, Any],
    theme: str,
    run_id: str,
) -> list[dict[str, Any]]:
    """Seed Teams 1:1 and group chats with messages.

    Gated behind ``chats.enabled`` config flag **and** the CLI
    ``--enable-beta-teams`` argument.

    Returns a list of action records (including chat IDs for cleanup).
    """
    chats_cfg = cfg.get("chats", {})
    if not chats_cfg.get("enabled"):
        logger.info("Teams chat seeding is disabled — skipping.")
        return []

    conversations = chats_cfg.get("conversations", [])
    if not conversations:
        logger.info("No chat conversations configured — skipping.")
        return []

    # Enrich config conversations with theme-specific messages
    theme_convs = {
        c["conversation_id"]: c
        for c in get_chat_conversations(theme)
        if "conversation_id" in c
    }
    for conv in conversations:
        cid = conv.get("conversation_id", "")
        if cid in theme_convs:
            tc = theme_convs[cid]
            if not conv.get("messages") and tc.get("messages"):
                conv["messages"] = tc["messages"]
            if not conv.get("topic") and tc.get("topic"):
                conv["topic"] = tc["topic"]

    actions: list[dict[str, Any]] = []

    # Pre-resolve user IDs (needed for member binding)
    all_upns: set[str] = set()
    for conv in conversations:
        all_upns.update(conv["members"])

    user_ids: dict[str, str] = {}
    for upn in all_upns:
        try:
            user_ids[upn] = _resolve_user_id(client, upn)
        except Exception as exc:
            logger.warning("Could not resolve user id for %s: %s", upn, exc)
            user_ids[upn] = upn  # fallback to UPN

    for conv in conversations:
        conv_id = conv.get("conversation_id", conv.get("topic", "unnamed"))
        logger.info(
            "[BETA] Creating chat '%s' with %d members",
            conv_id,
            len(conv["members"]),
        )

        try:
            chat_data = _create_chat(client, conv, run_id, user_ids)
            chat_id = chat_data.get("id", "dry-run-id")

            actions.append(
                {
                    "action": "create_chat",
                    "conversation_id": conv_id,
                    "chat_id": chat_id,
                    "type": conv.get("type", "group"),
                    "api": "beta",
                }
            )

            # Send messages
            for msg_cfg in conv.get("messages", []):
                message = msg_cfg["text"]
                logger.info(
                    "[BETA] Sending chat message in '%s'",
                    conv_id,
                )
                _send_chat_message(client, chat_id, message, run_id)
                actions.append(
                    {
                        "action": "send_chat_message",
                        "conversation_id": conv_id,
                        "chat_id": chat_id,
                        "api": "beta",
                    }
                )
        except Exception as exc:
            logger.error("Failed to create chat '%s': %s", conv_id, exc)
            actions.append(
                {
                    "action": "error",
                    "conversation_id": conv_id,
                    "error": str(exc),
                }
            )

    return actions
