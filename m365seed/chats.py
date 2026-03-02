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

from m365seed.graph import GraphClient, GRAPH_BETA, build_delegated_client
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
        # Teams chat topics cannot contain ':' characters
        safe_topic = topic.replace(":", "-")
        payload["topic"] = f"[DEMO-SEED-{run_id}] {safe_topic}"

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

    auth_mode = cfg.get("auth", {}).get("mode", "")
    chat_client = client

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

    # Pre-resolve user IDs using the *app* client (which has User.Read.All).
    # The delegated chat_client may not have that permission.
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

    # For delegated flows, Graph requires the signed-in caller to be a member
    # of chats they create.
    if auth_mode == "device_code":
        try:
            me_resp = chat_client.get("/me", params={"$select": "id,userPrincipalName"})
            me_data = me_resp.json()
            me_id = me_data.get("id", "")
            me_upn = me_data.get("userPrincipalName", "")
            if me_id and me_upn:
                user_ids[me_upn] = me_id
                for conv in conversations:
                    conv_type = conv.get("type", "group")
                    if conv_type == "oneOnOne":
                        continue
                    if me_upn not in conv["members"]:
                        conv["members"].append(me_upn)
        except Exception as exc:
            logger.warning("Could not resolve signed-in delegated user: %s", exc)

    for conv in conversations:
        conv_id = conv.get("conversation_id", conv.get("topic", "unnamed"))
        conv_type = conv.get("type", "group")

        if auth_mode == "device_code" and conv_type == "oneOnOne":
            try:
                me_resp = chat_client.get("/me", params={"$select": "userPrincipalName"})
                me_upn = (me_resp.json().get("userPrincipalName") or "").lower()
            except Exception:
                me_upn = ""

            members = [m.lower() for m in conv.get("members", [])]
            if me_upn and me_upn not in members:
                logger.info(
                    "Skipping oneOnOne chat '%s': delegated caller is not one of the configured members.",
                    conv_id,
                )
                actions.append(
                    {
                        "action": "skip_chat",
                        "conversation_id": conv_id,
                        "reason": "caller_not_member_for_oneOnOne",
                    }
                )
                continue

        logger.info(
            "[BETA] Creating chat '%s' with %d members",
            conv_id,
            len(conv["members"]),
        )

        try:
            try:
                chat_data = _create_chat(chat_client, conv, run_id, user_ids)
            except Exception as exc:
                # If app-only call is forbidden/unauthorized, try delegated once.
                exc_str = str(exc)
                authz_failure = (
                    "403" in exc_str
                    or "401" in exc_str
                    or "Forbidden" in exc_str
                    or "Unauthorized" in exc_str
                )
                if auth_mode == "client_secret" and authz_failure:
                    logger.info(
                        "Chat create hit auth boundary — attempting delegated auth …"
                    )
                    try:
                        chat_client = build_delegated_client(cfg, dry_run=client.dry_run)
                        chat_data = _create_chat(chat_client, conv, run_id, user_ids)
                    except Exception as delegated_exc:
                        logger.warning(
                            "Could not obtain delegated credentials for chat seeding — "
                            "skipping conversation '%s': %s",
                            conv_id,
                            delegated_exc,
                        )
                        actions.append(
                            {
                                "action": "skip_chat",
                                "conversation_id": conv_id,
                                "reason": "delegated_auth_failed",
                            }
                        )
                        continue
                else:
                    raise
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
            message_send_blocked = False
            for msg_cfg in conv.get("messages", []):
                # Accept both dict {"text": "..."} and plain string formats
                message = (
                    msg_cfg["text"] if isinstance(msg_cfg, dict) else str(msg_cfg)
                )
                logger.info(
                    "[BETA] Sending chat message in '%s'",
                    conv_id,
                )
                try:
                    _send_chat_message(chat_client, chat_id, message, run_id)
                except Exception as exc:
                    exc_str = str(exc)
                    detail = ""
                    if hasattr(exc, "response") and exc.response is not None:
                        try:
                            detail = (
                                exc.response.json().get("error", {}).get("message", "")
                            )
                        except Exception:
                            detail = exc.response.text or ""
                    app_only_message_block = (
                        auth_mode == "client_secret"
                        and ("401" in exc_str or "Unauthorized" in exc_str)
                        and "import purposes" in detail
                    )
                    if app_only_message_block:
                        logger.info(
                            "Skipping messages for chat '%s' in app-only mode: %s",
                            conv_id,
                            exc,
                        )
                        actions.append(
                            {
                                "action": "skip_chat_messages",
                                "conversation_id": conv_id,
                                "chat_id": chat_id,
                                "reason": "app_only_import_only",
                            }
                        )
                        message_send_blocked = True
                        break
                    raise
                actions.append(
                    {
                        "action": "send_chat_message",
                        "conversation_id": conv_id,
                        "chat_id": chat_id,
                        "api": "beta",
                    }
                )
            if message_send_blocked:
                continue
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
