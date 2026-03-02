"""Microsoft Graph client with retry / back-off and auth helpers."""

from __future__ import annotations

import json
import logging
import time
from typing import Any

import httpx
from azure.identity import (
    AzureCliCredential,
    ChainedTokenCredential,
    ClientSecretCredential,
    DeviceCodeCredential,
    TokenCachePersistenceOptions,
)

from m365seed.config import resolve_secret

logger = logging.getLogger("m365seed.graph")

GRAPH_BASE = "https://graph.microsoft.com"
GRAPH_V1 = f"{GRAPH_BASE}/v1.0"
GRAPH_BETA = f"{GRAPH_BASE}/beta"
SCOPES = ["https://graph.microsoft.com/.default"]

# Delegated scopes requested explicitly for device-code fallback.
# `.default` alone may not return dynamically-added delegated permissions
# for public client flows.
DELEGATED_SCOPES = [
    "https://graph.microsoft.com/ChannelMessage.Send",
    "https://graph.microsoft.com/Chat.Create",
    "https://graph.microsoft.com/Chat.ReadWrite",
    "https://graph.microsoft.com/User.Read",
]

# Retry settings
MAX_RETRIES = 5
DEFAULT_RETRY_AFTER = 5  # seconds


# ---------------------------------------------------------------------------
# Auth helpers
# ---------------------------------------------------------------------------


def build_credential(cfg: dict[str, Any]):
    """Return an ``azure-identity`` credential based on config ``auth.mode``."""
    mode = cfg["auth"]["mode"]
    tenant_id = cfg["tenant"]["tenant_id"]
    client_id = cfg["auth"]["client_id"]

    if mode == "client_secret":
        secret = resolve_secret(cfg)
        return ClientSecretCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            client_secret=secret,
        )
    elif mode == "device_code":
        cache_opts = TokenCachePersistenceOptions(
            name="m365seed-device-code-cache",
            allow_unencrypted_storage=True,
        )
        return DeviceCodeCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            cache_persistence_options=cache_opts,
        )
    else:
        raise ValueError(f"Unsupported auth mode: {mode}")


def build_delegated_client(
    cfg: dict[str, Any],
    dry_run: bool = False,
) -> "GraphClient":
    """Build a :class:`GraphClient` for delegated operations.

    Used when the primary auth mode is ``client_secret`` but the
    operation requires delegated permissions (e.g. posting Teams
    channel messages or creating chats).

    Auth strategy is non-interactive first:
    1) ``AzureCliCredential`` (reuses existing ``az login`` session)
    2) ``DeviceCodeCredential`` fallback if CLI token is unavailable
    """
    tenant_id = cfg["tenant"]["tenant_id"]
    client_id = cfg["auth"]["client_id"]

    cache_opts = TokenCachePersistenceOptions(
        name="m365seed-device-code-cache",
        allow_unencrypted_storage=True,
    )

    delegated_credential = ChainedTokenCredential(
        DeviceCodeCredential(
            tenant_id=tenant_id,
            client_id=client_id,
            cache_persistence_options=cache_opts,
        ),
        AzureCliCredential(tenant_id=tenant_id),
    )

    delegated_cfg = {
        **cfg,
        "auth": {**cfg["auth"], "mode": "device_code"},
    }
    return GraphClient(
        delegated_cfg,
        dry_run=dry_run,
        scopes=DELEGATED_SCOPES,
        credential=delegated_credential,
    )


# ---------------------------------------------------------------------------
# Graph client
# ---------------------------------------------------------------------------


class GraphClient:
    """Thin wrapper around ``httpx.Client`` for Microsoft Graph calls.

    Handles:
    - Bearer token acquisition via ``azure-identity``
    - Automatic retry with exponential back-off on HTTP 429 / 503 / 504
    - Dry-run mode (logs the request instead of sending it)
    """

    def __init__(
        self,
        cfg: dict[str, Any],
        dry_run: bool = False,
        scopes: list[str] | None = None,
        credential: Any | None = None,
    ) -> None:
        self.cfg = cfg
        self.dry_run = dry_run
        self._credential = credential or build_credential(cfg)
        self._scopes = scopes
        self._http = httpx.Client(timeout=60.0)

    # -- token ---------------------------------------------------------------

    def _get_token(self) -> str:
        scopes = self._scopes or SCOPES
        try:
            token = self._credential.get_token(*scopes)
        except Exception as exc:
            # AzureCliCredential requires exactly one scope per request.
            # When used in a chain for delegated flows, fall back to the
            # first scope to avoid hard failure on multi-scope calls.
            if len(scopes) > 1 and "exactly one scope" in str(exc):
                token = self._credential.get_token(scopes[0])
            else:
                raise
        return token.token

    def _auth_headers(self) -> dict[str, str]:
        return {
            "Authorization": f"Bearer {self._get_token()}",
            "Content-Type": "application/json",
        }

    # -- core request --------------------------------------------------------

    def request(
        self,
        method: str,
        url: str,
        *,
        json_body: dict | list | None = None,
        content: bytes | None = None,
        headers: dict[str, str] | None = None,
        params: dict[str, str] | None = None,
    ) -> httpx.Response:
        """Execute an HTTP request against Graph with retry logic.

        In ``dry_run`` mode the request is logged but never sent, and a
        synthetic 200-response is returned.
        """
        merged_headers = self._auth_headers()
        if headers:
            merged_headers.update(headers)

        if self.dry_run:
            body_preview = ""
            if json_body:
                body_preview = json.dumps(json_body, indent=2)[:500]
            logger.info(
                "[DRY-RUN] %s %s  body=%s",
                method,
                url,
                body_preview or "(none)",
            )
            # Return a synthetic response
            return httpx.Response(
                status_code=200,
                json={"id": "dry-run-id", "@dry_run": True},
                request=httpx.Request(method, url),
            )

        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = self._http.request(
                    method,
                    url,
                    headers=merged_headers,
                    json=json_body,
                    content=content,
                    params=params,
                )
            except httpx.TransportError as exc:
                logger.warning("Transport error (attempt %d): %s", attempt, exc)
                if attempt == MAX_RETRIES:
                    raise
                time.sleep(DEFAULT_RETRY_AFTER * attempt)
                continue

            if resp.status_code == 429 or resp.status_code in (503, 504):
                retry_after = int(
                    resp.headers.get("Retry-After", DEFAULT_RETRY_AFTER)
                )
                logger.warning(
                    "Throttled %d (attempt %d/%d) — retrying in %ds",
                    resp.status_code,
                    attempt,
                    MAX_RETRIES,
                    retry_after,
                )
                time.sleep(retry_after)
                continue

            if resp.status_code >= 400:
                # Log the response body for debugging before raising.
                # 404s are often expected (idempotency existence checks),
                # so log them at DEBUG; everything else at WARNING.
                try:
                    body = resp.json()
                    detail = body.get("error", {}).get("message", resp.text[:500])
                except Exception:
                    detail = resp.text[:500]
                duplicate_channel_name = (
                    resp.status_code == 400
                    and isinstance(detail, str)
                    and "Channel name already existed" in detail
                )
                app_only_chat_import_restriction = (
                    resp.status_code == 401
                    and isinstance(detail, str)
                    and "import purposes" in detail
                )
                log_level = (
                    logging.DEBUG
                    if resp.status_code in (404, 409)
                    or duplicate_channel_name
                    or app_only_chat_import_restriction
                    else logging.WARNING
                )
                logger.log(
                    log_level,
                    "Graph %d detail for %s %s: %s",
                    resp.status_code, method, url, detail,
                )
            resp.raise_for_status()
            return resp

        # Should not reach here, but just in case
        raise RuntimeError("Max retries exhausted")  # pragma: no cover

    # -- convenience ---------------------------------------------------------

    def get(self, path: str, *, base: str = GRAPH_V1, **kw) -> httpx.Response:
        return self.request("GET", f"{base}{path}", **kw)

    def post(self, path: str, *, base: str = GRAPH_V1, **kw) -> httpx.Response:
        return self.request("POST", f"{base}{path}", **kw)

    def put(self, path: str, *, base: str = GRAPH_V1, **kw) -> httpx.Response:
        return self.request("PUT", f"{base}{path}", **kw)

    def patch(self, path: str, *, base: str = GRAPH_V1, **kw) -> httpx.Response:
        return self.request("PATCH", f"{base}{path}", **kw)

    def delete(self, path: str, *, base: str = GRAPH_V1, **kw) -> httpx.Response:
        return self.request("DELETE", f"{base}{path}", **kw)

    # -- validation helpers --------------------------------------------------

    def check_auth(self) -> dict:
        """Validate authentication by calling ``/me`` (or ``/organization``).

        For app-only (client_secret) we call ``/organization``.
        """
        if self.cfg["auth"]["mode"] == "client_secret":
            resp = self.get("/organization")
        else:
            resp = self.get("/me")
        return resp.json()

    def check_user_exists(self, upn: str) -> bool:
        """Return True if the user exists in the tenant."""
        try:
            resp = self.get(f"/users/{upn}", params={"$select": "id,userPrincipalName"})
            return resp.status_code == 200
        except httpx.HTTPStatusError:
            return False

    def list_permissions(self) -> list[str]:
        """Best-effort: return the list of granted OAuth2 scopes/roles.

        Works only for delegated tokens; for app-only tokens this may
        return an empty list (Graph does not expose app role assignments
        on /me).
        """
        try:
            resp = self.get("/me")
            # Not directly available; return empty for app-only
            return []
        except Exception:
            return []
