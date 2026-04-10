"""Interactive setup wizard — ``m365seed setup``.

Walks the user through tenant configuration, theme selection,
content toggles, and config generation.  Optionally validates
and runs a dry-run at the end.
"""

from __future__ import annotations

import os
import re
import shutil
import subprocess
import json
import time
import uuid
from pathlib import Path
from typing import Any, Optional

import typer
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Confirm, Prompt
from rich.table import Table

console = Console()

# ── Defaults ─────────────────────────────────────────────────
THEMES = ["healthcare", "pharma", "medtech", "payor"]
THEME_LABELS = {
    "healthcare": "Health Provider — clinical ops, care coordination",
    "pharma": "Pharma / Life Science — research, clinical trials",
    "medtech": "MedTech — product dev, manufacturing, 510(k)",
    "payor": "Health Payor — claims, member services",
}

DEFAULT_USERS = [
    {"upn": "AllanD@{domain}", "role": "Clinical Ops Manager"},
    {"upn": "MeganB@{domain}", "role": "Care Manager — Dr. Donald Wilson"},
    {"upn": "NestorW@{domain}", "role": "Care Manager — Dr. Daniel Rodriguez"},
    {"upn": "LeeG@{domain}", "role": "Nurse Manager"},
    {"upn": "JoniS@{domain}", "role": "Compliance Officer"},
]

CONTENT_SECTIONS = [
    ("mail", "Email threads with attachments", True),
    ("files", "OneDrive / SharePoint files", True),
    ("calendar", "Calendar events", True),
    ("teams", "Teams channels & posts (beta)", False),
    ("chats", "Teams 1:1 and group chats (beta)", False),
    ("sharepoint", "SharePoint sites, pages, docs", False),
    ("planner", "Planner plans, buckets, tasks", False),
]

THEMED_TEAM_NAMES: dict[str, str] = {
    "healthcare": "Contoso Health System",
    "pharma": "Contoso Pharmaceuticals",
    "medtech": "Contoso Medical Devices",
    "payor": "Contoso Health Plans",
}

GUID_RE = re.compile(
    r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$",
    re.IGNORECASE,
)


# ── Helpers ──────────────────────────────────────────────────


def _ask_guid(label: str, env_var: str | None = None) -> str:
    """Prompt for a GUID, optionally pre-filling from an env var."""
    default = os.environ.get(env_var, "") if env_var else ""
    while True:
        value = Prompt.ask(
            f"  {label}",
            default=default or None,
        )
        if value and GUID_RE.match(value.strip()):
            return value.strip()
        console.print("  [red]Please enter a valid GUID (xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx).[/red]")


def _guess_tenant_domain(tenant_id: str) -> str | None:
    """Infer tenant default domain from current Azure CLI context when possible."""
    if shutil.which("az") is None:
        return None

    try:
        probe = _run_az_command(["account", "show", "--output", "json"])
        if probe.returncode == 0 and probe.stdout:
            acct = json.loads(probe.stdout)
            acct_tenant = (acct.get("tenantId") or "").strip().lower()
            if acct_tenant == tenant_id.lower():
                domain = (acct.get("tenantDefaultDomain") or "").strip()
                if domain and "." in domain:
                    return domain.lower()
    except Exception:
        return None

    return None


def _run_az_command(args: list[str]) -> subprocess.CompletedProcess:
    """Run an ``az`` command after purging Azure CLI's incompatible HTTP cache."""
    from m365seed.register import _ensure_msal_cache_healthy

    _ensure_msal_cache_healthy()
    return subprocess.run(
        ["az", "--only-show-errors", *args],
        capture_output=True,
        text=True,
        check=False,
    )


def _summarize_az_error(result: subprocess.CompletedProcess) -> str:
    """Extract the most useful Azure CLI error line for user-facing output."""
    lines = [line.strip() for line in (result.stderr or result.stdout or "").splitlines() if line.strip()]
    if not lines:
        return "unknown error"
    error_lines = [line for line in lines if line.startswith("ERROR:")]
    if error_lines:
        if (
            len(error_lines) > 1
            and error_lines[0].startswith("ERROR: The command failed with an unexpected error")
        ):
            return error_lines[1]
        return error_lines[0]
    if (
        len(lines) > 1
        and lines[0].startswith("ERROR: The command failed with an unexpected error")
    ):
        return lines[1]
    non_warning = [line for line in lines if not line.startswith("WARNING:")]
    if non_warning:
        return non_warning[0]
    return lines[0]


def _summarize_graph_error(exc: Exception) -> str:
    """Extract the most useful Microsoft Graph error detail."""
    response = getattr(exc, "response", None)
    if response is not None:
        try:
            payload = response.json()
            detail = payload.get("error", {}).get("message", "")
            if detail:
                return str(detail)
        except Exception:
            pass
        text = getattr(response, "text", "")
        if text:
            return str(text)[:300]
    return str(exc) or exc.__class__.__name__


def _build_setup_app_graph_client(
    tenant_id: str,
    client_id: str,
    secret_env: str,
) -> Any | None:
    """Build an app-only Graph client for tenant-scoped setup actions."""
    if not client_id or not secret_env or not os.environ.get(secret_env):
        return None

    try:
        from m365seed.graph import GraphClient

        return GraphClient(
            {
                "tenant": {"tenant_id": tenant_id},
                "auth": {
                    "mode": "client_secret",
                    "client_id": client_id,
                    "client_secret_env": secret_env,
                },
            }
        )
    except Exception:
        return None


def _build_setup_delegated_graph_client(tenant_id: str) -> Any | None:
    """Build a delegated Graph client backed by the current Azure CLI login."""
    if shutil.which("az") is None:
        return None

    try:
        from m365seed.graph import GraphClient, build_azure_cli_credential

        return GraphClient(
            {
                "tenant": {"tenant_id": tenant_id},
                "auth": {
                    "mode": "device_code",
                    "client_id": "azure-cli-setup",
                },
            },
            credential=build_azure_cli_credential(tenant_id),
            scopes=["https://graph.microsoft.com/.default"],
        )
    except Exception:
        return None


def _build_setup_graph_client(
    tenant_id: str,
    client_id: str,
    secret_env: str,
) -> Any | None:
    """Build the best available Graph client for interactive setup.

    Preference order:
    1. Delegated Graph via the current Azure CLI admin session
    2. App-only Graph via the supplied client ID and secret env var
    """
    delegated = _build_setup_delegated_graph_client(tenant_id)
    if delegated is not None:
        return delegated
    return _build_setup_app_graph_client(tenant_id, client_id, secret_env)


def _get_verified_tenant_domains(graph_client: Any | None) -> list[str]:
    """Return verified domains for the tenant, ordered by preference."""
    if graph_client is None:
        return []

    try:
        resp = graph_client.get("/organization", params={"$select": "verifiedDomains"})
        payload = resp.json()
        ranked: list[tuple[int, int, str]] = []
        order = 0
        for org in payload.get("value", []):
            for domain in org.get("verifiedDomains", []):
                name = (domain.get("name") or "").strip().lower()
                if not name:
                    continue
                priority = 3
                if domain.get("isDefault"):
                    priority = 0
                elif domain.get("isInitial"):
                    priority = 1
                elif name.endswith(".onmicrosoft.com"):
                    priority = 2
                ranked.append((priority, order, name))
                order += 1

        ranked.sort()
        resolved: list[str] = []
        seen: set[str] = set()
        for _, _, name in ranked:
            if name not in seen:
                seen.add(name)
                resolved.append(name)
        return resolved
    except Exception:
        return []


def _resolve_tenant_domain(domain: str, verified_domains: list[str]) -> tuple[str, bool]:
    """Return a verified tenant domain when the current value is invalid."""
    normalized = domain.strip().lower()
    if not verified_domains or normalized in verified_domains:
        return normalized, False
    return verified_domains[0], True


def _ask_labeled_choice(
    label: str,
    options: list[tuple[str, str]],
    *,
    default: str,
) -> str:
    """Prompt for a choice using friendly labels while accepting numbers or keys."""
    for i, (value, description) in enumerate(options, 1):
        tag = " [dim](default)[/dim]" if value == default else ""
        console.print(f"    {i}. {value} — {description}{tag}")

    values = [value for value, _ in options]
    while True:
        raw = Prompt.ask(label, default=default)
        stripped = raw.strip().lower()
        if stripped.isdigit():
            idx = int(stripped) - 1
            if 0 <= idx < len(options):
                return options[idx][0]
        for value in values:
            if stripped == value.lower():
                return value
        console.print(f"  [red]Choose 1-{len(options)} or type one of: {', '.join(values)}.[/red]")


def _display_name_from_upn(upn: str) -> str:
    """Best-effort display name derived from a UPN alias."""
    alias = upn.split("@", 1)[0].strip()
    if not alias:
        return "Demo User"
    humanized = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", alias)
    humanized = humanized.replace(".", " ").replace("_", " ").replace("-", " ")
    return re.sub(r"\s+", " ", humanized).strip() or alias


def _create_missing_demo_users(
    users: list[dict[str, str]],
    temp_password: str,
    *,
    force_change_next_sign_in: bool,
    graph_client: Any | None = None,
) -> tuple[int, int, list[dict[str, str]]]:
    """Create missing Entra ID users.

    Returns ``(success_count, failed_count, created_users)``.
    """
    if not users:
        return (0, 0, [])

    if graph_client is None and shutil.which("az") is None:
        console.print(
            "  [yellow]⚠ Azure CLI not found; skipping user creation step.[/yellow]"
        )
        return (0, len(users), [])

    success = 0
    failed = 0
    created: list[dict[str, str]] = []
    force_change = "true" if force_change_next_sign_in else "false"

    for user in users:
        upn = user["upn"]
        alias = upn.split("@", 1)[0]
        mail_nickname = re.sub(r"[^A-Za-z0-9._-]", "", alias) or f"demo{uuid.uuid4().hex[:8]}"
        display_name = user.get("display_name") or _display_name_from_upn(upn)
        if graph_client is not None:
            try:
                graph_client.post(
                    "/users",
                    json_body={
                        "accountEnabled": True,
                        "displayName": display_name,
                        "mailNickname": mail_nickname,
                        "userPrincipalName": upn,
                        "passwordProfile": {
                            "forceChangePasswordNextSignIn": force_change_next_sign_in,
                            "password": temp_password,
                        },
                    },
                )
                success += 1
                created.append(user)
                console.print(f"    [green]✓[/green] Created user {upn}")
                continue
            except Exception as exc:
                failed += 1
                reason = _summarize_graph_error(exc)
                console.print(f"    [red]✗[/red] Failed to create {upn}: {reason}")
                continue

        result = _run_az_command(
            [
                "ad",
                "user",
                "create",
                "--display-name",
                display_name,
                "--password",
                temp_password,
                "--user-principal-name",
                upn,
                "--force-change-password-next-sign-in",
                force_change,
                "--mail-nickname",
                mail_nickname,
                "--output",
                "json",
            ]
        )

        if result.returncode == 0:
            success += 1
            created.append(user)
            console.print(f"    [green]✓[/green] Created user {upn}")
        else:
            failed += 1
            reason = _summarize_az_error(result)
            console.print(f"    [red]✗[/red] Failed to create {upn}: {reason}")

    return (success, failed, created)


def _discover_tenant_users(graph_client: Any | None = None) -> list[dict[str, str]]:
    """Query the tenant for actual users.

    Returns a list of dicts with 'upn' and 'displayName' keys,
    or an empty list if discovery fails.
    """
    if graph_client is not None:
        try:
            resp = graph_client.get(
                "/users",
                params={
                    "$select": "userPrincipalName,displayName,mail",
                    "$top": "999",
                },
            )
            raw = resp.json().get("value", [])
            users = [
                {
                    "upn": u.get("userPrincipalName", ""),
                    "displayName": u.get("displayName", ""),
                    "mail": u.get("mail", ""),
                }
                for u in raw
                if u.get("userPrincipalName") and "#EXT#" not in u.get("userPrincipalName", "")
            ]
            return users
        except Exception:
            pass

    if shutil.which("az") is None:
        return []

    try:
        result = _run_az_command(
            [
                "ad", "user", "list",
                "--query",
                "[].{upn:userPrincipalName, displayName:displayName, mail:mail}",
                "--output", "json",
            ]
        )
        if result.returncode == 0 and result.stdout:
            raw = json.loads(result.stdout)
            # Filter out system/guest accounts — keep only *.onmicrosoft.com
            # or accounts without #EXT# in the UPN
            users = [
                u for u in raw
                if u.get("upn") and "#EXT#" not in u["upn"]
            ]
            return users
    except Exception:
        pass

    return []


def _match_default_users(
    domain: str,
    tenant_users: list[dict[str, str]],
) -> list[dict[str, str]]:
    """Match default demo user aliases against actual tenant users.

    Returns a list of dicts with 'upn', 'role', and 'matched' keys.
    If a default user is found in the tenant, their real UPN is used.
    """
    # Build a lookup by alias (part before @), case-insensitive
    alias_lookup: dict[str, str] = {}
    for tu in tenant_users:
        upn = tu.get("upn", "")
        if "@" in upn:
            alias = upn.split("@")[0].lower()
            alias_lookup[alias] = upn

    matched: list[dict[str, str]] = []
    for default in DEFAULT_USERS:
        template_upn = default["upn"].format(domain=domain)
        alias = template_upn.split("@")[0].lower()

        if alias in alias_lookup:
            matched.append({
                "upn": alias_lookup[alias],
                "role": default["role"],
                "matched": "true",
            })
        else:
            matched.append({
                "upn": template_upn,
                "role": default["role"],
                "matched": "false",
            })

    return matched


def _ask_choice(label: str, choices: list[str], default: str) -> str:
    """Present numbered choices and return the selection."""
    for i, c in enumerate(choices, 1):
        extra = THEME_LABELS.get(c, "")
        tag = " [dim](default)[/dim]" if c == default else ""
        desc = f" — {extra}" if extra else ""
        console.print(f"    {i}. {c}{desc}{tag}")

    while True:
        raw = Prompt.ask(f"  {label}", default=default)
        stripped = raw.strip()
        # Accept number or name
        if stripped.isdigit():
            idx = int(stripped) - 1
            if 0 <= idx < len(choices):
                return choices[idx]
        if stripped.lower() in [c.lower() for c in choices]:
            return stripped.lower()
        console.print(f"  [red]Choose 1‑{len(choices)} or type a name.[/red]")


def _ask_users(domain: str, graph_client: Any | None = None) -> list[dict[str, str]]:
    """Let the user accept defaults or enter custom UPNs.

    Queries Azure CLI to discover actual tenant users and matches
    default demo aliases against them.  Missing users are flagged.
    """
    # Discover real users in the tenant
    console.print("\n  [dim]Querying tenant for existing users…[/dim]")
    tenant_users = _discover_tenant_users(graph_client=graph_client)

    if tenant_users:
        console.print(
            f"  [dim]Found {len(tenant_users)} user(s) in tenant.[/dim]"
        )
        matched = _match_default_users(domain, tenant_users)
        found = [u for u in matched if u["matched"] == "true"]
        missing = [u for u in matched if u["matched"] == "false"]

        console.print("\n  [bold]Default demo users[/bold]:")
        for u in matched:
            if u["matched"] == "true":
                console.print(
                    f"    [green]✓[/green] {u['upn']}  ({u['role']})"
                )
            else:
                console.print(
                    f"    [red]✗[/red] {u['upn']}  ({u['role']}) — [red]not found in tenant[/red]"
                )

        if missing:
            console.print(
                f"\n  [yellow]⚠ {len(missing)} of {len(matched)} default "
                f"user(s) not found in tenant.[/yellow]"
            )
            if not found:
                console.print(
                    "  [yellow]  None of the defaults exist. "
                    "You can create them later in setup or choose different users now.[/yellow]"
                )

            options: list[tuple[str, str]] = []
            if found:
                options.append((
                    "use-existing",
                    "Use only the matched users that already exist in the tenant",
                ))
            options.append((
                "create-missing",
                "Use the standard demo users and create any missing accounts later in setup",
            ))
            if tenant_users:
                options.append((
                    "pick-existing",
                    "Select different users that already exist in the tenant",
                ))
            options.append((
                "manual",
                "Enter user principal names manually",
            ))
            choice = _ask_labeled_choice(
                "\n  How to proceed?",
                options,
                default="create-missing" if not found else "use-existing",
            )
        else:
            # All matched
            if Confirm.ask("\n  Use these defaults?", default=True):
                return [{k: v for k, v in u.items() if k != "matched"} for u in matched]
            choice = "manual"

        if choice == "use-existing":
            result = [{k: v for k, v in u.items() if k != "matched"} for u in found]
            console.print(
                f"  → Using {len(result)} verified user(s)."
            )
            return result
        elif choice == "create-missing":
            console.print(
                "  → Using the standard demo users. Missing accounts can be created in Step 6."
            )
            return [{k: v for k, v in u.items() if k != "matched"} for u in matched]
        elif choice == "pick-existing":
            return _pick_from_tenant(tenant_users)
        else:
            return _manual_user_entry()
    else:
        # Could not discover users — fall back to original behavior
        console.print(
            "  [yellow]⚠ Could not query tenant users automatically.[/yellow]"
        )
        defaults = [
            {**u, "upn": u["upn"].format(domain=domain)} for u in DEFAULT_USERS
        ]
        console.print("\n  [bold]Default demo users[/bold]:")
        for u in defaults:
            console.print(f"    • {u['upn']}  ({u['role']})")

        if Confirm.ask("\n  Use these defaults?", default=True):
            return defaults

        return _manual_user_entry()


def _pick_from_tenant(tenant_users: list[dict[str, str]]) -> list[dict[str, str]]:
    """Let the user pick from discovered tenant users."""
    console.print("\n  [bold]Available tenant users:[/bold]")
    for i, u in enumerate(tenant_users, 1):
        name = u.get("displayName", "")
        console.print(f"    {i:2d}. {u['upn']}  ({name})")

    console.print(
        "\n  Enter user numbers separated by commas (e.g. 1,3,5),"
        " or 'all' for everyone:"
    )
    raw = Prompt.ask("  Selection", default="all")

    if raw.strip().lower() == "all":
        selected = tenant_users
    else:
        indices: list[int] = []
        for part in raw.split(","):
            part = part.strip()
            if part.isdigit():
                idx = int(part) - 1
                if 0 <= idx < len(tenant_users):
                    indices.append(idx)
        selected = [tenant_users[i] for i in indices] if indices else tenant_users

    users: list[dict[str, str]] = []
    for u in selected:
        role = Prompt.ask(
            f"    Role for {u['upn']}",
            default="Demo User",
        )
        users.append({"upn": u["upn"], "role": role})

    return users


def _manual_user_entry() -> list[dict[str, str]]:
    """Prompt for manual UPN + role entry."""
    users: list[dict[str, str]] = []
    console.print("  Enter users (blank UPN to finish):")
    while True:
        upn = Prompt.ask("    UPN", default="")
        if not upn:
            break
        role = Prompt.ask("    Role", default="Demo User")
        users.append({"upn": upn, "role": role})
    return users


def _check_user_exists(upn: str, graph_client: Any | None = None) -> bool:
    """Check whether a user exists in the tenant."""
    if graph_client is not None:
        try:
            return bool(graph_client.check_user_exists(upn))
        except Exception:
            return False

    if shutil.which("az") is None:
        return True  # can't verify — optimistic
    try:
        result = _run_az_command(
            ["ad", "user", "show", "--id", upn, "--query", "id", "--output", "tsv"]
        )
        return result.returncode == 0 and bool(result.stdout.strip())
    except Exception:
        return True  # can't verify — optimistic


def _reset_demo_user_passwords(
    users: list[dict[str, str]],
    temp_password: str,
    *,
    force_change_next_sign_in: bool,
    graph_client: Any | None = None,
) -> tuple[int, int]:
    """Reset passwords for selected demo users.

    Returns (success_count, failed_count).
    """
    if not users:
        return (0, 0)

    if graph_client is None and shutil.which("az") is None:
        console.print(
            "  [yellow]⚠ Azure CLI not found; skipping password reset step.[/yellow]"
        )
        return (0, len(users))

    success = 0
    failed = 0
    force_change = "true" if force_change_next_sign_in else "false"

    for user in users:
        upn = user["upn"]
        if graph_client is not None:
            try:
                graph_client.patch(
                    f"/users/{upn}",
                    json_body={
                        "passwordProfile": {
                            "forceChangePasswordNextSignIn": force_change_next_sign_in,
                            "password": temp_password,
                        }
                    },
                )
                success += 1
                console.print(f"    [green]✓[/green] Reset password for {upn}")
                continue
            except Exception:
                # Fall through to Azure CLI as a best-effort fallback.
                if shutil.which("az") is None:
                    failed += 1
                    console.print(f"    [red]✗[/red] Failed to reset {upn}: insufficient permissions or Graph error")
                    continue

        result = _run_az_command(
            [
                "ad",
                "user",
                "update",
                "--id",
                upn,
                "--password",
                temp_password,
                "--force-change-password-next-sign-in",
                force_change,
            ]
        )

        if result.returncode == 0:
            success += 1
            console.print(f"    [green]✓[/green] Reset password for {upn}")
        else:
            failed += 1
            reason = _summarize_az_error(result)
            console.print(f"    [red]✗[/red] Failed to reset {upn}: {reason}")

    return (success, failed)

def _resolve_owner_refs(graph_client: Any | None, owner_upns: list[str] | None) -> list[str]:
    """Resolve user UPNs to Graph owner reference URLs."""
    if graph_client is None or not owner_upns:
        return []

    from m365seed.graph import GRAPH_V1

    owner_refs: list[str] = []
    seen: set[str] = set()
    for upn in owner_upns:
        try:
            resp = graph_client.get(f"/users/{upn}", params={"$select": "id"})
            user_id = (resp.json() or {}).get("id", "")
            if user_id and user_id not in seen:
                seen.add(user_id)
                owner_refs.append(f"{GRAPH_V1}/users/{user_id}")
        except Exception:
            continue
        if len(owner_refs) >= 2:
            break
    return owner_refs


def _discover_teams(
    tenant_id: str,
    graph_client: Any | None = None,
) -> list[dict[str, str]]:
    """List Microsoft 365 Groups that can back Teams or Planner."""
    if graph_client is not None:
        try:
            resp = graph_client.get(
                "/groups",
                params={
                    "$filter": "groupTypes/any(c:c eq 'Unified')",
                    "$select": "id,displayName",
                    "$top": "999",
                },
            )
            groups = resp.json().get("value", [])
            return [g for g in groups if g.get("id") and g.get("displayName")]
        except Exception:
            pass

    if shutil.which("az") is None:
        return []

    try:
        result = _run_az_command(
            [
                "ad", "group", "list",
                "--query",
                "[?contains(groupTypes,'Unified')].{id:id, displayName:displayName}",
                "--output", "json",
            ]
        )
        if result.returncode == 0 and result.stdout:
            groups = json.loads(result.stdout)
            return [g for g in groups if g.get("id") and g.get("displayName")]
    except Exception:
        pass

    return []


def _add_az_login_user_to_group(group_id: str) -> None:
    """Add the current ``az login`` user as an owner+member of *group_id*.

    This prevents 403 errors later when the delegated user (the same
    ``az login`` identity) tries to post Teams channel messages.
    """
    if shutil.which("az") is None:
        return

    try:
        acct_result = _run_az_command(["ad", "signed-in-user", "show", "--query", "id", "-o", "tsv"])
        if acct_result.returncode != 0 or not acct_result.stdout.strip():
            return
        user_id = acct_result.stdout.strip()
        ref_body = json.dumps({
            "@odata.id": f"https://graph.microsoft.com/v1.0/directoryObjects/{user_id}"
        })
        for role in ("owners", "members"):
            _run_az_command([
                "rest", "--method", "POST",
                "--url", f"https://graph.microsoft.com/v1.0/groups/{group_id}/{role}/$ref",
                "--body", ref_body,
            ])
        console.print(
            f"  [green]✓[/green] Added az-login user as team owner/member."
        )
    except Exception as exc:
        console.print(
            f"  [yellow]⚠ Could not auto-add az-login user to team: {exc}[/yellow]"
        )


def _create_team_group(
    display_name: str,
    description: str = "",
    *,
    owner_upns: list[str] | None = None,
    graph_client: Any | None = None,
) -> str | None:
    """Create an M365 Unified Group and Team-enable it via Microsoft Graph.

    Returns the new group's ID, or ``None`` on failure.
    """
    if graph_client is None and shutil.which("az") is None:
        console.print("  [red]Azure CLI not found — cannot create group.[/red]")
        return None

    mail_nick = re.sub(r"[^a-zA-Z0-9]", "", display_name).lower()[:40]
    # Append a short unique suffix to avoid mailNickname collisions
    mail_nick = f"{mail_nick}{uuid.uuid4().hex[:8]}"
    desc = description or f"Demo team for {display_name}"

    body = json.dumps({
        "displayName": display_name,
        "mailNickname": mail_nick,
        "description": desc,
        "groupTypes": ["Unified"],
        "mailEnabled": True,
        "securityEnabled": False,
        "visibility": "Private",
    })

    if graph_client is not None:
        owner_refs = _resolve_owner_refs(graph_client, owner_upns)
        group_body = {
            "displayName": display_name,
            "mailNickname": mail_nick,
            "description": desc,
            "groupTypes": ["Unified"],
            "mailEnabled": True,
            "securityEnabled": False,
            "visibility": "Private",
        }
        if owner_refs:
            group_body["owners@odata.bind"] = owner_refs

        try:
            resp = graph_client.post("/groups", json_body=group_body)
            group_data = resp.json()
            group_id = group_data.get("id", "")
            if not group_id:
                console.print("  [red]✗ Group created but no ID returned.[/red]")
                return None

            console.print(
                f"  [green]✓[/green] Created M365 Group "
                f"[cyan]{display_name}[/cyan] ({group_id[:8]}…)"
            )

            if not owner_refs:
                console.print(
                    "  [yellow]⚠ No existing user owner was available to assign to the group. "
                    "Teams enablement may fail until you add an owner.[/yellow]"
                )

            console.print("  [dim]Enabling Teams on the group…[/dim]")
            for attempt in range(3):
                try:
                    graph_client.put(f"/groups/{group_id}/team", json_body={})
                    console.print("  [green]✓[/green] Team enabled on group.")
                    _add_az_login_user_to_group(group_id)
                    return group_id
                except Exception:
                    if attempt < 2:
                        time.sleep(10)
                        console.print(
                            f"  [dim]  Waiting for provisioning (attempt {attempt + 2}/3)…[/dim]"
                        )

            console.print(
                "  [yellow]⚠ Group created but could not enable Teams. "
                "You may need to team-enable it manually in Teams admin.[/yellow]"
            )
            return group_id
        except Exception as exc:
            console.print(
                f"  [red]✗ Failed to create group: {_summarize_graph_error(exc)}[/red]"
            )
            return None

    try:
        result = _run_az_command(
            [
                "rest",
                "--method", "POST",
                "--url", "https://graph.microsoft.com/v1.0/groups",
                "--body", body,
            ]
        )
        if result.returncode != 0 or not result.stdout.strip():
            console.print(
                f"  [red]✗ Failed to create group: {_summarize_az_error(result)}[/red]"
            )
            return None

        group_data = json.loads(result.stdout)
        group_id = group_data.get("id", "")
        if not group_id:
            console.print("  [red]✗ Group created but no ID returned.[/red]")
            return None

        console.print(
            f"  [green]✓[/green] Created M365 Group "
            f"[cyan]{display_name}[/cyan] ({group_id[:8]}…)"
        )

        # Team-enable the group — requires a short delay for provisioning
        console.print("  [dim]Enabling Teams on the group…[/dim]")
        for attempt in range(5):
            time.sleep(3)
            team_result = _run_az_command(
                [
                    "rest",
                    "--method", "PUT",
                    "--url",
                    f"https://graph.microsoft.com/v1.0/groups/{group_id}/team",
                    "--body", "{}",
                ]
            )
            if team_result.returncode == 0:
                console.print("  [green]✓[/green] Team enabled on group.")
                _add_az_login_user_to_group(group_id)
                return group_id
            if attempt < 4:
                console.print(
                    f"  [dim]  Waiting for provisioning "
                    f"(attempt {attempt + 2}/5)…[/dim]"
                )

        # Group created but team-enable failed
        console.print(
            "  [yellow]⚠ Group created but could not enable Teams. "
            "You may need to team-enable it manually in Teams admin.[/yellow]"
        )
        return group_id

    except Exception as exc:
        console.print(f"  [red]✗ Error creating group: {exc}[/red]")

    return None


def _ask_team_id(
    tenant_id: str,
    theme: str = "healthcare",
    *,
    owner_upns: list[str] | None = None,
    graph_client: Any | None = None,
) -> str:
    """Let the user select, create, or enter a team_id for Teams seeding."""
    console.print("\n  [dim]Discovering M365 Groups / Teams in tenant…[/dim]")
    teams = _discover_teams(tenant_id, graph_client=graph_client)
    themed_name = THEMED_TEAM_NAMES.get(theme, "Contoso Demo Team")

    if teams:
        console.print(f"  [dim]Found {len(teams)} group(s).[/dim]")
        console.print("\n  [bold]Available M365 Groups (potential Teams):[/bold]")
        for i, t in enumerate(teams[:15], 1):
            console.print(f"    {i:2d}. {t['displayName']}  [dim]({t['id'][:8]}…)[/dim]")
        new_idx = min(len(teams), 15) + 1
        console.print(
            f"    {new_idx:2d}. [green]✚ Create new:[/green] "
            f"[cyan]{themed_name}[/cyan]  [dim](on-theme)[/dim]"
        )

        raw = Prompt.ask(
            "\n  Select a number, or paste a Team ID GUID",
            default=str(new_idx),
        )

        stripped = raw.strip()
        if stripped.isdigit():
            idx = int(stripped)
            if idx == new_idx:
                # Create a new themed group
                custom = Prompt.ask(
                    "  Group name",
                    default=themed_name,
                )
                # Check if a group with this name already exists
                existing = next(
                    (t for t in teams if t["displayName"].lower() == custom.lower()),
                    None,
                )
                if existing:
                    console.print(
                        f"  [yellow]⚠ A group named [cyan]{existing['displayName']}[/cyan] "
                        f"already exists ({existing['id'][:8]}…).[/yellow]"
                    )
                    if Confirm.ask("  Use the existing group instead?", default=True):
                        console.print(
                            f"  → Using: [cyan]{existing['displayName']}[/cyan] "
                            f"({existing['id']})"
                        )
                        return existing["id"]
                gid = _create_team_group(
                    custom,
                    owner_upns=owner_upns,
                    graph_client=graph_client,
                )
                if gid:
                    return gid
                # fall through to manual entry on failure
            else:
                real_idx = idx - 1
                if 0 <= real_idx < len(teams):
                    selected = teams[real_idx]
                    console.print(
                        f"  → Using: [cyan]{selected['displayName']}[/cyan] ({selected['id']})"
                    )
                    return selected["id"]

        if GUID_RE.match(stripped):
            return stripped
    else:
        console.print(
            "  [yellow]⚠ Could not discover existing Teams/Groups.[/yellow]"
        )
        if Confirm.ask(
            f"  Create a new group [cyan]{themed_name}[/cyan]?",
            default=True,
        ):
            custom = Prompt.ask("  Group name", default=themed_name)
            gid = _create_team_group(
                custom,
                owner_upns=owner_upns,
                graph_client=graph_client,
            )
            if gid:
                return gid

    # Fall back to manual entry
    while True:
        team_id = Prompt.ask("  Team ID (GUID)")
        if team_id and GUID_RE.match(team_id.strip()):
            return team_id.strip()
        console.print("  [red]Please enter a valid GUID.[/red]")


def _ask_group_id(
    tenant_id: str,
    team_id: str = "",
    theme: str = "healthcare",
    *,
    owner_upns: list[str] | None = None,
    graph_client: Any | None = None,
) -> str:
    """Let the user select, create, or enter a group_id for Planner.

    If a team_id was already chosen, offer to reuse the same group.
    """
    if team_id:
        if Confirm.ask(
            f"  Use the same group as Teams ({team_id[:8]}…) for Planner?",
            default=True,
        ):
            return team_id

    console.print("\n  [dim]Discovering M365 Groups for Planner…[/dim]")
    groups = _discover_teams(tenant_id, graph_client=graph_client)
    themed_name = THEMED_TEAM_NAMES.get(theme, "Contoso Demo Team")

    if groups:
        console.print(f"  [dim]Found {len(groups)} group(s).[/dim]")
        for i, g in enumerate(groups[:15], 1):
            console.print(f"    {i:2d}. {g['displayName']}  [dim]({g['id'][:8]}…)[/dim]")
        new_idx = min(len(groups), 15) + 1
        console.print(
            f"    {new_idx:2d}. [green]✚ Create new:[/green] "
            f"[cyan]{themed_name}[/cyan]  [dim](on-theme)[/dim]"
        )

        raw = Prompt.ask(
            "\n  Select a number, or paste a Group ID GUID",
            default=str(new_idx),
        )

        stripped = raw.strip()
        if stripped.isdigit():
            idx = int(stripped)
            if idx == new_idx:
                custom = Prompt.ask("  Group name", default=themed_name)
                # Check if a group with this name already exists
                existing = next(
                    (g for g in groups if g["displayName"].lower() == custom.lower()),
                    None,
                )
                if existing:
                    console.print(
                        f"  [yellow]⚠ A group named [cyan]{existing['displayName']}[/cyan] "
                        f"already exists ({existing['id'][:8]}…).[/yellow]"
                    )
                    if Confirm.ask("  Use the existing group instead?", default=True):
                        console.print(
                            f"  → Using: [cyan]{existing['displayName']}[/cyan] "
                            f"({existing['id']})"
                        )
                        return existing["id"]
                gid = _create_team_group(
                    custom,
                    owner_upns=owner_upns,
                    graph_client=graph_client,
                )
                if gid:
                    return gid
            else:
                real_idx = idx - 1
                if 0 <= real_idx < len(groups):
                    selected = groups[real_idx]
                    console.print(
                        f"  → Using: [cyan]{selected['displayName']}[/cyan] ({selected['id']})"
                    )
                    return selected["id"]

        if GUID_RE.match(stripped):
            return stripped
    else:
        if Confirm.ask(
            f"  Create a new group [cyan]{themed_name}[/cyan]?",
            default=True,
        ):
            custom = Prompt.ask("  Group name", default=themed_name)
            gid = _create_team_group(
                custom,
                owner_upns=owner_upns,
                graph_client=graph_client,
            )
            if gid:
                return gid

    # Fall back to manual entry
    while True:
        group_id = Prompt.ask("  Group ID (GUID)")
        if group_id and GUID_RE.match(group_id.strip()):
            return group_id.strip()
        console.print("  [red]Please enter a valid GUID.[/red]")


def _generate_config(
    *,
    tenant_id: str,
    client_id: str,
    secret_env: str,
    theme: str,
    run_id: str,
    users: list[dict[str, str]],
    sections: dict[str, bool],
    team_id: str = "",
    group_id: str = "",
) -> str:
    """Render a complete ``seed-config.yaml`` from wizard answers.

    When *team_id* or *group_id* are supplied the Teams / Planner sections
    are wired up automatically.  Theme content (channels, chats, sites,
    plans) is pulled from ``themes.json`` so the generated config is
    ready to seed without manual editing.
    """
    from m365seed.theme_content import (
        get_mail_threads,
        get_calendar_events,
        get_teams_channels,
        get_chat_conversations,
        get_sharepoint_sites,
        get_planner_plans,
    )

    lines: list[str] = []
    a = lines.append

    a("# ─────────────────────────────────────────────────────────────")
    a("# M365 Seed — Generated Configuration")
    a("# Generated by: m365seed setup")
    a("# ─────────────────────────────────────────────────────────────")
    a("")
    a("tenant:")
    a(f'  tenant_id: "{tenant_id}"')
    a(f'  authority: "https://login.microsoftonline.com/{tenant_id}"')
    a("")
    a("auth:")
    a('  mode: "client_secret"')
    a(f'  client_id: "{client_id}"')
    a(f'  client_secret_env: "{secret_env}"')
    a("")
    a("targets:")
    a("  users:")
    for u in users:
        a(f'    - upn: "{u["upn"]}"')
        a(f'      role: "{u["role"]}"')
    a("")
    a("content:")
    a(f'  theme: "{theme}"')
    a(f'  run_id: "{run_id}"')
    a("")

    # ── Mail ────────────────────────────────────────────────
    a("mail:")
    if sections.get("mail"):
        mail_threads = get_mail_threads(theme)
        a("  threads:")
        for mt in mail_threads:
            a(f'    - thread_id: "{mt["thread_id"]}"')
            a(f'      subject: "{mt.get("subject", mt["thread_id"])}"')
            a("      participants:")
            # Distribute users across threads
            for u in users[:3]:
                a(f'        - "{u["upn"]}"')
            a("      messages: 6")
            a("      include_attachments: true")
            a("")
    else:
        a("  threads: []")
    a("")

    # ── Files ───────────────────────────────────────────────
    a("files:")
    a("  oneDrive:")
    a(f"    enabled: {str(sections.get('files', True)).lower()}")
    if sections.get("files") and users:
        a(f'    target_user: "{users[0]["upn"]}"')
    else:
        a('    target_user: ""')
    a("    folders:")
    a('      - "Clinical Ops"')
    a('      - "Care Coordination"')
    a('      - "Compliance"')
    a('      - "Quality Improvement"')
    a("  sharePoint:")
    a("    enabled: false")
    a('    site_id: ""')
    a('    drive_id: ""')
    a("")

    # ── Calendar ────────────────────────────────────────────
    a("calendar:")
    a(f"  enabled: {str(sections.get('calendar', True)).lower()}")
    if sections.get("calendar"):
        cal_events = get_calendar_events(theme)
        a("  events:")
        for i, evt in enumerate(cal_events):
            a(f'    - event_id: "{evt["event_id"]}"')
            a(f'      subject: "{evt.get("subject", evt["event_id"])}"')
            if users:
                a(f'      organizer: "{users[i % len(users)]["upn"]}"')
            a("      attendees:")
            # Assign 2-3 attendees from the user pool
            attendees = [u for j, u in enumerate(users) if j != (i % len(users))]
            for att in attendees[:3]:
                a(f'        - "{att["upn"]}"')
            a("      duration_minutes: 30")
            a("      is_online_meeting: true")
            a("")
    a("")

    # ── Teams ───────────────────────────────────────────────
    a("teams:")
    a(f"  enabled: {str(sections.get('teams', False)).lower()}")
    a(f'  team_id: "{team_id}"')
    if sections.get("teams"):
        channels = get_teams_channels(theme)
        a("  channels:")
        for ch in channels:
            a(f'    - channel_id: "{ch["channel_id"]}"')
            a(f'      display_name: "{ch["display_name"]}"')
            if ch.get("description"):
                # Escape quotes in description
                desc = ch["description"].replace('"', '\\"')
                a(f'      description: "{desc}"')
            if ch.get("posts"):
                a("      posts:")
                for post in ch["posts"]:
                    escaped = post.replace('"', '\\"')
                    a(f'        - message: "{escaped}"')
            a("")
    else:
        a("  channels: []")
    a("")

    # ── Chats ───────────────────────────────────────────────
    a("chats:")
    a(f"  enabled: {str(sections.get('chats', False)).lower()}")
    if sections.get("chats") and users:
        conversations = get_chat_conversations(theme)
        a("  conversations:")
        for conv in conversations:
            a(f'    - conversation_id: "{conv["conversation_id"]}"')
            ctype = conv.get("type", "group")
            a(f'      type: "{ctype}"')
            if conv.get("topic"):
                topic = conv["topic"].replace('"', '\\"')
                a(f'      topic: "{topic}"')
            # Assign members from the user pool
            a("      members:")
            if ctype == "oneOnOne":
                for u in users[:2]:
                    a(f'        - "{u["upn"]}"')
            else:
                for u in users[:min(4, len(users))]:
                    a(f'        - "{u["upn"]}"')
            if conv.get("messages"):
                a("      messages:")
                # Assign senders round-robin from conversation members
                if ctype == "oneOnOne":
                    members_pool = [u["upn"] for u in users[:2]]
                else:
                    members_pool = [u["upn"] for u in users[:min(4, len(users))]]
                for mi, msg in enumerate(conv["messages"]):
                    sender = members_pool[mi % len(members_pool)] if members_pool else ""
                    escaped = msg.replace('"', '\\"')
                    a(f'        - sender: "{sender}"')
                    a(f'          text: "{escaped}"')
            a("")
    else:
        a("  conversations: []")
    a("")

    # ── SharePoint ──────────────────────────────────────────
    a("sharepoint:")
    a(f"  enabled: {str(sections.get('sharepoint', False)).lower()}")
    if sections.get("sharepoint") and users:
        a(f'  owner: "{users[0]["upn"]}"')
        sites = get_sharepoint_sites(theme)
        a("  sites:")
        for site in sites:
            a(f'    - display_name: "{site["display_name"]}"')
            if site.get("description"):
                desc = site["description"].replace('"', '\\"')
                a(f'      description: "{desc}"')
            # Pages and documents are enriched from theme at runtime
            # — we only need display_name in config
            a("")
    else:
        a('  owner: ""')
        a("  sites: []")
    a("")

    # ── Planner ─────────────────────────────────────────────
    a("planner:")
    a(f"  enabled: {str(sections.get('planner', False)).lower()}")
    a(f'  group_id: "{group_id}"')
    if sections.get("planner"):
        plans = get_planner_plans(theme)
        a("  plans:")
        for plan in plans:
            a(f'    - title: "{plan["title"]}"')
            # Buckets and tasks are enriched from theme at runtime
            a("")
    else:
        a("  plans: []")
    a("")

    return "\n".join(lines)


# ── Main wizard ──────────────────────────────────────────────


def run_setup(config_path: str = "seed-config.yaml") -> None:
    """Interactive setup wizard entry point."""
    # Proactively clear any corrupt MSAL cache so all az CLI calls succeed.
    from m365seed.register import _ensure_msal_cache_healthy
    _ensure_msal_cache_healthy()

    console.print(
        Panel.fit(
            "[bold cyan]M365 Demo Tenant Seeding Tool — Setup Wizard[/bold cyan]\n\n"
            "This wizard will walk you through configuring your demo tenant.\n"
            "All values are saved to [bold]seed-config.yaml[/bold].",
            border_style="cyan",
        )
    )

    # ── Step 1: Tenant ──────────────────────────────────────
    console.print("\n[bold]Step 1 — Tenant[/bold]")
    tenant_id = _ask_guid("Tenant ID (GUID)", env_var="M365SEED_TENANT_ID")

    # Derive domain hint (prefer real tenant default domain from az context)
    inferred_domain = _guess_tenant_domain(tenant_id)
    domain_default = inferred_domain or f"M365x{tenant_id[:6]}.onmicrosoft.com"
    domain = Prompt.ask("  Tenant domain", default=domain_default)

    # ── Step 2: App Registration ────────────────────────────
    console.print("\n[bold]Step 2 — App Registration[/bold]")
    auto_register = Confirm.ask(
        "  Register a new app automatically via Azure CLI?", default=False
    )

    secret_env = "M365SEED_CLIENT_SECRET"
    if auto_register:
        from m365seed.register import register_app

        include_sp = sections_preview_sp = Confirm.ask(
            "    Include SharePoint + Planner permissions?", default=True
        )
        include_teams_perms = Confirm.ask(
            "    Include Teams permissions (beta)?", default=False
        )
        reg = register_app(
            tenant_id,
            include_teams=include_teams_perms,
            include_sharepoint_planner=include_sp,
        )
        if reg:
            from m365seed.register import _format_env_export

            client_id = reg["client_id"]
            # Set secret in current process env for downstream steps
            os.environ[secret_env] = reg["client_secret"]
            console.print(
                f"\n  [green]✓[/green] App registered — Client ID: [cyan]{client_id}[/cyan]"
            )
            console.print(
                f"  [green]✓[/green] Client secret saved to ${secret_env} for this session."
            )
            console.print(
                f"  [green]✓[/green] Export command: [dim]{_format_env_export(secret_env, reg['client_secret'])}[/dim]"
            )
            console.print(
                "  [yellow]⚠ Save this secret now — it cannot be retrieved later.[/yellow]"
            )
        else:
            console.print(
                "  [yellow]Auto-registration failed — falling back to manual entry.[/yellow]"
            )
            client_id = _ask_guid("Client (Application) ID", env_var="M365SEED_CLIENT_ID")
    else:
        client_id = _ask_guid("Client (Application) ID", env_var="M365SEED_CLIENT_ID")
        secret_env = Prompt.ask(
            "  Client secret env var name",
            default="M365SEED_CLIENT_SECRET",
        )

    # Verify secret is set
    if not os.environ.get(secret_env):
        console.print(
            f"  [yellow]⚠ Environment variable ${secret_env} is not set.[/yellow]\n"
            f"  Set it before running seed commands:\n"
            f"    export {secret_env}=\"your-client-secret\""
        )

    setup_graph_client = _build_setup_graph_client(tenant_id, client_id, secret_env)
    verified_domains = _get_verified_tenant_domains(setup_graph_client)
    resolved_domain, domain_was_corrected = _resolve_tenant_domain(domain, verified_domains)
    if domain_was_corrected:
        domain_matches_default = domain.strip().lower() == domain_default.lower()
        console.print(
            f"  [yellow]⚠ Tenant domain [cyan]{domain}[/cyan] is not verified for this tenant.[/yellow]"
        )
        if domain_matches_default:
            domain = resolved_domain
            console.print(
                f"  [green]✓[/green] Using verified domain [cyan]{domain}[/cyan] instead."
            )
        else:
            preview = ", ".join(verified_domains[:3])
            if preview:
                console.print(f"  [dim]Verified domains: {preview}[/dim]")
            domain = Prompt.ask("  Tenant domain", default=resolved_domain).strip().lower()
    else:
        domain = resolved_domain

    # ── Step 3: Theme ───────────────────────────────────────
    console.print("\n[bold]Step 3 — Content Theme[/bold]")
    theme = _ask_choice("Select theme", THEMES, default="healthcare")
    console.print(f"  → Theme: [cyan]{theme}[/cyan]")

    # ── Step 4: Run ID ──────────────────────────────────────
    console.print("\n[bold]Step 4 — Run ID[/bold]")
    run_id = Prompt.ask(
        "  Run ID (used for idempotency & cleanup)",
        default="hls-demo-001",
    )

    # ── Step 5: Users ───────────────────────────────────────
    console.print("\n[bold]Step 5 — Demo Users[/bold]")
    users = _ask_users(domain, graph_client=setup_graph_client)

    # ── Step 6: User Sign-In Passwords (optional) ───────────
    console.print("\n[bold]Step 6 — User Sign-In Passwords (optional)[/bold]")
    if Confirm.ask("  Reset passwords for these demo users now?", default=False):
        # Pre-validate: check which users actually exist
        console.print("  [dim]Verifying users exist in tenant…[/dim]")
        verified: list[dict[str, str]] = []
        missing: list[dict[str, str]] = []
        for u in users:
            if _check_user_exists(u["upn"], graph_client=setup_graph_client):
                verified.append(u)
            else:
                missing.append(u)
                console.print(
                    f"    [red]✗[/red] {u['upn']} — not found in tenant"
                )

        if not missing and verified:
            console.print(
                f"  [green]✓[/green] All {len(verified)} user(s) verified."
            )
        elif missing and verified:
            console.print(
                f"\n  [yellow]⚠ {len(missing)} user(s) not found. "
                f"Password reset will apply to the {len(verified)} existing user(s).[/yellow]"
            )
        elif missing:
            console.print(
                "\n  [yellow]⚠ None of the selected users exist in the tenant yet.[/yellow]"
            )

        create_missing = False
        if missing:
            console.print(
                "  [dim]Setup can create missing Entra ID users now, but mailboxes, "
                "OneDrive, Teams, and Planner still require Microsoft 365 licensing "
                "and provisioning.[/dim]"
            )
            create_missing = Confirm.ask(
                "  Create missing Entra ID users now?",
                default=not verified,
            )

        if verified or create_missing:
            while True:
                temp_password = Prompt.ask(
                    "  Temporary password to set for created or selected demo users",
                    password=True,
                )
                if len(temp_password) >= 8:
                    break
                console.print("  [red]Password must be at least 8 characters.[/red]")

            force_change = Confirm.ask(
                "  Force users to change password at next sign-in?",
                default=False,
            )

            if create_missing:
                console.print("  Creating missing users…")
                created_success, created_failed, created_users = _create_missing_demo_users(
                    missing,
                    temp_password,
                    force_change_next_sign_in=force_change,
                    graph_client=setup_graph_client,
                )
                console.print(
                    f"  [green]User creation:[/green] {created_success} created, {created_failed} failed."
                )
                verified.extend(created_users)

            existing_users = [u for u in verified if u not in missing]
            if existing_users:
                console.print("  Applying password reset…")
                success, failed = _reset_demo_user_passwords(
                    existing_users,
                    temp_password,
                    force_change_next_sign_in=force_change,
                    graph_client=setup_graph_client,
                )
                console.print(f"  [green]Password reset:[/green] {success} updated, {failed} failed.")
        else:
            console.print(
                "  [dim]Skipped password reset and user creation because none of the selected users exist yet.[/dim]"
            )
    else:
        console.print("  [dim]Skipped password reset. Existing tenant passwords remain unchanged.[/dim]")

    # ── Step 7: Content modules ─────────────────────────────
    console.print("\n[bold]Step 7 — Content Modules[/bold]")
    console.print("  Select which content types to seed:")
    sections: dict[str, bool] = {}
    for key, label, default in CONTENT_SECTIONS:
        sections[key] = Confirm.ask(f"    {label}?", default=default)

    if sections.get("files"):
        console.print(
            "\n  [yellow]⚠ OneDrive note:[/yellow] Each user's OneDrive must be accessed at least once\n"
            "  before files can be uploaded. If seed-files returns a 404, sign in as the\n"
            "  target user at [bold]https://<tenant>-my.sharepoint.com/[/bold] to provision their drive."
        )

    # ── Step 8: Teams / Planner IDs ─────────────────────────
    team_id = ""
    group_id = ""

    if sections.get("teams") or sections.get("chats"):
        console.print("\n[bold]Step 8 — Teams Configuration[/bold]")
        console.print(
            "  [dim]Teams channels and chats require a Team ID "
            "(the M365 Group backing the team).[/dim]"
        )
        team_id = _ask_team_id(
            tenant_id,
            theme,
            owner_upns=[u["upn"] for u in users],
            graph_client=setup_graph_client,
        )

    if sections.get("planner"):
        if not team_id:
            console.print("\n[bold]Step 8 — Planner Configuration[/bold]")
        else:
            console.print("\n[bold]Planner Configuration[/bold]")
        console.print(
            "  [dim]Planner requires an M365 Group to own the plans.[/dim]"
        )
        group_id = _ask_group_id(
            tenant_id,
            team_id,
            theme,
            owner_upns=[u["upn"] for u in users],
            graph_client=setup_graph_client,
        )

    # ── Generate config ─────────────────────────────────────
    yaml_str = _generate_config(
        tenant_id=tenant_id,
        client_id=client_id,
        secret_env=secret_env,
        theme=theme,
        run_id=run_id,
        users=users,
        sections=sections,
        team_id=team_id,
        group_id=group_id,
    )

    # ── Write config ────────────────────────────────────────
    target = Path(config_path)
    if target.exists():
        overwrite = Confirm.ask(
            f"\n  [yellow]{target} already exists. Overwrite?[/yellow]",
            default=False,
        )
        if not overwrite:
            backup = target.with_suffix(".yaml.bak")
            shutil.copy2(target, backup)
            console.print(f"  Backed up existing config to {backup}")

    target.write_text(yaml_str, encoding="utf-8")
    console.print(f"\n  [green]✓[/green] Wrote [bold]{target}[/bold]")

    # ── Summary ─────────────────────────────────────────────
    table = Table(title="Configuration Summary", show_header=True)
    table.add_column("Setting", style="bold")
    table.add_column("Value")
    table.add_row("Tenant ID", tenant_id)
    table.add_row("Domain", domain)
    table.add_row("Client ID", client_id)
    table.add_row("Secret Env", secret_env)
    table.add_row("Theme", theme)
    table.add_row("Run ID", run_id)
    table.add_row("Users", str(len(users)))
    enabled = [k for k, v in sections.items() if v]
    table.add_row("Content", ", ".join(enabled) or "none")
    console.print()
    console.print(table)

    # ── Next steps ──────────────────────────────────────────
    console.print("\n[bold]Next Steps[/bold]")

    # Offer to validate
    if Confirm.ask("  Run validation now?", default=True):
        console.print()
        _run_child_command(["m365seed", "validate", "-c", str(target)])

    # Offer dry run
    if Confirm.ask("  Run a dry-run seed?", default=True):
        cmd = ["m365seed", "seed-all", "-c", str(target), "--dry-run"]
        if sections.get("teams") or sections.get("chats"):
            cmd.append("--enable-beta-teams")
        console.print()
        _run_child_command(cmd)

    # Offer live run
    if Confirm.ask("  Run the live seed now?", default=False):
        cmd = ["m365seed", "seed-all", "-c", str(target)]
        if sections.get("teams") or sections.get("chats"):
            cmd.append("--enable-beta-teams")
        console.print()
        _run_child_command(cmd)

    console.print(
        "\n[bold green]Setup complete![/bold green] "
        "You can rerun any command individually — see [bold]m365seed --help[/bold]."
    )


def _run_child_command(cmd: list[str]) -> None:
    """Execute an m365seed sub-command, showing output in real time."""
    import subprocess
    import sys

    console.print(f"  [dim]$ {' '.join(cmd)}[/dim]\n")
    try:
        subprocess.run(cmd, check=False)
    except FileNotFoundError:
        # Fall back to module invocation
        subprocess.run(
            [sys.executable, "-m", "m365seed.cli"] + cmd[1:],
            check=False,
        )
