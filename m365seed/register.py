"""Automated Entra ID App Registration via Azure CLI.

Creates the app registration, assigns Graph permissions, generates
a client secret, and grants admin consent — all via ``az`` commands
with device-code login.
"""

from __future__ import annotations

import json
import shutil
import subprocess
import sys
from typing import Optional

from rich.console import Console
from rich.panel import Panel
from rich.prompt import Confirm, Prompt
from rich.table import Table

console = Console()

# ── Microsoft Graph resource ID ──────────────────────────────
GRAPH_API_ID = "00000003-0000-0000-c000-000000000000"

# ── Application-type permission GUIDs (stable, from MS docs) ─
CORE_PERMISSIONS: dict[str, str] = {
    "User.Read.All": "df021288-bdef-4463-88db-98f22de89214",
    "Organization.Read.All": "498476ce-e0fe-48b0-b801-37ba7e2685c6",
    "Mail.Send": "b633e1c5-b582-4048-a93e-9f11b44c7e96",
    "Mail.ReadWrite": "e2a3a72e-5f79-4c64-b1b1-878b674786c9",
    "Files.ReadWrite.All": "75359482-378d-4052-8f01-80520e7db3cd",
    "Sites.ReadWrite.All": "9492366f-7969-46a4-8d15-ed1a20078fff",
    "Calendars.ReadWrite": "ef54d2bf-783f-4e0f-bca1-3210c0444d99",
    "OnlineMeetings.ReadWrite.All": "b8bb2037-6e08-44ac-a4ea-4674e010e2a4",
}

EXTENDED_PERMISSIONS: dict[str, str] = {
    "Group.ReadWrite.All": "62a82d76-70ea-41e2-9197-370581804d09",
    "Sites.Manage.All": "0c0bf378-bf22-4f51-be20-13571c76a2c4",
    "Tasks.ReadWrite.All": "44e666d1-d276-445b-a5fc-8815eeb81d55",
}

TEAMS_PERMISSIONS: dict[str, str] = {
    "Chat.Create": "d9c48af6-9ad9-47ad-82c3-63757137b9af",
    "Chat.ReadWrite.All": "294ce7c9-31ba-490a-ad7d-97a7d075e4ed",
    "Channel.Create": "f3a65bd4-b703-46df-8f7e-0174571a4049",
    "Channel.Delete.All": "6a118a39-1227-45d4-b534-b9ce547f9313",
    "ChannelMessage.Send": "7ab1d787-bae7-4d5d-8b12-37d448e0bdb4",
}


# ── Helpers ──────────────────────────────────────────────────


def _az(*args: str, capture: bool = True) -> subprocess.CompletedProcess:
    """Run an ``az`` CLI command, returning parsed output."""
    cmd = ["az", *args]
    if capture:
        cmd.extend(["--output", "json"])
    result = subprocess.run(
        cmd,
        capture_output=capture,
        text=True,
    )
    return result


def _az_json(*args: str) -> dict | list | None:
    """Run ``az`` and parse JSON output.  Returns None on failure."""
    r = _az(*args, capture=True)
    if r.returncode != 0:
        return None
    try:
        return json.loads(r.stdout)
    except (json.JSONDecodeError, TypeError):
        return None


def _check_az_cli() -> bool:
    """Return True if Azure CLI is available."""
    return shutil.which("az") is not None


def _is_logged_in(tenant_id: str) -> bool:
    """Return True if already logged into the target tenant."""
    acct = _az_json("account", "show")
    if acct and isinstance(acct, dict):
        return acct.get("tenantId", "").lower() == tenant_id.lower()
    return False


# ── Public API ───────────────────────────────────────────────


def register_app(
    tenant_id: str,
    *,
    app_name: str = "M365 Demo Seed Tool",
    include_teams: bool = False,
    include_sharepoint_planner: bool = False,
    secret_years: int = 1,
) -> dict[str, str] | None:
    """Create an Entra ID App Registration via Azure CLI.

    Returns a dict with ``client_id``, ``client_secret``, ``tenant_id``
    on success, or ``None`` on failure.
    """

    # ── 1. Verify Azure CLI ─────────────────────────────────
    if not _check_az_cli():
        console.print(
            "[red]✗ Azure CLI (az) is not installed.[/red]\n"
            "  Install it: https://aka.ms/installazurecli\n"
            "  Or use the dev container (Azure CLI is pre-installed)."
        )
        return None

    # ── 2. Login via device code ────────────────────────────
    if not _is_logged_in(tenant_id):
        console.print(
            "\n  [bold]Signing in to Azure…[/bold]\n"
            "  A device-code prompt will appear — follow the instructions\n"
            "  to authenticate as a [bold]Global Administrator[/bold] of the demo tenant.\n"
        )
        result = _az(
            "login",
            "--tenant", tenant_id,
            "--use-device-code",
            "--allow-no-subscriptions",
            capture=False,
        )
        if result.returncode != 0:
            console.print("[red]✗ Azure login failed.[/red]")
            return None
        console.print("[green]✓[/green] Logged in successfully.")
    else:
        console.print(f"[green]✓[/green] Already logged into tenant {tenant_id}.")

    # ── 3. Create the app registration ──────────────────────
    console.print(f"\n  Creating app registration: [cyan]{app_name}[/cyan]…")
    app = _az_json(
        "ad", "app", "create",
        "--display-name", app_name,
        "--sign-in-audience", "AzureADMyOrg",
    )
    if not app or not isinstance(app, dict):
        console.print("[red]✗ Failed to create app registration.[/red]")
        return None

    client_id = app["appId"]
    object_id = app["id"]
    console.print(f"  [green]✓[/green] App created — Client ID: [cyan]{client_id}[/cyan]")

    # ── 4. Build permission set ─────────────────────────────
    permissions = dict(CORE_PERMISSIONS)
    if include_sharepoint_planner:
        permissions.update(EXTENDED_PERMISSIONS)
    if include_teams:
        permissions.update(TEAMS_PERMISSIONS)

    # ── 5. Add Graph permissions ────────────────────────────
    console.print(f"\n  Adding {len(permissions)} Graph API permissions…")
    perm_string = " ".join(f"{guid}=Role" for guid in permissions.values())
    r = _az(
        "ad", "app", "permission", "add",
        "--id", client_id,
        "--api", GRAPH_API_ID,
        "--api-permissions", perm_string,
    )
    if r.returncode != 0:
        console.print(f"[yellow]⚠ Permission add returned non-zero — continuing…[/yellow]")
        if r.stderr:
            console.print(f"  [dim]{r.stderr.strip()[:200]}[/dim]")
    else:
        for name in permissions:
            console.print(f"    [green]✓[/green] {name}")

    # ── 6. Create service principal ─────────────────────────
    console.print("\n  Creating service principal…")
    sp = _az_json("ad", "sp", "create", "--id", client_id)
    if sp is None:
        # May already exist
        console.print("  [dim](service principal may already exist — OK)[/dim]")
    else:
        console.print("  [green]✓[/green] Service principal created.")

    # ── 7. Generate client secret ───────────────────────────
    console.print(f"\n  Generating client secret (valid for {secret_years} year(s))…")
    cred = _az_json(
        "ad", "app", "credential", "reset",
        "--id", client_id,
        "--years", str(secret_years),
        "--display-name", "m365seed-auto",
    )
    if not cred or not isinstance(cred, dict):
        console.print("[red]✗ Failed to generate client secret.[/red]")
        return None

    client_secret = cred.get("password", "")
    console.print("  [green]✓[/green] Client secret generated.")

    # ── 8. Grant admin consent ──────────────────────────────
    console.print("\n  Granting admin consent…")
    # Small delay to allow propagation
    import time
    time.sleep(3)

    r = _az(
        "ad", "app", "permission", "admin-consent",
        "--id", client_id,
    )
    if r.returncode != 0:
        console.print(
            "[yellow]⚠ Admin consent may need to be granted manually.[/yellow]\n"
            f"  Run: az ad app permission admin-consent --id {client_id}\n"
            "  Or grant consent in the Azure Portal → API permissions."
        )
    else:
        console.print("  [green]✓[/green] Admin consent granted.")

    # ── 9. Summary ──────────────────────────────────────────
    result = {
        "client_id": client_id,
        "client_secret": client_secret,
        "tenant_id": tenant_id,
        "object_id": object_id,
    }

    console.print()
    console.print(
        Panel(
            f"[bold green]App Registration Complete[/bold green]\n\n"
            f"  Client ID:     [cyan]{client_id}[/cyan]\n"
            f"  Tenant ID:     [cyan]{tenant_id}[/cyan]\n"
            f"  Object ID:     [dim]{object_id}[/dim]\n\n"
            f"  Client Secret: [bold yellow](saved — set as env var below)[/bold yellow]\n\n"
            f"  Permissions:   {len(permissions)} Graph API (Application)\n"
            f"  Consent:       {'Granted' if r.returncode == 0 else 'Manual step required'}",
            title="Registration Summary",
            border_style="green",
        )
    )

    return result


def run_registration_wizard(tenant_id: str | None = None) -> dict[str, str] | None:
    """Interactive wrapper around ``register_app`` for the CLI."""
    console.print(
        Panel.fit(
            "[bold cyan]Entra ID App Registration — Automated Setup[/bold cyan]\n\n"
            "This will create an app registration in your demo tenant\n"
            "using Azure CLI with device-code authentication.\n\n"
            "[dim]Requires: Global Administrator role on the tenant.[/dim]",
            border_style="cyan",
        )
    )

    import os
    import re

    guid_re = re.compile(
        r"^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$",
        re.IGNORECASE,
    )

    # Tenant ID
    if not tenant_id:
        default = os.environ.get("M365SEED_TENANT_ID", "")
        while True:
            tenant_id = Prompt.ask("  Tenant ID (GUID)", default=default or None)
            if tenant_id and guid_re.match(tenant_id.strip()):
                tenant_id = tenant_id.strip()
                break
            console.print("  [red]Enter a valid GUID.[/red]")

    # App name
    app_name = Prompt.ask("  App display name", default="M365 Demo Seed Tool")

    # Which permissions?
    include_sp = Confirm.ask("  Include SharePoint + Planner permissions?", default=True)
    include_teams = Confirm.ask("  Include Teams permissions (beta)?", default=False)

    console.print()
    if not Confirm.ask("  Proceed with app registration?", default=True):
        console.print("  [dim]Cancelled.[/dim]")
        return None

    result = register_app(
        tenant_id,
        app_name=app_name,
        include_teams=include_teams,
        include_sharepoint_planner=include_sp,
    )

    if result:
        # Offer to set env var for current session
        secret = result["client_secret"]
        console.print(
            "\n  [bold]Set the client secret in your environment:[/bold]\n"
        )
        if sys.platform == "win32":
            console.print(f'    $env:M365SEED_CLIENT_SECRET = "{secret}"')
        else:
            console.print(f'    export M365SEED_CLIENT_SECRET="{secret}"')
        console.print(
            "\n  [yellow]⚠ Save this secret now — it cannot be retrieved later.[/yellow]"
        )

    return result
