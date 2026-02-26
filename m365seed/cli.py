"""CLI entry point — ``m365seed`` commands."""

from __future__ import annotations

import json
import logging
import sys
from pathlib import Path
from typing import Optional

import typer
from rich.console import Console
from rich.logging import RichHandler
from rich.table import Table

from m365seed.config import load_config, get_run_id, get_theme, get_users

app = typer.Typer(
    name="m365seed",
    help=(
        "Safe, idempotent seeding tool for Microsoft 365 demo tenants "
        "with synthetic, theme-aware HLS content."
    ),
    add_completion=False,
)

console = Console()

# ---------------------------------------------------------------------------
# Shared options
# ---------------------------------------------------------------------------

CONFIG_OPT = typer.Option(
    "seed-config.yaml",
    "--config",
    "-c",
    help="Path to the seed-config YAML file.",
)

DRY_RUN_OPT = typer.Option(
    False,
    "--dry-run",
    help="Print intended actions without modifying the tenant.",
)

LOG_FILE_OPT = typer.Option(
    None,
    "--log-file",
    help="Optional path for JSONL structured log output.",
)

VERBOSE_OPT = typer.Option(
    False,
    "--verbose",
    "-v",
    help="Enable debug-level logging.",
)


def _setup_logging(verbose: bool, log_file: Optional[str] = None) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    handlers: list[logging.Handler] = [
        RichHandler(console=console, show_path=False, rich_tracebacks=True)
    ]
    if log_file:
        fh = logging.FileHandler(log_file, mode="a", encoding="utf-8")
        fh.setFormatter(logging.Formatter("%(message)s"))

        class JSONLFilter(logging.Filter):
            def filter(self, record: logging.LogRecord) -> bool:
                record.msg = json.dumps(
                    {
                        "ts": record.created,
                        "level": record.levelname,
                        "name": record.name,
                        "msg": record.getMessage(),
                    }
                )
                return True

        fh.addFilter(JSONLFilter())
        handlers.append(fh)

    logging.basicConfig(level=level, handlers=handlers, force=True)


def _build_client(cfg: dict, dry_run: bool):
    from m365seed.graph import GraphClient

    return GraphClient(cfg, dry_run=dry_run)


def _print_actions(actions: list[dict]) -> None:
    if not actions:
        console.print("[dim]No actions recorded.[/dim]")
        return
    table = Table(title="Actions")
    table.add_column("Action")
    table.add_column("Details")
    for a in actions:
        action = a.pop("action", "?")
        table.add_row(action, json.dumps(a, default=str))
    console.print(table)


# ═══════════════════════════════════════════════════════════════
# Commands
# ═══════════════════════════════════════════════════════════════


@app.command()
def validate(
    config: str = CONFIG_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    dry_run: bool = DRY_RUN_OPT,
) -> None:
    """Validate config, auth, permissions, and target users."""
    _setup_logging(verbose, log_file)
    logger = logging.getLogger("m365seed.cli")

    # 1. Config schema validation
    try:
        cfg = load_config(config)
        console.print("[green]✓[/green] Config schema is valid.")
    except Exception as exc:
        console.print(f"[red]✗[/red] Config validation failed: {exc}")
        raise typer.Exit(1)

    # 2. Auth check
    client = _build_client(cfg, dry_run)
    try:
        result = client.check_auth()
        console.print("[green]✓[/green] Graph authentication succeeded.")
        if isinstance(result, dict) and "value" in result:
            org = result["value"][0] if result["value"] else {}
            console.print(f"  Tenant: {org.get('displayName', 'N/A')}")
    except Exception as exc:
        console.print(f"[red]✗[/red] Auth check failed: {exc}")
        raise typer.Exit(1)

    # 3. User existence
    users = get_users(cfg)
    for u in users:
        upn = u["upn"]
        exists = client.check_user_exists(upn)
        mark = "[green]✓[/green]" if exists else "[red]✗[/red]"
        console.print(f"  {mark} User {upn}")
        if not exists:
            logger.warning("User %s not found in tenant.", upn)

    console.print("\n[bold]Validation complete.[/bold]")


@app.command("seed-mail")
def seed_mail_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Send synthetic theme-specific email threads."""
    _setup_logging(verbose, log_file)

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.mail import seed_mail

    actions = seed_mail(client, cfg, resolved_theme, run_id)
    _print_actions(actions)


@app.command("seed-files")
def seed_files_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Upload synthetic theme-specific files to OneDrive/SharePoint."""
    _setup_logging(verbose, log_file)

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.files import seed_files

    actions = seed_files(client, cfg, resolved_theme, run_id)
    _print_actions(actions)


@app.command("seed-calendar")
def seed_calendar_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Create synthetic theme-specific calendar events."""
    _setup_logging(verbose, log_file)

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.calendar import seed_calendar

    actions = seed_calendar(client, cfg, resolved_theme, run_id)
    _print_actions(actions)


@app.command("seed-teams")
def seed_teams_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
    enable_beta_teams: bool = typer.Option(
        False,
        "--enable-beta-teams",
        help="Enable Teams seeding (requires Graph /beta — unstable).",
    ),
) -> None:
    """Seed Teams channels and posts (beta — off by default)."""
    _setup_logging(verbose, log_file)

    if not enable_beta_teams:
        console.print(
            "[yellow]Teams seeding is disabled.[/yellow] "
            "Pass --enable-beta-teams to enable (uses Graph /beta)."
        )
        raise typer.Exit(0)

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.teams import seed_teams

    actions = seed_teams(client, cfg, resolved_theme, run_id)
    _print_actions(actions)


@app.command("seed-chats")
def seed_chats_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
    enable_beta_teams: bool = typer.Option(
        False,
        "--enable-beta-teams",
        help="Enable Teams chat seeding (requires Graph /beta).",
    ),
) -> None:
    """Seed Teams 1:1 and group chats with messages (beta — off by default)."""
    _setup_logging(verbose, log_file)

    if not enable_beta_teams:
        console.print(
            "[yellow]Chat seeding is disabled.[/yellow] "
            "Pass --enable-beta-teams to enable (uses Graph /beta)."
        )
        raise typer.Exit(0)

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.chats import seed_chats

    actions = seed_chats(client, cfg, resolved_theme, run_id)
    _print_actions(actions)


@app.command("seed-sharepoint")
def seed_sharepoint_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Create SharePoint sites, pages, and upload documents."""
    _setup_logging(verbose, log_file)

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.sharepoint import seed_sharepoint

    actions = seed_sharepoint(client, cfg, resolved_theme, run_id)
    _print_actions(actions)


@app.command("seed-planner")
def seed_planner_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Create Planner plans, buckets, and tasks."""
    _setup_logging(verbose, log_file)

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.planner import seed_planner

    actions = seed_planner(client, cfg, resolved_theme, run_id)
    _print_actions(actions)


@app.command("seed-all")
def seed_all_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
    enable_beta_teams: bool = typer.Option(
        False,
        "--enable-beta-teams",
        help="Enable Teams seeding (Graph /beta).",
    ),
) -> None:
    """Run all seeding commands in sequence."""
    _setup_logging(verbose, log_file)

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    all_actions: list[dict] = []

    from m365seed.mail import seed_mail
    from m365seed.files import seed_files
    from m365seed.calendar import seed_calendar

    console.print("\n[bold]── Mail ──[/bold]")
    all_actions.extend(seed_mail(client, cfg, resolved_theme, run_id))

    console.print("\n[bold]── Files ──[/bold]")
    all_actions.extend(seed_files(client, cfg, resolved_theme, run_id))

    console.print("\n[bold]── Calendar ──[/bold]")
    all_actions.extend(seed_calendar(client, cfg, resolved_theme, run_id))

    if enable_beta_teams:
        from m365seed.teams import seed_teams
        from m365seed.chats import seed_chats

        console.print("\n[bold]── Teams Channels (beta) ──[/bold]")
        all_actions.extend(seed_teams(client, cfg, resolved_theme, run_id))

        console.print("\n[bold]── Teams Chats (beta) ──[/bold]")
        all_actions.extend(seed_chats(client, cfg, resolved_theme, run_id))

    from m365seed.sharepoint import seed_sharepoint
    from m365seed.planner import seed_planner

    console.print("\n[bold]── SharePoint ──[/bold]")
    all_actions.extend(seed_sharepoint(client, cfg, resolved_theme, run_id))

    console.print("\n[bold]── Planner ──[/bold]")
    all_actions.extend(seed_planner(client, cfg, resolved_theme, run_id))

@app.command()
def setup(
    config: str = CONFIG_OPT,
) -> None:
    """Interactive setup wizard — configure tenant, theme, and content."""
    from m365seed.setup import run_setup

    run_setup(config_path=config)


@app.command()
def register(
    tenant_id: Optional[str] = typer.Option(
        None, "--tenant", "-t", help="Tenant ID (GUID). Prompted if omitted.",
    ),
) -> None:
    """Create an Entra ID app registration via Azure CLI (device-code login)."""
    from m365seed.register import run_registration_wizard

    run_registration_wizard(tenant_id=tenant_id)


@app.command()
def cleanup(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    mail: bool = typer.Option(True, help="Clean up seeded mail."),
    files: bool = typer.Option(True, help="Clean up seeded files."),
    calendar: bool = typer.Option(True, help="Clean up seeded calendar events."),
    teams: bool = typer.Option(True, help="Clean up seeded Teams channels."),
    chats: bool = typer.Option(True, help="Clean up seeded Teams chats."),
    sharepoint: bool = typer.Option(True, help="Clean up seeded SharePoint sites."),
    planner: bool = typer.Option(True, help="Clean up seeded Planner plans."),
) -> None:
    """Remove all seeded content tagged with the configured run_id."""
    _setup_logging(verbose, log_file)

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.cleanup import cleanup as do_cleanup

    console.print(
        f"[bold yellow]Cleaning up all seeded content for run_id={run_id}[/bold yellow]"
    )

    actions = do_cleanup(
        client,
        cfg,
        run_id,
        mail=mail,
        files=files,
        calendar=calendar,
        teams=teams,
        chats=chats,
        sharepoint=sharepoint,
        planner=planner,
    )
    _print_actions(actions)


# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app()
