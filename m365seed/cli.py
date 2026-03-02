"""CLI entry point — ``m365seed`` commands."""

from __future__ import annotations

import json
import logging
import sys
from datetime import datetime
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
    help="Path for JSONL structured log output. Auto-generated in logs/ if omitted.",
)

VERBOSE_OPT = typer.Option(
    False,
    "--verbose",
    "-v",
    help="Enable debug-level logging.",
)

# Tracks the active log file path so we can print it at the end of a command.
_active_log_file: str | None = None


def _auto_log_path(command: str = "seed") -> str:
    """Return an auto-generated log file path under ``logs/``."""
    logs_dir = Path("logs")
    logs_dir.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%dT%H%M%S")
    return str(logs_dir / f"{command}_{ts}.jsonl")


def _setup_logging(
    verbose: bool,
    log_file: Optional[str] = None,
    *,
    command: str = "seed",
) -> None:
    global _active_log_file
    level = logging.DEBUG if verbose else logging.INFO
    handlers: list[logging.Handler] = [
        RichHandler(console=console, show_path=False, rich_tracebacks=True)
    ]

    # Auto-create a log file if the caller didn't supply one
    resolved = log_file or _auto_log_path(command)
    _active_log_file = resolved

    fh = logging.FileHandler(resolved, mode="a", encoding="utf-8")
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
            # Clear args so the logging framework won't attempt
            # %-style formatting on the already-serialised JSON string.
            record.args = None
            return True

    fh.addFilter(JSONLFilter())
    handlers.append(fh)

    logging.basicConfig(level=level, handlers=handlers, force=True)

    # Silence noisy Azure SDK loggers on the console — they still go to the file
    for noisy in ("azure", "azure.core", "azure.identity", "httpx", "httpcore"):
        logging.getLogger(noisy).setLevel(logging.WARNING)


def _print_log_path() -> None:
    """Print a summary of issues and the log file location."""
    _print_run_summary()
    if _active_log_file:
        console.print(f"\n[dim]Log written to:[/dim] [cyan]{_active_log_file}[/cyan]")


def _print_run_summary() -> None:
    """Parse the active log file and print a summary of warnings/errors."""
    if not _active_log_file:
        return
    try:
        warnings: list[str] = []
        errors: list[str] = []
        with open(_active_log_file, encoding="utf-8") as fh:
            for line in fh:
                try:
                    rec = json.loads(line.strip())
                    if isinstance(rec, str):
                        rec = json.loads(rec)
                except (json.JSONDecodeError, TypeError):
                    continue
                level = rec.get("level", "")
                if level == "WARNING":
                    warnings.append(rec.get("msg", ""))
                elif level in ("ERROR", "CRITICAL"):
                    errors.append(rec.get("msg", ""))

        total_issues = len(warnings) + len(errors)
        if total_issues == 0:
            console.print("\n[bold green]✓ Run completed with no warnings or errors.[/bold green]")
            return

        console.print(f"\n[bold yellow]⚠ Run completed with {total_issues} issue(s):[/bold yellow]")

        # Deduplicate and summarize — group by module + HTTP status
        import re as _re
        from collections import Counter as _Counter

        def _categorise(msg: str) -> str:
            """Return a short category string from a log message."""
            # Extract module from msg if present
            status_match = _re.search(r"(\d{3})\s+(Forbidden|Bad Request|Not Found|Unauthorized|Conflict)", msg)
            status = status_match.group(0) if status_match else ""
            # Extract the first meaningful phrase
            short = msg.split(":")[0].strip() if ":" in msg else msg[:80]
            return f"{short} ({status})" if status else short

        if errors:
            err_counts = _Counter(_categorise(m) for m in errors)
            console.print(f"  [red]Errors ({len(errors)}):[/red]")
            for cat, count in err_counts.most_common(10):
                console.print(f"    [red]✗[/red] {cat}  [dim]×{count}[/dim]" if count > 1 else f"    [red]✗[/red] {cat}")

        if warnings:
            warn_counts = _Counter(_categorise(m) for m in warnings)
            console.print(f"  [yellow]Warnings ({len(warnings)}):[/yellow]")
            for cat, count in warn_counts.most_common(10):
                console.print(f"    [yellow]⚠[/yellow] {cat}  [dim]×{count}[/dim]" if count > 1 else f"    [yellow]⚠[/yellow] {cat}")

    except Exception:
        pass  # Don't let summary logic break the CLI


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
        action = a.get("action", "?")
        details = {k: v for k, v in a.items() if k != "action"}
        table.add_row(action, json.dumps(details, default=str))
    console.print(table)


def _theme_label(theme: str) -> str:
    labels = {
        "healthcare": "Healthcare Provider",
        "pharma": "Pharma",
        "medtech": "MedTech",
        "payor": "Payor",
    }
    return labels.get(theme, theme.replace("_", " ").title())


def _print_seed_summary(theme: str, actions: list[dict]) -> None:
    """Print a compact seed rollup summary based on recorded actions."""
    action_counts: dict[str, int] = {}
    for item in actions:
        name = item.get("action", "")
        if not isinstance(name, str):
            continue
        action_counts[name] = action_counts.get(name, 0) + 1

    def count(*names: str) -> int:
        return sum(action_counts.get(n, 0) for n in names)

    summary = {
        "Theme": _theme_label(theme),
        "Users updated": count("update-profile"),
        "Groups/Sites created": count("create_site"),
        "Emails sent": count("send_mail"),
        "Teams channels created": count("create_channel"),
        "Meetings created": count("create_event"),
        "Files uploaded": count("upload", "upload_document"),
        "Chats created": count("create_chat"),
        "Chat messages sent": count("send_chat_message"),
        "Planner plans created": count("create_plan"),
    }

    table = Table(title="Seed Summary")
    table.add_column("Metric")
    table.add_column("Value", justify="right")
    for metric, value in summary.items():
        table.add_row(metric, str(value))
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
    _setup_logging(verbose, log_file, command="validate")
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
    _print_log_path()


@app.command("seed-profiles")
def seed_profiles_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Update user profiles (jobTitle, department, company) to match the theme."""
    _setup_logging(verbose, log_file, command="seed-profiles")

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.profiles import seed_profiles

    actions = seed_profiles(client, cfg, resolved_theme, run_id)
    _print_actions(actions)
    _print_log_path()


@app.command("seed-mail")
def seed_mail_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Send synthetic theme-specific email threads."""
    _setup_logging(verbose, log_file, command="seed-mail")

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.mail import seed_mail

    actions = seed_mail(client, cfg, resolved_theme, run_id)
    _print_actions(actions)
    _print_log_path()


@app.command("seed-files")
def seed_files_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Upload synthetic theme-specific files to OneDrive/SharePoint."""
    _setup_logging(verbose, log_file, command="seed-files")

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.files import seed_files

    actions = seed_files(client, cfg, resolved_theme, run_id)
    _print_actions(actions)
    _print_log_path()


@app.command("seed-calendar")
def seed_calendar_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Create synthetic theme-specific calendar events."""
    _setup_logging(verbose, log_file, command="seed-calendar")

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.calendar import seed_calendar

    actions = seed_calendar(client, cfg, resolved_theme, run_id)
    _print_actions(actions)
    _print_log_path()


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
    _setup_logging(verbose, log_file, command="seed-teams")

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
    _print_log_path()


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
    _setup_logging(verbose, log_file, command="seed-chats")

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
    _print_log_path()


@app.command("seed-sharepoint")
def seed_sharepoint_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Create SharePoint sites, pages, and upload documents."""
    _setup_logging(verbose, log_file, command="seed-sharepoint")

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.sharepoint import seed_sharepoint

    actions = seed_sharepoint(client, cfg, resolved_theme, run_id)
    _print_actions(actions)
    _print_log_path()


@app.command("seed-planner")
def seed_planner_cmd(
    config: str = CONFIG_OPT,
    dry_run: bool = DRY_RUN_OPT,
    verbose: bool = VERBOSE_OPT,
    log_file: Optional[str] = LOG_FILE_OPT,
    theme: Optional[str] = typer.Option(None, help="Override content theme."),
) -> None:
    """Create Planner plans, buckets, and tasks."""
    _setup_logging(verbose, log_file, command="seed-planner")

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    from m365seed.planner import seed_planner

    actions = seed_planner(client, cfg, resolved_theme, run_id)
    _print_actions(actions)
    _print_log_path()


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
    _setup_logging(verbose, log_file, command="seed-all")

    cfg = load_config(config)
    run_id = get_run_id(cfg)
    resolved_theme = theme or get_theme(cfg)
    client = _build_client(cfg, dry_run)

    all_actions: list[dict] = []

    from m365seed.profiles import seed_profiles
    from m365seed.mail import seed_mail
    from m365seed.files import seed_files
    from m365seed.calendar import seed_calendar

    console.print("\n[bold]── Profiles ──[/bold]")
    all_actions.extend(seed_profiles(client, cfg, resolved_theme, run_id))

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

    _print_seed_summary(resolved_theme, all_actions)
    _print_actions(all_actions)
    _print_log_path()

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
    team_group: bool = typer.Option(True, help="Clean up M365 Group created for Teams/Planner."),
) -> None:
    """Remove all seeded content tagged with the configured run_id."""
    _setup_logging(verbose, log_file, command="cleanup")

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
        team_group=team_group,
    )
    _print_actions(actions)
    _print_log_path()


# ═══════════════════════════════════════════════════════════════

if __name__ == "__main__":
    app()
