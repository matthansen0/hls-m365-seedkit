"""Tests for m365seed.setup helpers."""

from __future__ import annotations

import subprocess
from unittest.mock import patch

import yaml

from m365seed.setup import (
    _build_setup_graph_client,
    _guess_tenant_domain,
    _run_az_command,
    _create_missing_demo_users,
    _reset_demo_user_passwords,
    _discover_tenant_users,
    _match_default_users,
    _check_user_exists,
    _discover_teams,
    _create_team_group,
    _generate_config,
    THEMED_TEAM_NAMES,
)


def _cp(returncode: int, stdout: str = "", stderr: str = "") -> subprocess.CompletedProcess:
    return subprocess.CompletedProcess(args=["az"], returncode=returncode, stdout=stdout, stderr=stderr)


def test_guess_tenant_domain_from_matching_az_context() -> None:
    tenant_id = "2c627739-3b65-451a-ac0d-d3ecea353a55"
    payload = (
        '{"tenantId":"2c627739-3b65-451a-ac0d-d3ecea353a55",'
        '"tenantDefaultDomain":"M365x06303451.onmicrosoft.com"}'
    )

    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", return_value=_cp(0, stdout=payload)
    ):
        assert _guess_tenant_domain(tenant_id) == "m365x06303451.onmicrosoft.com"


def test_guess_tenant_domain_returns_none_on_tenant_mismatch() -> None:
    tenant_id = "2c627739-3b65-451a-ac0d-d3ecea353a55"
    payload = (
        '{"tenantId":"19a1b89f-a9ba-4694-ae2b-bd1dc0ee369b",'
        '"tenantDefaultDomain":"contoso.onmicrosoft.com"}'
    )

    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", return_value=_cp(0, stdout=payload)
    ):
        assert _guess_tenant_domain(tenant_id) is None


def test_guess_tenant_domain_returns_none_when_az_missing() -> None:
    with patch("m365seed.setup.shutil.which", return_value=None):
        assert _guess_tenant_domain("2c627739-3b65-451a-ac0d-d3ecea353a55") is None


def test_build_setup_graph_client_prefers_delegated() -> None:
    delegated = object()
    app_client = object()

    with patch(
        "m365seed.setup._build_setup_delegated_graph_client",
        return_value=delegated,
    ), patch(
        "m365seed.setup._build_setup_app_graph_client",
        return_value=app_client,
    ):
        client = _build_setup_graph_client(
            "2c627739-3b65-451a-ac0d-d3ecea353a55",
            "client-id",
            "M365SEED_CLIENT_SECRET",
        )

    assert client is delegated


def test_build_setup_graph_client_falls_back_to_app_client() -> None:
    app_client = object()

    with patch(
        "m365seed.setup._build_setup_delegated_graph_client",
        return_value=None,
    ), patch(
        "m365seed.setup._build_setup_app_graph_client",
        return_value=app_client,
    ):
        client = _build_setup_graph_client(
            "2c627739-3b65-451a-ac0d-d3ecea353a55",
            "client-id",
            "M365SEED_CLIENT_SECRET",
        )

    assert client is app_client


def test_reset_demo_user_passwords_success() -> None:
    users = [
        {"upn": "a@contoso.com", "role": "User A"},
        {"upn": "b@contoso.com", "role": "User B"},
    ]

    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", side_effect=[_cp(0), _cp(0)]
    ):
        success, failed = _reset_demo_user_passwords(
            users,
            "TempPass!234",
            force_change_next_sign_in=False,
        )

    assert success == 2
    assert failed == 0


def test_reset_demo_user_passwords_partial_failure() -> None:
    users = [
        {"upn": "a@contoso.com", "role": "User A"},
        {"upn": "b@contoso.com", "role": "User B"},
    ]

    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", side_effect=[_cp(0), _cp(1, stderr="boom")]
    ):
        success, failed = _reset_demo_user_passwords(
            users,
            "TempPass!234",
            force_change_next_sign_in=True,
        )

    assert success == 1
    assert failed == 1


def test_reset_demo_user_passwords_without_az_cli() -> None:
    users = [{"upn": "a@contoso.com", "role": "User A"}]

    with patch("m365seed.setup.shutil.which", return_value=None):
        success, failed = _reset_demo_user_passwords(
            users,
            "TempPass!234",
            force_change_next_sign_in=False,
        )

    assert success == 0
    assert failed == 1


def test_create_missing_demo_users_success() -> None:
    users = [{"upn": "a@contoso.com", "role": "User A"}]

    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", return_value=_cp(0, stdout="{}")
    ):
        success, failed, created = _create_missing_demo_users(
            users,
            "TempPass!234",
            force_change_next_sign_in=True,
        )

    assert success == 1
    assert failed == 0
    assert created == users


def test_run_az_command_purges_http_cache() -> None:
    with patch("m365seed.register._ensure_msal_cache_healthy") as ensure_healthy, patch(
        "m365seed.setup.subprocess.run", return_value=_cp(0, stdout="{}")
    ):
        result = _run_az_command(["account", "show", "--output", "json"])

    assert result.returncode == 0
    ensure_healthy.assert_called_once()


# ---------------------------------------------------------------------------
# User discovery / matching tests
# ---------------------------------------------------------------------------


def test_discover_tenant_users_returns_users() -> None:
    payload = '[{"upn":"AllanD@contoso.onmicrosoft.com","displayName":"Allan Deyoung","mail":null}]'
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", return_value=_cp(0, stdout=payload)
    ):
        users = _discover_tenant_users()
    assert len(users) == 1
    assert users[0]["upn"] == "AllanD@contoso.onmicrosoft.com"


def test_discover_tenant_users_filters_ext_guests() -> None:
    payload = (
        '[{"upn":"AllanD@contoso.onmicrosoft.com","displayName":"Allan","mail":null},'
        '{"upn":"guest_partner.com#EXT#@contoso.onmicrosoft.com","displayName":"Guest","mail":null}]'
    )
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", return_value=_cp(0, stdout=payload)
    ):
        users = _discover_tenant_users()
    assert len(users) == 1
    assert "EXT" not in users[0]["upn"]


def test_discover_tenant_users_empty_when_az_missing() -> None:
    with patch("m365seed.setup.shutil.which", return_value=None):
        users = _discover_tenant_users()
    assert users == []


def test_discover_tenant_users_empty_on_az_failure() -> None:
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", return_value=_cp(1, stderr="not logged in")
    ):
        users = _discover_tenant_users()
    assert users == []


def test_match_default_users_all_found() -> None:
    tenant_users = [
        {"upn": "AllanD@contoso.onmicrosoft.com", "displayName": "Allan"},
        {"upn": "MeganB@contoso.onmicrosoft.com", "displayName": "Megan"},
        {"upn": "NestorW@contoso.onmicrosoft.com", "displayName": "Nestor"},
        {"upn": "LeeG@contoso.onmicrosoft.com", "displayName": "Lee"},
        {"upn": "JoniS@contoso.onmicrosoft.com", "displayName": "Joni"},
    ]
    result = _match_default_users("contoso.onmicrosoft.com", tenant_users)
    assert len(result) == 5
    assert all(u["matched"] == "true" for u in result)
    assert result[0]["upn"] == "AllanD@contoso.onmicrosoft.com"


def test_match_default_users_partial_match() -> None:
    tenant_users = [
        {"upn": "AllanD@contoso.onmicrosoft.com", "displayName": "Allan"},
        {"upn": "MeganB@contoso.onmicrosoft.com", "displayName": "Megan"},
    ]
    result = _match_default_users("contoso.onmicrosoft.com", tenant_users)
    found = [u for u in result if u["matched"] == "true"]
    missing = [u for u in result if u["matched"] == "false"]
    assert len(found) == 2
    assert len(missing) == 3


def test_match_default_users_none_found() -> None:
    tenant_users = [
        {"upn": "CustomUser1@contoso.onmicrosoft.com", "displayName": "Custom"},
    ]
    result = _match_default_users("contoso.onmicrosoft.com", tenant_users)
    assert all(u["matched"] == "false" for u in result)


def test_match_default_users_case_insensitive() -> None:
    tenant_users = [
        {"upn": "alland@contoso.onmicrosoft.com", "displayName": "Allan"},
    ]
    result = _match_default_users("contoso.onmicrosoft.com", tenant_users)
    allan = [u for u in result if "allan" in u["upn"].lower()]
    assert len(allan) == 1
    assert allan[0]["matched"] == "true"
    assert allan[0]["upn"] == "alland@contoso.onmicrosoft.com"


def test_check_user_exists_true() -> None:
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run",
        return_value=_cp(0, stdout="12345678-abcd-ef01-2345-000000000000\n"),
    ):
        assert _check_user_exists("AllanD@contoso.onmicrosoft.com") is True


def test_check_user_exists_false() -> None:
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run",
        return_value=_cp(1, stderr="Resource does not exist"),
    ):
        assert _check_user_exists("gone@contoso.onmicrosoft.com") is False


def test_check_user_exists_optimistic_when_az_missing() -> None:
    with patch("m365seed.setup.shutil.which", return_value=None):
        assert _check_user_exists("anyone@contoso.onmicrosoft.com") is True


# ---------------------------------------------------------------------------
# Team discovery tests
# ---------------------------------------------------------------------------


def test_discover_teams_returns_groups() -> None:
    payload = (
        '[{"id":"11111111-1111-1111-1111-111111111111",'
        '"displayName":"Clinical Ops Team"},'
        '{"id":"22222222-2222-2222-2222-222222222222",'
        '"displayName":"Research Team"}]'
    )
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", return_value=_cp(0, stdout=payload)
    ):
        teams = _discover_teams("some-tenant-id")
    assert len(teams) == 2
    assert teams[0]["displayName"] == "Clinical Ops Team"
    assert teams[1]["id"] == "22222222-2222-2222-2222-222222222222"


def test_discover_teams_empty_when_az_missing() -> None:
    with patch("m365seed.setup.shutil.which", return_value=None):
        teams = _discover_teams("some-tenant-id")
    assert teams == []


def test_discover_teams_empty_on_az_failure() -> None:
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", return_value=_cp(1, stderr="not logged in")
    ):
        teams = _discover_teams("some-tenant-id")
    assert teams == []


def test_discover_teams_filters_entries_without_id() -> None:
    payload = (
        '[{"id":"11111111-1111-1111-1111-111111111111",'
        '"displayName":"Good Team"},'
        '{"id":"","displayName":"Bad Team"}]'
    )
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run", return_value=_cp(0, stdout=payload)
    ):
        teams = _discover_teams("some-tenant-id")
    assert len(teams) == 1
    assert teams[0]["displayName"] == "Good Team"


# ---------------------------------------------------------------------------
# Config generation tests
# ---------------------------------------------------------------------------

_USERS = [
    {"upn": "AllanD@contoso.onmicrosoft.com", "role": "Clinical Ops Manager"},
    {"upn": "MeganB@contoso.onmicrosoft.com", "role": "Physician Lead"},
    {"upn": "NestorW@contoso.onmicrosoft.com", "role": "Care Coordinator"},
]

_BASE_SECTIONS = {
    "mail": True,
    "files": True,
    "calendar": True,
    "teams": False,
    "chats": False,
    "sharepoint": False,
    "planner": False,
}


def _gen(theme: str = "healthcare", sections: dict | None = None, **kwargs):
    """Helper to call _generate_config with sensible defaults."""
    s = dict(_BASE_SECTIONS)
    if sections:
        s.update(sections)
    raw = _generate_config(
        tenant_id="aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
        client_id="11111111-2222-3333-4444-555555555555",
        secret_env="M365SEED_CLIENT_SECRET",
        theme=theme,
        run_id="test-run-001",
        users=kwargs.pop("users", _USERS),
        sections=s,
        team_id=kwargs.pop("team_id", ""),
        group_id=kwargs.pop("group_id", ""),
    )
    return yaml.safe_load(raw)


def test_generate_config_has_required_top_level_keys() -> None:
    cfg = _gen()
    for key in ("tenant", "auth", "targets", "content", "mail", "files", "calendar",
                "teams", "chats", "sharepoint", "planner"):
        assert key in cfg, f"Missing top-level key: {key}"


def test_generate_config_uses_correct_tenant_and_auth() -> None:
    cfg = _gen()
    assert cfg["tenant"]["tenant_id"] == "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"
    assert cfg["auth"]["client_id"] == "11111111-2222-3333-4444-555555555555"
    assert cfg["auth"]["client_secret_env"] == "M365SEED_CLIENT_SECRET"


def test_generate_config_includes_theme_mail_threads() -> None:
    cfg = _gen(theme="healthcare")
    threads = cfg["mail"]["threads"]
    assert len(threads) == 3
    assert threads[0]["thread_id"] == "care-coordination-001"


def test_generate_config_uses_theme_specific_mail_thread_ids() -> None:
    """Pharma config must NOT use healthcare thread IDs."""
    cfg = _gen(theme="pharma")
    thread_ids = [t["thread_id"] for t in cfg["mail"]["threads"]]
    assert all("care-coordination" not in tid for tid in thread_ids)
    # Pharma should have its own IDs
    assert any("pipeline" in tid or "trial" in tid or "regulatory" in tid for tid in thread_ids)


def test_generate_config_calendar_events_from_theme() -> None:
    cfg = _gen(theme="healthcare")
    events = cfg["calendar"]["events"]
    assert len(events) == 5
    assert events[0]["event_id"] == "clinical-ops-standup-001"
    # Each event should have organizer and attendees
    assert "organizer" in events[0]
    assert "attendees" in events[0]


def test_generate_config_teams_channels_populated_when_enabled() -> None:
    cfg = _gen(
        theme="healthcare",
        sections={"teams": True},
        team_id="aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa",
    )
    assert cfg["teams"]["enabled"] is True
    assert cfg["teams"]["team_id"] == "aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa"
    channels = cfg["teams"]["channels"]
    assert len(channels) == 4  # healthcare has 4 channels
    assert channels[0]["channel_id"] == "care-updates-001"
    assert channels[0]["display_name"] == "Care Updates"
    # Posts should be populated as dicts with 'message' key
    assert "posts" in channels[0]
    assert len(channels[0]["posts"]) > 0
    assert "message" in channels[0]["posts"][0]


def test_generate_config_teams_empty_when_disabled() -> None:
    cfg = _gen(sections={"teams": False})
    assert cfg["teams"]["enabled"] is False
    assert cfg["teams"]["channels"] == []


def test_generate_config_chats_populated_when_enabled() -> None:
    cfg = _gen(
        theme="healthcare",
        sections={"chats": True},
    )
    assert cfg["chats"]["enabled"] is True
    convs = cfg["chats"]["conversations"]
    assert len(convs) == 4  # healthcare has 4 chat conversations
    assert convs[0]["conversation_id"] == "care-handoff-chat-001"
    assert convs[0]["type"] == "group"
    assert "topic" in convs[0]
    # Members should be populated from user pool
    assert "members" in convs[0]
    assert len(convs[0]["members"]) > 0
    # Messages should be dicts with 'sender' and 'text' keys
    assert "messages" in convs[0]
    assert len(convs[0]["messages"]) > 0
    assert "text" in convs[0]["messages"][0]
    assert "sender" in convs[0]["messages"][0]
    # Sender must be a UPN from the user pool
    assert "@" in convs[0]["messages"][0]["sender"]


def test_generate_config_chats_empty_when_disabled() -> None:
    cfg = _gen(sections={"chats": False})
    assert cfg["chats"]["enabled"] is False
    assert cfg["chats"]["conversations"] == []


def test_generate_config_sharepoint_populated_when_enabled() -> None:
    cfg = _gen(
        theme="healthcare",
        sections={"sharepoint": True},
    )
    assert cfg["sharepoint"]["enabled"] is True
    sites = cfg["sharepoint"]["sites"]
    assert len(sites) == 3  # healthcare has 3 SP sites
    assert sites[0]["display_name"] == "Clinical Operations Hub"
    assert "description" in sites[0]
    # Owner should be set from first user
    assert cfg["sharepoint"]["owner"] == "AllanD@contoso.onmicrosoft.com"


def test_generate_config_sharepoint_empty_when_disabled() -> None:
    cfg = _gen(sections={"sharepoint": False})
    assert cfg["sharepoint"]["enabled"] is False
    assert cfg["sharepoint"]["sites"] == []


def test_generate_config_planner_populated_when_enabled() -> None:
    cfg = _gen(
        theme="healthcare",
        sections={"planner": True},
        group_id="bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb",
    )
    assert cfg["planner"]["enabled"] is True
    assert cfg["planner"]["group_id"] == "bbbbbbbb-bbbb-bbbb-bbbb-bbbbbbbbbbbb"
    plans = cfg["planner"]["plans"]
    assert len(plans) == 3  # healthcare has 3 plans
    assert plans[0]["title"] == "Clinical Operations Sprint"


def test_generate_config_planner_empty_when_disabled() -> None:
    cfg = _gen(sections={"planner": False})
    assert cfg["planner"]["enabled"] is False
    assert cfg["planner"]["plans"] == []


def test_generate_config_medtech_theme() -> None:
    """Verify config generation works for a non-default theme."""
    cfg = _gen(
        theme="medtech",
        sections={"teams": True, "chats": True, "sharepoint": True, "planner": True},
        team_id="cccccccc-cccc-cccc-cccc-cccccccccccc",
        group_id="cccccccc-cccc-cccc-cccc-cccccccccccc",
    )
    assert cfg["content"]["theme"] == "medtech"
    # Mail should have medtech-specific threads
    thread_ids = [t["thread_id"] for t in cfg["mail"]["threads"]]
    assert all("care-coordination" not in tid for tid in thread_ids)
    # Teams channels should be medtech-specific
    assert len(cfg["teams"]["channels"]) == 4
    # Chats should be populated
    assert len(cfg["chats"]["conversations"]) == 4
    # SharePoint sites should be populated
    assert len(cfg["sharepoint"]["sites"]) == 3
    # Planner plans should be populated
    assert len(cfg["planner"]["plans"]) == 3


def test_generate_config_all_four_themes_produce_valid_yaml() -> None:
    """Every theme should produce parseable YAML with the right structure."""
    for theme in ("healthcare", "pharma", "medtech", "payor"):
        cfg = _gen(
            theme=theme,
            sections={"teams": True, "chats": True, "sharepoint": True, "planner": True},
            team_id="eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee",
            group_id="eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee",
        )
        assert cfg["content"]["theme"] == theme
        assert len(cfg["mail"]["threads"]) >= 2, f"{theme} has too few mail threads"
        assert len(cfg["calendar"]["events"]) >= 2, f"{theme} has too few events"
        assert len(cfg["teams"]["channels"]) >= 2, f"{theme} has too few channels"
        assert len(cfg["chats"]["conversations"]) >= 2, f"{theme} has too few chats"
        assert len(cfg["sharepoint"]["sites"]) >= 2, f"{theme} has too few SP sites"
        assert len(cfg["planner"]["plans"]) >= 2, f"{theme} has too few plans"


# ---------------------------------------------------------------------------
# Team/group creation tests
# ---------------------------------------------------------------------------


def test_themed_team_names_all_four_themes() -> None:
    for theme in ("healthcare", "pharma", "medtech", "payor"):
        assert theme in THEMED_TEAM_NAMES
        assert len(THEMED_TEAM_NAMES[theme]) > 0


def test_create_team_group_success() -> None:
    group_json = '{"id":"aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee","displayName":"Contoso Health System"}'
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run",
        side_effect=[
            _cp(0, stdout=group_json),   # az rest POST (create group)
            _cp(0, stdout="{}"),          # az rest PUT  (team-enable)
        ],
    ), patch("m365seed.setup.time.sleep"):   # skip delay
        gid = _create_team_group("Contoso Health System")
    assert gid == "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"


def test_create_team_group_failure_returns_none() -> None:
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run",
        return_value=_cp(1, stderr="Insufficient permissions"),
    ):
        gid = _create_team_group("Bad Group")
    assert gid is None


def test_create_team_group_no_az_returns_none() -> None:
    with patch("m365seed.setup.shutil.which", return_value=None):
        gid = _create_team_group("No CLI")
    assert gid is None


def test_create_team_group_team_enable_retries() -> None:
    """Group created but team-enable needs retries before succeeding."""
    group_json = '{"id":"aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee","displayName":"Test"}'
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run",
        side_effect=[
            _cp(0, stdout=group_json),   # create group
            _cp(1, stderr="Not ready"),   # team-enable attempt 1 fails
            _cp(1, stderr="Not ready"),   # team-enable attempt 2 fails
            _cp(0, stdout="{}"),          # team-enable attempt 3 succeeds
        ],
    ), patch("m365seed.setup.time.sleep"):
        gid = _create_team_group("Test")
    assert gid == "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"


def test_create_team_group_returns_id_even_if_team_enable_fails() -> None:
    """Group created, all team-enable attempts fail — still returns group ID."""
    group_json = '{"id":"aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee","displayName":"Test"}'
    with patch("m365seed.setup.shutil.which", return_value="/usr/bin/az"), patch(
        "m365seed.setup.subprocess.run",
        side_effect=[
            _cp(0, stdout=group_json),
            _cp(1, stderr="fail"), _cp(1, stderr="fail"),
            _cp(1, stderr="fail"), _cp(1, stderr="fail"),
            _cp(1, stderr="fail"),
        ],
    ), patch("m365seed.setup.time.sleep"):
        gid = _create_team_group("Test")
    # Group was created, so ID is still returned even if team-enable failed
    assert gid == "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee"


# ---------------------------------------------------------------------------
# Schema validation integration test
# ---------------------------------------------------------------------------


def test_generate_config_passes_schema_validation_all_themes() -> None:
    """Generated config must pass jsonschema validation for every theme."""
    import jsonschema
    from m365seed.config import CONFIG_SCHEMA

    for theme in ("healthcare", "pharma", "medtech", "payor"):
        cfg = _gen(
            theme=theme,
            sections={
                "mail": True, "files": True, "calendar": True,
                "teams": True, "chats": True,
                "sharepoint": True, "planner": True,
            },
            team_id="eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee",
            group_id="eeeeeeee-eeee-eeee-eeee-eeeeeeeeeeee",
        )
        try:
            jsonschema.validate(cfg, CONFIG_SCHEMA)
        except jsonschema.ValidationError as e:
            raise AssertionError(
                f"Schema validation failed for theme '{theme}': "
                f"{e.message} at path {list(e.absolute_path)}"
            ) from e
