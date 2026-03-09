# M365 Demo Tenant Seeding Tool — Task Plan

## Phase 1: Scaffold & Config
- [x] Create directory structure
- [x] Create `pyproject.toml` with dependencies
- [x] Create `.gitignore`
- [x] Create `seed-config.example.yaml`

## Phase 2: Core Modules
- [x] `m365seed/__init__.py` — package init
- [x] `m365seed/config.py` — YAML schema validation
- [x] `m365seed/graph.py` — Graph client, retry, auth
- [x] `m365seed/mail.py` — email seeding
- [x] `m365seed/files.py` — file seeding
- [x] `m365seed/calendar.py` — calendar seeding
- [x] `m365seed/teams.py` — teams seeding (beta-gated)
- [x] `m365seed/cleanup.py` — cleanup mode
- [x] `m365seed/cli.py` — CLI entry point

## Phase 3: Templates & Data
- [x] Email templates (Jinja2) — healthcare, pharma, medtech, payor
- [x] Document templates (SOP, compliance, discharge, staffing, etc.)
- [x] Synthetic data payloads per vertical (`data/themes.json`)

## Phase 4: Tests
- [x] `tests/test_config.py` — 17 tests
- [x] `tests/test_graph.py` — 10 tests
- [x] `tests/test_mail.py` — 10 tests
- [x] `tests/test_files.py` — 8 tests
- [x] `tests/test_calendar.py` — 6 tests

## Phase 5: Documentation
- [x] `README.md`
- [x] `AGENTS.md`
- [x] `spec.md`
- [x] `SKILL.md`

## Phase 6: Dev Container & Verification
- [x] Add `.devcontainer/devcontainer.json` (Python 3.12 + Azure CLI)
- [x] Run all tests — **58 passed, 0 warnings**
- [x] Fix deprecation warning (`datetime.utcnow` → `datetime.now(timezone.utc)`)
- [x] Fix `pyproject.toml` build-backend

## Phase 7: Theme-Aware Content System
- [x] Rewrite `themes.json` (83 KB — 4 verticals, all content types)
- [x] Create `theme_content.py` — typed, LRU-cached accessors
- [x] Create 21 document templates (7 per non-healthcare theme)
- [x] Enhance email templates to 6 variants each
- [x] Update all 7 modules: files, calendar, mail, teams, chats, sharepoint, planner
- [x] Update CLI to pass theme to all seed functions
- [x] Update all tests — **89 tests passing**
- [x] Update seed-config.example.yaml with theme enrichment comments

## Phase 8: Dev Container + Deployment Automation
- [x] Create custom `Dockerfile` (Python 3.12 + Azure CLI + all deps pre-baked)
- [x] Create `post-create.sh` auto-setup script
- [x] Update `devcontainer.json` (build from Dockerfile, GitHub CLI, env vars, extensions)
- [x] Create `m365seed/setup.py` — interactive setup wizard
- [x] Integrate `setup` command into CLI
- [x] Update `AGENTS.md` — full architecture, theme system, dev container, deployment workflow
- [x] Update `README.md` — dev container quick start, setup wizard
- [x] Update `SKILL.md` — theme system, dev container steps
- [x] Update `spec.md` — architecture, automation

## Phase 9: Autonomous Repair Loop (2026-02-27)
- [x] Identify non-interactive auth blocker (`M365SEED_CLIENT_SECRET` missing in shell)
- [x] Validate + run `seed-all --enable-beta-teams` + cleanup + reseed loop
- [x] Repair config drift: update `teams.team_id` and `planner.group_id` after team deletion
- [x] Patch delegated token handling for Azure CLI single-scope behavior (`m365seed/graph.py`)
- [x] Patch Teams channel existence to handle pagination (`m365seed/teams.py`)
- [x] Patch chat seeding to app-only first, delegated fallback on auth boundary (`m365seed/chats.py`)
- [x] Harden delegated Graph auth to purge Azure CLI HTTP cache before `AzureCliCredential` token requests (`m365seed/graph.py`)
- [x] Relax `validate` so app-only `/organization` `403` is reported as a permission warning instead of a hard auth failure (`m365seed/cli.py`)
- [x] Reroute live OneDrive file seeding to the provisioned admin personal site after `admin@M365CPI56568282.onmicrosoft.com` completed first-run SharePoint sign-in
- [ ] Remaining tenant-side blocker: Teams/Chats posting authorization still not fully clean in app-only runs
- [ ] Re-run live tenant validation after re-exporting `M365SEED_CLIENT_SECRET`
- [ ] Provision AllanD OneDrive later if persona-specific file placement is still required
