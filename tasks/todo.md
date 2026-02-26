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
