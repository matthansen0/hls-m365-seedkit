# AGENTS.md — Operating Layer

> **Purpose**: Defines collaboration norms and working agreements for this project.
> **Load Behavior**: Always loaded (every session).

---

## Project Identity

**M365 Demo Tenant Seeding Tool (hls-m365-seedkit)**

A safe, idempotent seeding tool for Microsoft 365 demo tenants with synthetic, theme-aware content across four HLS verticals (Healthcare, Pharma, MedTech, Payor). Supports Work IQ, Foundry IQ, and Fabric IQ demonstrations in Healthcare and Life Sciences.

---

## Architecture Overview

```
hls-m365-seedkit/
├── .devcontainer/          # Containerised dev environment
│   ├── Dockerfile          # Python 3.12 + Azure CLI + all deps
│   ├── devcontainer.json   # VS Code dev-container config
│   └── post-create.sh      # Auto-setup on container creation
├── m365seed/               # Core Python package
│   ├── cli.py              # Typer CLI — commands + setup wizard
│   ├── setup.py            # Interactive setup wizard (m365seed setup)
│   ├── config.py           # YAML schema validation
│   ├── graph.py            # Graph client, retry, auth
│   ├── theme_content.py    # Theme content provider (typed, LRU-cached)
│   ├── profiles.py         # User profile branding (jobTitle, dept, company)
│   ├── mail.py             # Email seeding (theme-aware)
│   ├── files.py            # File seeding — OneDrive + SharePoint
│   ├── calendar.py         # Calendar seeding (+ online meetings)
│   ├── teams.py            # Teams channels/posts (beta)
│   ├── chats.py            # Teams 1:1/group chats (beta)
│   ├── sharepoint.py       # SharePoint sites, pages, docs
│   ├── planner.py          # Planner plans, buckets, tasks
│   ├── cleanup.py          # Cleanup mode (7 content types)
│   ├── templates/          # Jinja2 templates per theme
│   │   ├── healthcare/     # 3 email + 7 doc templates
│   │   ├── pharma/         # 3 email + 7 doc templates
│   │   ├── medtech/        # 3 email + 7 doc templates
│   │   └── payor/          # 3 email + 7 doc templates
│   └── data/
│       └── themes.json     # 83 KB — all theme content (4 verticals)
├── tests/                  # 89 unit tests (mocked Graph calls)
├── seed-config.example.yaml
├── AGENTS.md               # This file
├── SKILL.md                # Execution playbook
├── spec.md                 # Feature specification
└── README.md               # User-facing documentation
```

---

## Theme System

Content is fully theme-aware across all 7 seeding modules. The `theme_content.py` module provides typed, LRU-cached accessors:

| Accessor | Returns |
|----------|---------|
| `get_user_profiles(theme)` | User profile attributes (jobTitle, department, company, office) |
| `get_file_manifest(theme)` | File manifest (folder, name, template, description) |
| `get_mail_threads(theme)` | Email thread metadata + attachment content |
| `get_calendar_events(theme)` | Calendar event body text |
| `get_teams_channels(theme)` | Channel descriptions + post content |
| `get_chat_conversations(theme)` | Chat conversation messages + topics |
| `get_sharepoint_sites(theme)` | Site pages + documents |
| `get_planner_plans(theme)` | Plan buckets + tasks |

Source data lives in `m365seed/data/themes.json`. Templates live in `m365seed/templates/{theme}/`.

---

## Dev Container

The `.devcontainer/` directory provides a fully configured development environment:

- **Python 3.12** with all runtime and dev dependencies pre-installed
- **Azure CLI** for interactive tenant management
- **GitHub CLI** for repo operations
- **VS Code extensions**: Python, Pylance, Ruff, Azure Resources, YAML
- **Auto-setup**: `post-create.sh` installs the project, copies example config, runs tests

### Environment Variables

Set these locally before opening the dev container:

| Variable | Purpose | Required |
|----------|---------|----------|
| `M365SEED_CLIENT_SECRET` | Entra ID app client secret | Yes |
| `M365SEED_TENANT_ID` | Target tenant GUID | Optional (setup wizard) |
| `M365SEED_CLIENT_ID` | App registration client ID | Optional (setup wizard) |

### Azure CLI Login in Dev Containers

The dev container has no browser, so you must use `--use-device-code`.
M365-only tenants have no Azure subscription, so add `--allow-no-subscriptions`:

```bash
az login --tenant <TENANT_ID> --allow-no-subscriptions --use-device-code
```

This identity is used by the **delegated auth** path (Teams channel messages, Teams chats). The `m365seed setup` wizard automatically adds this user as an owner/member of newly created Teams so that `seed-teams` and `seed-chats` don't get 403 errors.

---

## Deployment Workflow

### Quick Start (Dev Container)

1. Open in VS Code with the Dev Containers extension
2. The container auto-builds with all dependencies
3. Run `m365seed setup` for interactive configuration
4. Run `m365seed seed-all --dry-run` to verify
5. Run `m365seed seed-all` to seed the tenant

### Quick Start (Local)

```bash
python -m venv .venv && source .venv/bin/activate
pip install -e ".[dev]"
m365seed setup                    # Interactive wizard
m365seed validate                 # Verify auth + users
m365seed seed-all --dry-run       # Preview actions
m365seed seed-all                 # Seed the tenant
```

### CLI Commands

| Command | Description |
|---------|-------------|
| `m365seed setup` | Interactive setup wizard — generates seed-config.yaml |
| `m365seed register` | Automated Entra ID app registration via Azure CLI |
| `m365seed validate` | Validate config, auth, permissions, users |
| `m365seed seed-profiles` | Brand user profiles to match theme |
| `m365seed seed-mail` | Seed email threads |
| `m365seed seed-files` | Seed OneDrive/SharePoint files |
| `m365seed seed-calendar` | Seed calendar events |
| `m365seed seed-teams` | Seed Teams channels (beta) |
| `m365seed seed-chats` | Seed Teams chats (beta) |
| `m365seed seed-sharepoint` | Seed SharePoint sites |
| `m365seed seed-planner` | Seed Planner plans |
| `m365seed seed-all` | Run all seeders |
| `m365seed cleanup` | Remove seeded content by run_id |

---

## Working Agreements

### 1. Safety First
- **No real data**: All content is synthetic. No PHI, no PII.
- **Idempotent by design**: Every operation must be safe to rerun.
- **Dry-run default**: Always test with `--dry-run` before live execution.
- **No secrets in code**: Auth credentials must come from environment variables.

### 2. Code Standards
- **Python 3.11+** — use modern Python features (type hints, f-strings, `match` where appropriate).
- **Structured logging** — use the `logging` module with `m365seed.*` logger names.
- **Tests for every module** — unit tests with mocked Graph calls; no live API calls in tests.
- Follow **PEP 8** and use docstrings on all public functions.

### 3. Graph API Discipline
- **v1.0 endpoints only** unless explicitly gated behind `--enable-beta-teams`.
- **Handle 429 Retry-After** — the Graph client must implement exponential back-off.
- **Document every permission** — every Graph permission used must be listed in README.md.
- **No undocumented APIs** — only Microsoft Learn-documented endpoints.

### 4. Configuration Over Code
- Behavior is driven by `seed-config.yaml`, not hardcoded values.
- Templates are Jinja2 files, not inline strings (when practical).
- Themes are pluggable — adding a new theme means adding a template folder + themes.json entry.
- The `m365seed setup` wizard generates config interactively.

### 5. Collaboration Norms
- Update `tasks/todo.md` with progress as you work.
- After any correction or mistake, capture the pattern in `tasks/lessons.md`.
- Plan before building — enter plan mode for non-trivial tasks.
- Verify before declaring done — run tests, check outputs.

### 6. Synthetic Persona Changes

When requesting changes to patient/provider personas in the demo data, use this structured format:

```
PATIENT:       Name, age, location, language preference
PROVIDER:      Name, specialty
CLINICAL FOCUS: Condition, medications, key mechanism
CONTENT TONE:  What kind of interactions (e.g., virtual rounding,
               care coordination, discharge planning)
EXCLUSIONS:    What NOT to generate (e.g., EHR/EMR clinical data)
```

When the user requests a persona change without using this format, respond with the template above and ask them to confirm or fill in any missing fields before proceeding. This ensures all content modules get the right clinical context.

**Files affected by persona changes** (all must be updated together):

| File | What to change |
|------|---------------|
| `m365seed/data/themes.json` | Roles, user_profiles, mail_threads, calendar_events, teams_channels, chat_conversations, sharepoint_sites, planner_plans |
| `m365seed/templates/healthcare/*.j2` | discharge_planning, handoff_checklist, email_body templates |
| `seed-config.yaml` | Role references, chat messages, teams post messages |
| `seed-config.example.yaml` | Role reference |
| `m365seed/setup.py` | DEFAULT_USERS role mapping |
| `tests/test_profiles.py` | Profile assertions and test fixtures |

**Active personas** (healthcare theme):

| Role | Patient | Provider | Specialty |
|------|---------|----------|-----------|
| Care Manager — Dr. Donald Wilson | Jennifer Moore | Dr. Donald Wilson | General |
| Care Manager — Dr. Daniel Rodriguez | Elizabeth Brown (age 62, Lansing MI, Spanish) | Dr. Daniel Rodriguez | Cardiology |

After changes, always run `python -m pytest tests/ -v` to verify.

---

## Decision Log

| Date | Decision | Rationale |
|------|----------|-----------|
| 2026-02-26 | Use `httpx` over `requests` | Better async support potential, type hints |
| 2026-02-26 | Use `typer` for CLI | Rich help output, less boilerplate than argparse |
| 2026-02-26 | Deterministic subject tags | Enables idempotent mail seeding without external state |
| 2026-02-26 | Run-ID file prefix | Enables idempotent file seeding and easy cleanup |
| 2026-02-26 | Teams behind feature flag | Graph Teams APIs require `/beta` for some operations |
| 2026-02-26 | Cleanup mode implemented | All seeded content is tagged, making cleanup feasible |
| 2026-02-26 | Theme-aware content system | `theme_content.py` centralises typed accessors; all 7 modules enriched |
| 2026-02-26 | Interactive setup wizard | `m365seed setup` generates config via CLI prompts — reduces onboarding friction |
| 2026-02-26 | Custom Dockerfile dev container | Pre-bakes all deps; Azure CLI + Python 3.12 + project install |
| 2026-02-26 | Automated app registration | `m365seed register` via Azure CLI + device-code; zero portal needed |

---

## Out of Scope
- Real patient data or PHI — never, under any circumstances
- Production tenant operations — this tool is for demo/dev tenants only
- Advanced Teams features (tabs, apps) — beyond the seeding scope
- OAuth delegated flows in automation — service principal is the primary auth mode
