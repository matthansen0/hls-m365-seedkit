# M365 Seed — Microsoft 365 Demo Tenant Seeding Tool

**Safe, idempotent, configurable** seeding tool to populate a Microsoft 365 demo tenant with **synthetic, theme-aware** content for Work IQ, Foundry IQ, and Fabric IQ demonstrations across four Healthcare and Life Sciences verticals.

> **All content is explicitly synthetic. No real patient data, no PHI.**

---

## Features

| Command | Description | Status |
|---------|-------------|--------|
| `m365seed setup` | Interactive setup wizard — generates `seed-config.yaml` | GA |
| `m365seed validate` | Validate config, auth, permissions, and users | GA |
| `m365seed seed-profiles` | Brand user profiles (jobTitle, department, company) to match theme | GA |
| `m365seed seed-mail` | Send synthetic email threads with attachments | GA |
| `m365seed seed-files` | Upload documents to OneDrive / SharePoint | GA |
| `m365seed seed-calendar` | Create calendar events (with optional Teams meeting links) | GA |
| `m365seed seed-teams` | Seed Teams channels / posts | Beta (`--enable-beta-teams`) |
| `m365seed seed-chats` | Seed Teams 1:1 and group chats | Beta (`--enable-beta-teams`) |
| `m365seed seed-sharepoint` | Create SharePoint sites, pages, and documents | GA |
| `m365seed seed-planner` | Create Planner plans, buckets, and tasks | GA |
| `m365seed seed-all` | Run all seeders in sequence | GA |
| `m365seed cleanup` | Remove all seeded content by run_id | GA |

### Content Themes

| Theme | Flag | Description |
|-------|------|-------------|
| `healthcare` | (default) | Health Provider — clinical ops, care coordination |
| `pharma` | `--theme pharma` | Pharma / Life Science — research, clinical trials |
| `medtech` | `--theme medtech` | MedTech — product dev, manufacturing |
| `payor` | `--theme payor` | Health Payor — claims, member services |

---

## Quick Start

Everything runs inside the dev container — no local Python, pip, or package installs required.

### 1. Prerequisites

- [VS Code](https://code.visualstudio.com/) with the [Dev Containers extension](https://marketplace.visualstudio.com/items?itemName=ms-vscode-remote.remote-containers)
- [Docker Desktop](https://www.docker.com/products/docker-desktop) running
- An Entra ID App Registration in the demo tenant — **create one automatically** with `m365seed register`, or set up manually (see [Graph Permissions](#microsoft-graph-permissions))

### 2. Set Environment Variables (if using an existing app)

If you already have an app registration, set these **before** reopening in the container — the dev container forwards them automatically. If you plan to use `m365seed register` or the setup wizard's auto-register option, you can skip this step.

```powershell
# PowerShell
$env:M365SEED_CLIENT_SECRET = "your-client-secret"
$env:M365SEED_TENANT_ID     = "your-tenant-guid"      # optional — wizard prompts if missing
$env:M365SEED_CLIENT_ID     = "your-app-client-id"     # optional — wizard prompts if missing
```
```bash
# macOS / Linux
export M365SEED_CLIENT_SECRET="your-client-secret"
export M365SEED_TENANT_ID="your-tenant-guid"
export M365SEED_CLIENT_ID="your-app-client-id"
```

### 3. Open in Dev Container

```
Ctrl+Shift+P  →  "Dev Containers: Reopen in Container"
```

The container auto-builds with Python 3.12, Azure CLI, and all project dependencies.  
On first launch it installs the project, copies the example config, and runs the test suite.

### 4. Run the Setup Wizard

```bash
m365seed setup
```

The wizard walks through tenant ID → app registration → theme → demo users → optional user password reset → content modules, then generates `seed-config.yaml`.

Passwords are never written to `seed-config.yaml`; the optional reset step updates selected tenant users directly via Azure CLI.

### 5. Seed the Tenant

```bash
m365seed seed-all --dry-run       # Preview actions (no changes)
m365seed seed-all                 # Go live
```

### 6. Cleanup (when done)

```bash
m365seed cleanup --dry-run        # Preview deletions
m365seed cleanup                  # Remove seeded content
```

Cleanup supports **7 content types** — each can be toggled independently:

| Flag | Content Type | Cleanup Strategy |
|------|-------------|-----------------|
| `--mail` / `--no-mail` | Email | Delete messages with `DEMO-SEED:<run_id>` in subject |
| `--files` / `--no-files` | OneDrive files | Delete files prefixed with `<run_id>_` |
| `--calendar` / `--no-calendar` | Calendar events | Delete events with `[DEMO-SEED:<run_id>:` in subject |
| `--teams` / `--no-teams` | Teams channels | Delete channels matching configured display_name (beta) |
| `--chats` / `--no-chats` | Teams chats | Delete group chats with `DEMO-SEED:<run_id>` in topic (beta) |
| `--sharepoint` / `--no-sharepoint` | SharePoint sites | Delete M365 Groups (cascades to site, pages, docs) |
| `--planner` / `--no-planner` | Planner plans | Delete plans with `[DEMO-SEED:<run_id>]` prefix (cascades to buckets/tasks) |

---

## Microsoft Graph Permissions

### App Registration Setup

#### Option A — Automated (recommended)

Run the registration command inside the dev container:

```bash
m365seed register               # Standalone — creates app, adds permissions, grants consent
m365seed setup                   # Or use the setup wizard — it offers auto-registration at Step 2
```

This uses Azure CLI with device-code login to:
1. Authenticate as a Global Administrator
2. Create the app registration (`M365 Demo Seed Tool`, single-tenant)
3. Add all required Graph API permissions (Application type)
4. Create a service principal and client secret
5. Grant admin consent

> **Requires**: Azure CLI (`az`) — pre-installed in the dev container.

#### Option B — Manual (Azure Portal)

1. Go to [Azure Portal → Entra ID → App registrations](https://portal.azure.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade)
2. Click **New registration**
3. Name: `M365 Demo Seed Tool`
4. Supported account types: **Single tenant**
5. Click **Register**
6. Note the **Application (client) ID** and **Directory (tenant) ID**
7. Under **Certificates & secrets**, create a new client secret
8. Under **API permissions**, add the permissions below

### Required Permissions (Application type)

| Feature | Permission | Type | Justification |
|---------|-----------|------|---------------|
| Validate | `User.Read.All` | Application | Verify demo users exist |
| Profiles | `User.ReadWrite.All` | Application | Update user jobTitle, department, companyName, officeLocation |
| Validate | `Organization.Read.All` | Application | Verify auth works |
| Email | `Mail.Send` | Application | Send mail as demo users |
| Email | `Mail.ReadWrite` | Application | Idempotency checks + cleanup |
| Files (OneDrive) | `Files.ReadWrite.All` | Application | Upload files to users' OneDrive |
| Files (SharePoint) | `Sites.ReadWrite.All` | Application | Upload files to SharePoint |
| Calendar | `Calendars.ReadWrite` | Application | Create/read/delete calendar events |
| Calendar | `OnlineMeetings.ReadWrite.All` | Application | Create Teams meeting links on events |
| Teams (beta) | `ChannelMessage.Send` | Application | Post messages to channels |
| Teams (beta) | `Channel.Create` | Application | Create channels |
| Teams (beta) | `Channel.Delete.All` | Application | Cleanup: delete seeded channels |
| Teams Chats (beta) | `Chat.Create` | Application | Create 1:1 and group chats |
| Teams Chats (beta) | `Chat.ReadWrite.All` | Application | Send messages & cleanup chats |
| SharePoint | `Group.ReadWrite.All` | Application | Create M365 Groups (auto-provisions sites) |
| SharePoint | `Sites.Manage.All` | Application | Create/delete site pages |
| Planner | `Tasks.ReadWrite.All` | Application | Create/delete plans, buckets, tasks |

> **Least privilege**: Only grant the permissions for the features you plan to use.
> Teams / Chat permissions are only needed if you enable `--enable-beta-teams`.

After adding permissions, click **Grant admin consent** for the tenant.

---

## Running Against MDX/CDX Tenants

1. **MDX (Microsoft Demo Experience)** or **CDX (Customer Digital Experience)** tenants are pre-provisioned Microsoft 365 environments for demos.
2. Create your App Registration **inside the demo tenant** (not your corporate tenant).
3. Use the demo tenant's **tenant ID** in `seed-config.yaml`.
4. Demo users (e.g., `AllanDe@M365x...`, `MeganB@M365x...`) come pre-provisioned.
   - Update user UPNs in `seed-config.yaml` to match the actual demo users.
5. **Safety**: The tool is idempotent — rerunning will skip already-created content. Use `cleanup` to remove seeded content.

---

## Known Issues

### OneDrive Provisioning Requirement

Microsoft 365 does not provision a user's OneDrive personal site until that user (or an admin) accesses it for the first time. If the `seed-files` command returns a `404` for a user's drive, the target user must sign in to OneDrive (or SharePoint) at least once before file seeding will work.

**Workarounds:**
- Sign in as the target user at `https://<tenant>-my.sharepoint.com/` to trigger provisioning
- Or use `m365seed setup` — the wizard lets you pick a user whose OneDrive is already provisioned

See [Microsoft Learn — OneDrive provisioning](https://learn.microsoft.com/en-us/sharepoint/troubleshoot/administration/personal-site-not-created) for details.

---

## Idempotency Strategy

The tool uses **deterministic tagging** to ensure reruns are safe:

### Email
- Subject prefix: `[DEMO-SEED:<run_id>:<thread_id>]`
- Custom headers: `X-DemoSeed-RunId`, `X-DemoSeed-ThreadId`
- Before sending, the tool searches the sender's mailbox for existing messages with the tag

### Files
- Filenames are prefixed with `<run_id>_` (e.g., `hls-demo-001_Standard_Operating_Procedures.txt`)
- Before uploading, the tool checks if the file already exists in OneDrive/SharePoint

### Calendar
- Subjects are prefixed: `[DEMO-SEED:<run_id>:<event_id>]`
- Before creating, the tool queries existing events with the same prefix
- Online meetings (`is_online_meeting: true`) create Teams join links via v1.0 GA

### Teams Channels
- Channel names are checked for existence before creation
- Reuses existing channels instead of creating duplicates

### Teams Chats
- Group chat topics are prefixed: `[DEMO-SEED:<run_id>] <topic>`
- 1:1 chats are created each time (no dedup — chats lack unique identifiers)

### SharePoint Sites
- Sites created via M365 Groups; group `displayName` prefixed: `[DEMO-SEED:<run_id>]`
- Group `mailNickname` prefixed with `seed<run_id>` for uniqueness checks
- Pages and documents are tagged with run_id for cleanup

### Planner
- Plan titles are prefixed: `[DEMO-SEED:<run_id>] <title>`
- Before creating, the tool queries existing plans in the group

---

## Configuration Reference

See [seed-config.example.yaml](seed-config.example.yaml) for the full schema.

Key sections:
- `tenant` — Tenant ID and authority URL
- `auth` — Authentication mode and client credentials
- `targets.users` — List of demo users (UPN + role)
- `content` — Theme (`healthcare`, `pharma`, `medtech`, `payor`) and `run_id`
- `mail.threads` — Email threads to seed
- `files` — OneDrive and SharePoint file seeding config
- `calendar` — Calendar events config (supports `is_online_meeting` for Teams links)
- `teams` — Teams channels/posts (beta)
- `chats` — Teams 1:1 and group chats with messages (beta)
- `sharepoint` — SharePoint sites, pages, and document uploads
- `planner` — Planner plans, buckets, and tasks

---

## CLI Reference

```
m365seed --help
m365seed validate --help
m365seed seed-profiles --help
m365seed seed-mail --help
m365seed seed-files --help
m365seed seed-calendar --help
m365seed seed-teams --help
m365seed seed-chats --help
m365seed seed-sharepoint --help
m365seed seed-planner --help
m365seed seed-all --help
m365seed cleanup --help
```

### Common Flags

| Flag | Description |
|------|-------------|
| `--config`, `-c` | Path to config YAML (default: `seed-config.yaml`) |
| `--dry-run` | Print actions without modifying the tenant |
| `--verbose`, `-v` | Debug-level logging |
| `--log-file` | Path for JSONL structured log output |
| `--theme` | Override content theme (seed commands only) |
| `--enable-beta-teams` | Enable Teams seeding (beta APIs) |

---

## Development

All development happens inside the dev container. Open in VS Code → "Reopen in Container" — deps are pre-installed and tests run automatically on first launch.

```bash
# Run tests
pytest -v

# Run with coverage
pytest --cov=m365seed --cov-report=term-missing
```

---

## Project Structure

```
HLS-M365-Seed/
├── .devcontainer/
│   ├── Dockerfile           # Python 3.12 + Azure CLI + all deps
│   ├── devcontainer.json    # VS Code dev-container config
│   └── post-create.sh       # Auto-setup on container creation
├── README.md
├── AGENTS.md
├── spec.md
├── SKILL.md
├── pyproject.toml
├── seed-config.example.yaml
├── .gitignore
├── m365seed/
│   ├── __init__.py
│   ├── cli.py               # Typer CLI entry point
│   ├── setup.py             # Interactive setup wizard
│   ├── config.py            # YAML schema validation
│   ├── graph.py             # Graph client, retry, auth
│   ├── theme_content.py     # Theme content provider (typed, LRU-cached)
│   ├── mail.py              # Email seeding (theme-aware)
│   ├── files.py             # File seeding (OneDrive + SharePoint)
│   ├── calendar.py          # Calendar seeding (+ online meetings)
│   ├── teams.py             # Teams channel/post seeding (beta)
│   ├── chats.py             # Teams 1:1/group chat seeding (beta)
│   ├── sharepoint.py        # SharePoint sites, pages, documents
│   ├── planner.py           # Planner plans, buckets, tasks
│   ├── cleanup.py           # Cleanup mode (7 content types)
│   ├── templates/           # Jinja2 templates per theme
│   │   ├── healthcare/      # 3 email + 7 doc templates
│   │   ├── pharma/          # 3 email + 7 doc templates
│   │   ├── medtech/         # 3 email + 7 doc templates
│   │   └── payor/           # 3 email + 7 doc templates
│   └── data/                # Synthetic data payloads
│       └── themes.json      # 83 KB — all theme content (4 verticals)
├── tests/
│   ├── test_config.py
│   ├── test_graph.py
│   ├── test_mail.py
│   ├── test_files.py
│   ├── test_calendar.py
│   ├── test_chats.py
│   ├── test_sharepoint.py
│   ├── test_planner.py
│   └── test_cleanup.py
```

---

## License

MIT — See [LICENSE](LICENSE).

---

> **Disclaimer**: All content generated by this tool is synthetic and intended solely for demo environments. No real patient data, no PHI, no PII. All names, organizations, and clinical data are fictitious.
