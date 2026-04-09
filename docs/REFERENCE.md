# Reference

Detailed reference for Graph permissions, cleanup, idempotency, CLI flags, configuration, known issues, and project structure.

---

## Graph Permissions

All permissions are **Application** type. Only grant permissions for the modules you plan to use.

| Module | Permission | Why |
|--------|-----------|-----|
| Validate | `User.Read.All` | Verify demo users exist |
| Validate | `Organization.Read.All` | Verify auth works |
| Profiles | `User.ReadWrite.All` | Update jobTitle, department, companyName |
| Email | `Mail.Send` | Send as demo users |
| Email | `Mail.ReadWrite` | Idempotency checks + cleanup |
| Files | `Files.ReadWrite.All` | Upload to OneDrive |
| Files | `Sites.ReadWrite.All` | Upload to SharePoint |
| Calendar | `Calendars.ReadWrite` | Create/delete events |
| Calendar | `OnlineMeetings.ReadWrite.All` | Teams meeting links |
| Teams | `ChannelMessage.Send` | Post to channels |
| Teams | `Channel.Create` | Create channels |
| Teams | `Channel.Delete.All` | Cleanup channels |
| Chats | `Chat.Create` | Create 1:1/group chats |
| Chats | `Chat.ReadWrite.All` | Send messages + cleanup |
| SharePoint | `Group.ReadWrite.All` | Create M365 Groups → sites |
| SharePoint | `Sites.Manage.All` | Create/delete pages |
| Planner | `Tasks.ReadWrite.All` | Create/delete plans, buckets, tasks |

After adding permissions, **Grant admin consent** for the tenant.

---

## Cleanup Flags

Each content type can be toggled independently:

| Flag | Cleanup Strategy |
|------|-----------------|
| `--mail` / `--no-mail` | Delete messages with `DEMO-SEED:<run_id>` in subject |
| `--files` / `--no-files` | Delete files prefixed with `<run_id>_` |
| `--calendar` / `--no-calendar` | Delete events with `[DEMO-SEED:<run_id>:` in subject |
| `--teams` / `--no-teams` | Delete channels matching configured names |
| `--chats` / `--no-chats` | Delete group chats with `DEMO-SEED:<run_id>` in topic |
| `--sharepoint` / `--no-sharepoint` | Delete M365 Groups (cascades to site, pages, docs) |
| `--planner` / `--no-planner` | Delete plans with `[DEMO-SEED:<run_id>]` prefix |

---

## Idempotency

The tool is safe to rerun. Every content item is tagged with a deterministic identifier so duplicates are detected and skipped.

| Content | Tagging Strategy |
|---------|-----------------|
| Email | Subject prefix `[DEMO-SEED:<run_id>:<thread_id>]` + custom headers |
| Files | Filename prefix `<run_id>_` |
| Calendar | Subject prefix `[DEMO-SEED:<run_id>:<event_id>]` |
| Teams Channels | Match by `display_name` |
| Teams Chats | Topic prefix `[DEMO-SEED:<run_id>]` (1:1 chats can't be deduped) |
| SharePoint Sites | Group `mailNickname` prefix `seed<run_id>` |
| Planner | Plan title prefix `[DEMO-SEED:<run_id>]` |

---

## CLI Flags

| Flag | Description |
|------|-------------|
| `--config`, `-c` | Path to config YAML (default: `seed-config.yaml`) |
| `--dry-run` | Preview actions without modifying the tenant |
| `--verbose`, `-v` | Debug-level logging |
| `--log-file` | Path for JSONL structured log |
| `--theme` | Override content theme |
| `--enable-beta-teams` | Enable Teams/Chats seeding (beta Graph APIs) |

---

## Configuration

See [seed-config.example.yaml](../seed-config.example.yaml) for the full schema. Key sections:

| Section | Purpose |
|---------|---------|
| `tenant` | Tenant ID and authority URL |
| `auth` | Auth mode, client ID, secret env var name |
| `targets.users` | Demo users (UPN + role) |
| `content` | Theme and `run_id` |
| `mail.threads` | Email threads to seed |
| `files` | OneDrive/SharePoint file config |
| `calendar` | Calendar events |
| `teams` | Teams channels and posts |
| `chats` | Teams 1:1 and group chats |
| `sharepoint` | SharePoint sites, pages, documents |
| `planner` | Planner plans, buckets, tasks |

---

## Known Issues

**OneDrive provisioning** — M365 doesn't provision a user's OneDrive until they (or an admin) access it for the first time. If `seed-files` returns a 404, sign in as the target user at `https://<tenant>-my.sharepoint.com/` to trigger provisioning. See [Microsoft Learn](https://learn.microsoft.com/en-us/sharepoint/troubleshoot/administration/personal-site-not-created).

**Teams beta APIs** — Teams channel posting and chat creation use Microsoft Graph beta endpoints. These require `--enable-beta-teams` and may have authorization constraints depending on tenant configuration.

---

## Project Structure

```
hls-m365-seedkit/
├── .devcontainer/
│   ├── Dockerfile           # Python 3.12 + Azure CLI
│   ├── devcontainer.json    # VS Code config
│   └── post-create.sh       # Auto-setup
├── m365seed/
│   ├── cli.py               # Typer CLI
│   ├── setup.py             # Interactive wizard
│   ├── config.py            # YAML validation
│   ├── graph.py             # Graph client + retry
│   ├── theme_content.py     # Theme content provider
│   ├── mail.py              # Email seeding
│   ├── files.py             # OneDrive/SharePoint files
│   ├── calendar.py          # Calendar events
│   ├── teams.py             # Teams channels
│   ├── chats.py             # Teams chats
│   ├── sharepoint.py        # SharePoint sites
│   ├── planner.py           # Planner boards
│   ├── cleanup.py           # Cleanup (7 types)
│   ├── templates/           # Jinja2 templates (4 themes)
│   └── data/themes.json     # Synthetic content (4 verticals)
└── tests/                   # 180 unit tests (mocked Graph)
```
