# SKILL.md — Capability Layer

> **Purpose**: Encodes reusable execution knowledge for the M365 Demo Tenant Seeding Tool.
> **Load Behavior**: Loaded when relevant to the task at hand.

---

## Skill: Microsoft Graph API Seeding

### When to Use
- Populating a Microsoft 365 demo tenant with synthetic content
- Demonstrating Work IQ, Foundry IQ, or Fabric IQ in HLS contexts
- Testing Graph API integrations in non-production environments

### Prerequisites
1. Python 3.11+ installed (or use the dev container)
2. Entra ID App Registration with required permissions (see README.md)
3. Client secret stored in `M365SEED_CLIENT_SECRET` env var
4. `seed-config.yaml` configured with target tenant and users

### Execution Playbook

#### Step 0: Environment Setup (Dev Container)
```bash
# Open the repo in VS Code, then Ctrl+Shift+P →
# "Dev Containers: Reopen in Container"
# Container auto-installs all deps and runs tests.
```

#### Step 1: Interactive Setup
```bash
m365seed setup
# Wizard walks through: tenant → app registration → theme → users → modules
# Generates seed-config.yaml automatically
```

#### Step 2: Validate Environment
```bash
m365seed validate -c seed-config.yaml
```
- Confirms: config schema, auth, user existence
- Fix any reported issues before proceeding

#### Step 3: Dry Run
```bash
m365seed seed-all -c seed-config.yaml --dry-run -v
```
- Review every intended action in the output
- No changes are made to the tenant

#### Step 4: Seed Content
```bash
m365seed seed-all -c seed-config.yaml -v --log-file logs/seed.jsonl
```
- Monitor the log output for errors
- Check the tenant to verify content was created

#### Step 5: Verify Idempotency
```bash
# Run again — should show "already_exists" for all items
m365seed seed-all -c seed-config.yaml -v
```

#### Step 6: Cleanup (when done)
```bash
m365seed cleanup -c seed-config.yaml --dry-run
m365seed cleanup -c seed-config.yaml
```

### Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|-------------|-----|
| `401 Unauthorized` | Bad or expired token | Regenerate client secret, check env var |
| `403 Forbidden` | Missing Graph permissions | Add permissions + grant admin consent |
| `429 Too Many Requests` | Throttled by Graph | Tool retries automatically; reduce batch size if persistent |
| `404 Not Found` on user | UPN doesn't exist in tenant | Update config with correct UPNs |
| Config validation error | Schema mismatch | Compare against `seed-config.example.yaml` |

### Key Patterns

#### Idempotency via Deterministic Tags
- **Email**: Subject prefix `[DEMO-SEED:<run_id>:<thread_id>]`
- **Files**: Filename prefix `<run_id>_`
- **Calendar**: Subject prefix `[DEMO-SEED:<run_id>:<event_id>]`
- **Teams**: Channel display name match

#### Retry Strategy
- HTTP 429 → sleep for `Retry-After` header value
- HTTP 503/504 → retry with exponential back-off
- Transport errors → retry up to 5 times
- Non-retryable errors (400, 401, 403) → raise immediately

#### Theme System
All 7 seeding modules are theme-aware via `theme_content.py`:
- `get_file_manifest(theme)` — file manifests per vertical
- `get_mail_threads(theme)` — email attachment metadata
- `get_calendar_events(theme)` — event body text
- `get_teams_channels(theme)` — channel descriptions + posts
- `get_chat_conversations(theme)` — chat messages
- `get_sharepoint_sites(theme)` — site pages + documents
- `get_planner_plans(theme)` — plan buckets + tasks

Themes: `healthcare` (default), `pharma`, `medtech`, `payor`

#### Adding a New Theme
1. Create folder: `m365seed/templates/<theme_name>/`
2. Add email templates (3 variants)
3. Add document templates (7 types)
4. Add theme data to `m365seed/data/themes.json` (all 7 sections)
5. Add theme to `CONFIG_SCHEMA` enum in `m365seed/config.py`

### Graph API Quick Reference

| Operation | Method | Endpoint |
|-----------|--------|----------|
| Send mail | POST | `/v1.0/users/{upn}/sendMail` |
| Search mail | GET | `/v1.0/users/{upn}/messages?$search=...` |
| Delete mail | DELETE | `/v1.0/users/{upn}/messages/{id}` |
| Upload file (OneDrive) | PUT | `/v1.0/users/{upn}/drive/root:/{path}:/content` |
| List folder | GET | `/v1.0/users/{upn}/drive/root:/{path}:/children` |
| Delete file | DELETE | `/v1.0/users/{upn}/drive/items/{id}` |
| Create event | POST | `/v1.0/users/{upn}/events` |
| List events | GET | `/v1.0/users/{upn}/events?$filter=...` |
| Delete event | DELETE | `/v1.0/users/{upn}/events/{id}` |
| Create channel | POST | `/beta/teams/{id}/channels` |
| Post message | POST | `/beta/teams/{id}/channels/{id}/messages` |
