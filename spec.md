# spec.md — Feature Layer

> **Purpose**: Defines the intent and scope of the M365 Demo Tenant Seeding Tool.
> **Load Behavior**: Loaded when starting a feature or reviewing scope.

---

## Problem Statement

Microsoft Healthcare and Life Sciences (HLS) demos require a populated Microsoft 365 tenant with realistic-looking content — emails, files, calendar events, and Teams messages — to demonstrate Work IQ, Foundry IQ, and Fabric IQ capabilities. Manually creating this content is tedious, error-prone, and not reproducible.

## Solution

A Python CLI tool (`m365seed`) that programmatically seeds a demo tenant with **synthetic, theme-aware content** via the Microsoft Graph API across four HLS verticals. The tool is:

- **Safe**: All content is explicitly synthetic — no real data or PHI.
- **Idempotent**: Deterministic tagging ensures reruns don't create duplicates.
- **Configurable**: YAML-driven configuration supports multiple themes and scenarios.
- **Reversible**: Cleanup mode removes all seeded content by run ID.
- **Automated**: Interactive setup wizard (`m365seed setup`) and dev container streamline deployment.

## Scope

### In Scope

| Capability | Graph Endpoint | API Version |
|---|---|---|
| Email threads with attachments | `/users/{id}/sendMail` | v1.0 |
| OneDrive file uploads | `/users/{id}/drive/root:/{path}:/content` | v1.0 |
| SharePoint file uploads | `/sites/{id}/drives/{id}/root:/{path}:/content` | v1.0 |
| Calendar events (single + recurring) | `/users/{id}/events` | v1.0 |
| Teams channels and posts | `/teams/{id}/channels` | **beta** |
| Validation (auth, users, config) | `/organization`, `/users/{upn}` | v1.0 |
| Cleanup (mail, files, calendar) | DELETE endpoints | v1.0 |

### Content Themes

| Theme | Vertical | Key Artifacts |
|---|---|---|
| `healthcare` | Health Provider | SOPs, discharge planning, staffing, compliance |
| `pharma` | Life Science | Research protocols, clinical trial reports, lab data |
| `medtech` | Medical Devices | Design reviews, manufacturing QA, 510(k) prep |
| `payor` | Health Insurance | Claims ops, member services, network management |

### Out of Scope
- Real patient data or PHI
- Production tenant targeting
- Advanced Teams apps/tabs/bots
- Power Platform / Dynamics 365 seeding
- Multi-tenant support (one tenant per config)

## Architecture

```
CLI (typer)
 └── Commands: setup | validate | seed-* | cleanup
      └── Setup Wizard (interactive config generation)
      └── GraphClient (httpx + azure-identity)
           ├── Auth: ClientSecretCredential or DeviceCodeCredential
           ├── Retry: exponential back-off on 429/503/504
           └── Dry-run: log-only mode
      └── ThemeContent (typed, LRU-cached accessors per vertical)
      └── Templates (Jinja2, per theme — 4 verticals × 10 templates)
      └── Config (YAML, validated via jsonschema)

Dev Container
 └── Dockerfile (Python 3.12 + Azure CLI + all deps)
 └── post-create.sh (auto-install, config copy, test run)
```

## Acceptance Criteria

1. `m365seed validate` succeeds with a valid config (placeholders OK for dry-run).
2. `--dry-run` prints all intended actions without modifying the tenant.
3. Rerunning `seed-mail` and `seed-files` does **not** create duplicates.
4. Graph throttling (429) is handled with `Retry-After` back-off.
5. All Graph permissions are documented in README.md.
6. Unit tests pass for config validation, retry logic, and idempotency.
7. Teams seeding is behind `--enable-beta-teams` and labeled as unstable.
8. Cleanup mode removes seeded content tagged with the run ID.

## Dependencies

| Package | Purpose |
|---|---|
| `azure-identity` | Entra ID auth (ClientSecretCredential / DeviceCodeCredential) |
| `httpx` | HTTP client for Graph API calls |
| `pyyaml` | YAML config parsing |
| `typer` | CLI framework |
| `jinja2` | Template rendering for emails and documents |
| `jsonschema` | Config schema validation |
| `rich` | Pretty console output and logging |
| `python-docx` | (Optional) DOCX file generation |
