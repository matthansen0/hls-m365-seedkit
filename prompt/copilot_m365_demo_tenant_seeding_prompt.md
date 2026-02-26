# Copilot Agent Mode Prompt — M365 Demo Tenant Seeding (Healthcare-Themed, Synthetic)

---

You are GitHub Copilot running in **Agent mode**. Create a new repository (or implement in the current repo) that provides a **safe, idempotent, configurable seeding tool** to populate a Microsoft 365 demo tenant with **synthetic healthcare-themed** content: emails, calendar events (optional), and OneDrive/SharePoint files; optionally Teams channels/posts where feasible. It is important to know the purpose and theme - this environment will be used solely to fund the functionality of Microsoft Work IQ, and ultimately integrate with Foundry IQ and Fabric IQ, all with a Healthcare and Life Sciences specific focus.

## Hard constraints (must follow)
1. **No real customer data, no PHI.** All content must be explicitly synthetic and clearly labeled in templates.
2. **Idempotent**: the tool must be safe to rerun without duplicating everything (use a deterministic tagging strategy).
3. **Least privilege**: document required Microsoft Graph permissions; do not request more than needed.
4. **No undocumented APIs**: only use Microsoft Graph endpoints documented in Microsoft Learn. If a Teams operation requires `/beta`, isolate it behind a feature flag and label it as “beta”.
5. **No secrets committed**: all auth must come from environment variables or a local config excluded by `.gitignore`.
6. Include a **dry-run mode** that prints intended actions but does not modify the tenant.
7. Use a dev container for everything, I don't want to install prereqs.
8. We need to make sure that all the content that's created (job titles, messages, files, meetings, sites, etc. etc.) is all full fledged contextually relevant to the themes. Including data generated in files created in OneDrive, etc. This might be more complicated to do, but when I merge this environment with Fabric IQ, Fabric will have all the theme-relevant content generated (by another project).
9. The deployment process should be as automated as possible, even better if it was CLI prompt driven.

## Nice to have
1. Determine if it is possible to include a **cleanup mode** mode that allows for removal of all created; if so, implement this functionality.
2. Include a flag, and structure, for Health Provider mode (default), but also Pharma/Life Science (research, lab), MedTech (product, manufacturing), and Health Payor.


## Target platform & stack
- Language: **Python 3.11+**
- Auth: `azure-identity` with **ClientSecretCredential** (service principal) OR device code (optional)
- HTTP: `requests` or `httpx`
- Config: YAML (`seed-config.yaml`)
- CLI: `typer` (or `argparse`)
- Logging: structured logs to console + optional JSONL file

## Capabilities (deliverables)
Implement a CLI named `m365seed` with commands:

### 1) `m365seed validate`
- Validates config schema
- Validates Graph auth works
- Checks required permissions are present (best-effort)
- Verifies the target users exist in the tenant

### 2) `m365seed seed-mail`
- Sends synthetic “healthcare” email threads between specified demo users
- Supports attachments (small text/docx/pdf placeholders)
- Uses deterministic message identifiers:
  - e.g., add a unique header `X-DemoSeed-RunId` and `X-DemoSeed-ThreadId`
  - and/or add a subject prefix like `[DEMO-SEED:<threadId>]`
- Must support batching and throttling (handle 429 with `Retry-After`)

### 3) `m365seed seed-files`
- Uploads healthcare-themed files into:
  - OneDrive of specific users and/or a target SharePoint site library
- Creates folder structure (e.g., `Clinical Ops/`, `Care Coordination/`, `Compliance/`)
- Uses deterministic filenames and a manifest to avoid duplicates
- Adds simple metadata where possible (e.g., file descriptions in a manifest)

### 4) `m365seed seed-calendar` (optional but preferred)
- Creates synthetic meetings (e.g., “Clinical Ops Review”, “Care Coordination Standup”)
- Creates them in selected users’ calendars with deterministic subjects for idempotency

### 5) `m365seed seed-teams` (optional, behind flag `--enable-beta-teams`)
- Creates channels/posts or chat messages only if supported
- If it requires Microsoft Graph `/beta`, implement but label as unstable and **off by default**

## Repo structure
Create this structure:

- `README.md` (clear setup + examples + permissions)
- `pyproject.toml` (dependency management)
- `m365seed/`
  - `__init__.py`
  - `cli.py`
  - `config.py` (YAML schema validation)
  - `graph.py` (Graph client, retry, auth)
  - `mail.py` (email seeding)
  - `files.py` (file seeding)
  - `calendar.py` (calendar seeding)
  - `teams.py` (optional; gated + beta)
  - `templates/` (Jinja2 templates for emails + docs)
  - `data/` (synthetic sample payloads)
- `tests/` with unit tests for:
  - config validation
  - idempotency key generation
  - Graph retry/backoff logic (mocked)

## Config file schema
Create `seed-config.example.yaml` with these sections:

```yaml
tenant:
  tenant_id: "<GUID>"
  authority: "https://login.microsoftonline.com/<GUID>"
auth:
  mode: "client_secret"   # or "device_code"
  client_id: "<GUID>"
  client_secret_env: "M365SEED_CLIENT_SECRET"  # env var name

targets:
  users:
    - upn: "allande@m365x123456.onmicrosoft.com"
      role: "Clinical Ops Manager"
    - upn: "dr.patel@m365x123456.onmicrosoft.com"
      role: "Physician Lead"
    - upn: "care.coord@m365x123456.onmicrosoft.com"
      role: "Care Coordinator"

content:
  theme: "healthcare"
  run_id: "hls-demo-001"

mail:
  threads:
    - thread_id: "care-coordination-001"
      subject: "Care Coordination: Discharge planning follow-up"
      participants: ["dr.patel@...", "care.coord@..."]
      messages: 6
      include_attachments: true

files:
  oneDrive:
    enabled: true
    folders:
      - "Clinical Ops"
      - "Care Coordination"
      - "Compliance"
  sharePoint:
    enabled: false

calendar:
  enabled: true

teams:
  enabled: false
```

## Synthetic healthcare content requirements
Generate templates that look realistic but explicitly synthetic:
- Use disclaimers in footers like: **“Demo content — synthetic, no patient data.”**
- Include artifacts such as:
  - SOP doc
  - compliance checklist
  - discharge planning worksheet
  - clinic staffing roster (fake)
- Keep language plausible but avoid any reference to real organizations or real people.

## Acceptance criteria
- Running `m365seed validate` succeeds with example config (with placeholders)
- Dry run prints intended actions
- Rerunning `seed-mail` and `seed-files` does **not** create duplicates
- Tool handles Graph throttling (429) and retries
- README documents:
  - Graph permissions needed per feature
  - how to create/authorize an Entra app registration
  - how to run against an MDX/CDX tenant safely

# Workflow Orchestration

## 1. Plan Mode Default
- Enter plan mode for ANY non-trivial task (3+ steps or architectural decisions)
- If something goes sideways, STOP and re-plan immediately — don’t keep pushing
- Use plan mode for verification steps, not just building
- Write detailed specs upfront to reduce ambiguity

## 2. Subagent Strategy
- Use subagents liberally to keep main context window clean
- Offload research, exploration, and parallel analysis to subagents
- For complex problems, throw more compute at it via subagents
- One task per subagent for focused execution

## 3. Self-Improvement Loop
- After ANY correction from the user: update `tasks/lessons.md` with the pattern
- Write rules for yourself that prevent the same mistake
- Ruthlessly iterate on these lessons until mistake rate drops
- Review lessons at session start for relevant project

## 4. Verification Before Done
- Never mark a task complete without proving it works
- Diff behavior between main and your changes when relevant
- Ask yourself: “Would a staff engineer approve this?”
- Run tests, check logs, demonstrate correctness

## 5. Demand Elegance (Balanced)
- For non-trivial changes: pause and ask “is there a more elegant way?”
- If a fix feels hacky: “Knowing everything I know now, implement the elegant solution”
- Skip this for simple, obvious fixes — don’t over-engineer
- Challenge your own work before presenting it

## 6. Autonomous Bug Fixing
- When given a bug report: just fix it. Don’t ask for hand-holding
- Point at logs, errors, failing tests — then resolve them
- Zero context switching required from the user
- Go fix failing CI tests without being told how

---

# Task Management

1. **Plan First:** Write plan to `tasks/todo.md` with checkable items
2. **Verify Plan:** Check in before starting implementation
3. **Track Progress:** Mark items complete as you go
4. **Explain Changes:** High-level summary at each step
5. **Document Results:** Add review section to `tasks/todo.md`
6. **Capture Lessons:** Update `tasks/lessons.md` after corrections

---

# Core Principles

- **Simplicity First:** Make every change as simple as possible. Impact minimal code.
- **No Laziness:** Find root causes. No temporary fixes. Senior developer standards.
- **Minimal Impact:** Changes should only touch what’s necessary. Avoid introducing bugs.

## Final instruction
Now implement the repo end-to-end, including tests and README examples. As you build, run unit tests locally and fix failures. If Teams seeding requires beta APIs, implement only behind `--enable-beta-teams` and document that it may break. As you go, create and modify a copilot agent file.

1. Operating Layer
File: AGENTS.md
Purpose: Defines collaboration norms and working agreements
Load Behavior: Always loaded (every session)
Analogy: Employee handbook

2. Feature Layer
File: spec.md
Purpose: Defines the intent and scope of a specific build
Load Behavior: Loaded when starting a feature
Analogy: Project brief

3. Capability Layer
File: SKILL.md
Purpose: Encodes reusable execution knowledge
Load Behavior: Loaded when relevant to the task
Analogy: Playbook