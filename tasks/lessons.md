# Lessons Learned

_This file tracks patterns and corrections to prevent repeated mistakes._

## Session Start Checklist
- Review this file before starting work
- Check `tasks/todo.md` for current progress

## 2026-02-27 — Autonomous Repair Loop Lessons
- For unattended runs, always verify `M365SEED_CLIENT_SECRET` is exported in the current shell before `validate`/`seed-all`.
- Do not run cleanup with `--team-group` if `seed-config.yaml` points to a persistent existing Team; deleting it causes downstream `team_id`/`group_id` drift.
- Teams channel idempotency checks must handle pagination; single-page checks can miss existing channels and trigger duplicate-name `400` errors.
- `AzureCliCredential` requires a single scope per token request; multi-scope delegated requests need fallback handling.
- `AzureCliCredential` shells out through `az account get-access-token`, so it needs the same Azure CLI HTTP-cache purge guard as direct `az` subprocess calls.
- Delegated device-code fallback must use a public client application ID; reusing a confidential app registration's client ID causes `AADSTS7000218` unless public client flows are explicitly enabled.
- App-only token acquisition and `GET /organization` authorization are separate checks; treat a `403` there as a permission diagnostic, not a generic auth failure.
- In this tenant, Teams/Chats message posting surfaced authorization constraints even with additional app roles; treat posting failures separately from create/seed idempotency.
- OneDrive provisioning is per user: opening SharePoint/OneDrive as `admin@M365CPI56568282.onmicrosoft.com` provisioned only the admin personal site, so a configured `target_user` that still returns `404` must either sign in separately or be swapped to a provisioned account for live seeding.
