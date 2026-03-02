"""Planner seeding — create plans, buckets, and tasks.

Uses Microsoft Graph v1.0 Planner endpoints with application permission
``Tasks.ReadWrite.All``.  All content is tagged with the run_id for cleanup.
"""

from __future__ import annotations

import logging
from typing import Any

from m365seed.graph import GraphClient
from m365seed.theme_content import get_planner_plans

logger = logging.getLogger("m365seed.planner")

DISCLAIMER = "Demo content — synthetic, no patient data."


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _plan_exists(
    client: GraphClient,
    group_id: str,
    title: str,
) -> dict[str, Any] | None:
    """Check if a plan with the given title exists in the group."""
    try:
        resp = client.get(
            f"/groups/{group_id}/planner/plans",
            params={"$select": "id,title", "$top": "50"},
        )
        for plan in resp.json().get("value", []):
            if plan.get("title") == title:
                return plan
        return None
    except Exception as exc:
        logger.warning("Plan existence check failed: %s", exc)
        return None


def _create_plan(
    client: GraphClient,
    group_id: str,
    title: str,
) -> dict[str, Any]:
    """Create a Planner plan in the specified group."""
    payload = {
        "owner": group_id,
        "title": title,
    }
    resp = client.post("/planner/plans", json_body=payload)
    return resp.json()


def _create_bucket(
    client: GraphClient,
    plan_id: str,
    name: str,
    order_hint: str = " !",
) -> dict[str, Any]:
    """Create a bucket within a plan."""
    payload = {
        "name": name,
        "planId": plan_id,
        "orderHint": order_hint,
    }
    resp = client.post("/planner/buckets", json_body=payload)
    return resp.json()


def _create_task(
    client: GraphClient,
    plan_id: str,
    bucket_id: str,
    task_cfg: dict[str, Any],
    run_id: str,
    assignee_ids: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Create a Planner task in a bucket."""
    title = f"[DEMO-SEED:{run_id}] {task_cfg['title']}"

    payload: dict[str, Any] = {
        "planId": plan_id,
        "bucketId": bucket_id,
        "title": title,
    }

    # Priority: 1=Urgent, 3=Important, 5=Medium, 9=Low
    if "priority" in task_cfg:
        payload["priority"] = task_cfg["priority"]

    # Percent complete: 0, 50, 100
    if "percent_complete" in task_cfg:
        payload["percentComplete"] = task_cfg["percent_complete"]

    # Assignments
    if task_cfg.get("assignees") and assignee_ids:
        assignments = {}
        for upn in task_cfg["assignees"]:
            uid = assignee_ids.get(upn)
            if uid:
                assignments[uid] = {
                    "@odata.type": "#microsoft.graph.plannerAssignment",
                    "orderHint": " !",
                }
        if assignments:
            payload["assignments"] = assignments

    resp = client.post("/planner/tasks", json_body=payload)
    return resp.json()


def _resolve_user_id(client: GraphClient, upn: str) -> str:
    """Resolve a UPN to a directory object id."""
    resp = client.get(f"/users/{upn}", params={"$select": "id"})
    return resp.json().get("id", upn)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------


def seed_planner(
    client: GraphClient,
    cfg: dict[str, Any],
    theme: str,
    run_id: str,
) -> list[dict[str, Any]]:
    """Create Planner plans, buckets, and tasks.

    Requires a pre-existing Microsoft 365 Group (``group_id`` in config).
    Returns a list of action records with plan/task IDs for cleanup.
    """
    planner_cfg = cfg.get("planner", {})
    if not planner_cfg.get("enabled"):
        logger.info("Planner seeding is disabled — skipping.")
        return []

    group_id = planner_cfg.get("group_id", "")
    if not group_id:
        logger.warning("No group_id configured for Planner — skipping.")
        return []

    # Verify the group still exists before attempting plan creation
    if not client.dry_run:
        try:
            client.get(f"/groups/{group_id}", params={"$select": "id"})
        except Exception:
            logger.error(
                "Group '%s' not found — was it deleted during cleanup? "
                "Re-run 'm365seed setup' to recreate the team/group, "
                "then update planner.group_id in seed-config.yaml.",
                group_id,
            )
            return []

    plans = planner_cfg.get("plans", [])
    if not plans:
        logger.info("No Planner plans configured — skipping.")
        return []

    # Enrich config plans with theme-specific buckets and tasks
    theme_plans_list = get_planner_plans(theme)
    theme_plans_map = {p["title"]: p for p in theme_plans_list if "title" in p}
    for plan_cfg in plans:
        ptitle = plan_cfg.get("title", "")
        if ptitle in theme_plans_map:
            tp = theme_plans_map[ptitle]
            if not plan_cfg.get("buckets") and tp.get("buckets"):
                plan_cfg["buckets"] = tp["buckets"]

    actions: list[dict[str, Any]] = []

    # Pre-resolve user IDs for task assignments
    all_upns: set[str] = set()
    for plan_cfg in plans:
        for bucket_cfg in plan_cfg.get("buckets", []):
            for task_cfg in bucket_cfg.get("tasks", []):
                all_upns.update(task_cfg.get("assignees", []))

    assignee_ids: dict[str, str] = {}
    for upn in all_upns:
        try:
            assignee_ids[upn] = _resolve_user_id(client, upn)
        except Exception as exc:
            logger.warning("Could not resolve user '%s': %s", upn, exc)

    for plan_cfg in plans:
        plan_title = f"[DEMO-SEED:{run_id}] {plan_cfg['title']}"

        # Idempotency
        existing = (
            None if client.dry_run else _plan_exists(client, group_id, plan_title)
        )

        if existing:
            plan_id = existing["id"]
            logger.info("Plan '%s' already exists — reusing.", plan_cfg["title"])
            actions.append(
                {
                    "action": "skip_plan",
                    "plan": plan_cfg["title"],
                    "plan_id": plan_id,
                    "reason": "already_exists",
                }
            )
        else:
            logger.info("Creating Planner plan '%s' …", plan_cfg["title"])
            try:
                plan_data = _create_plan(client, group_id, plan_title)
                plan_id = plan_data.get("id", "dry-run-id")
                actions.append(
                    {
                        "action": "create_plan",
                        "plan": plan_cfg["title"],
                        "plan_id": plan_id,
                        "group_id": group_id,
                    }
                )
            except Exception as exc:
                logger.error("Failed to create plan '%s': %s", plan_cfg["title"], exc)
                actions.append(
                    {"action": "error", "plan": plan_cfg["title"], "error": str(exc)}
                )
                continue

        # Create buckets and tasks
        for bucket_cfg in plan_cfg.get("buckets", []):
            bucket_name = bucket_cfg["name"]
            logger.info("Creating bucket '%s' in plan '%s'", bucket_name, plan_cfg["title"])

            try:
                bucket_data = _create_bucket(client, plan_id, bucket_name)
                bucket_id = bucket_data.get("id", "dry-run-id")
                actions.append(
                    {
                        "action": "create_bucket",
                        "bucket": bucket_name,
                        "bucket_id": bucket_id,
                        "plan_id": plan_id,
                    }
                )
            except Exception as exc:
                logger.error("Failed to create bucket '%s': %s", bucket_name, exc)
                continue

            # Create tasks in this bucket
            for task_cfg in bucket_cfg.get("tasks", []):
                logger.info("Creating task '%s' …", task_cfg["title"])
                try:
                    task_data = _create_task(
                        client, plan_id, bucket_id, task_cfg, run_id, assignee_ids
                    )
                    actions.append(
                        {
                            "action": "create_task",
                            "task": task_cfg["title"],
                            "task_id": task_data.get("id", ""),
                            "bucket": bucket_name,
                            "plan_id": plan_id,
                        }
                    )
                except Exception as exc:
                    logger.error("Failed to create task '%s': %s", task_cfg["title"], exc)

    return actions
