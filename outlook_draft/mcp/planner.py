"""Microsoft Planner tools. Writes preview unless confirm=true."""

from __future__ import annotations

from .. import planner
from .clients import clamp, graph_client
from .shaping import JsonDict


def _plan(plan: JsonDict) -> JsonDict:
    return {"id": plan.get("id", ""), "title": plan.get("title", ""), "owner": plan.get("owner", ""),
            "created": plan.get("createdDateTime", "")}


def planner_plans(group_id: str | None = None, count: int = 50) -> list[JsonDict]:
    """List Planner plans you can see. Without group_id it lists only plans you recently used or pinned, so pass group_id to list a specific group. Read-only."""
    with graph_client() as client:
        plans = client.list_planner_plans(group_id=group_id, top=clamp(count))
    return [_plan(p) for p in plans]


def planner_groups(search: str | None = None, count: int = 100) -> list[JsonDict]:
    """List the Microsoft 365 groups you belong to, optionally filtered by a name search. Only these groups can own a plan, so use the group ID with planner_plans. Read-only."""
    with graph_client() as client:
        groups = client.list_planner_groups(search=search, top=clamp(count))
    return [{"id": g.get("id", ""), "name": g.get("displayName", ""), "visibility": g.get("visibility") or "",
             "description": g.get("description") or ""} for g in groups]


def planner_buckets(plan_id: str) -> list[JsonDict]:
    """List the buckets (columns) of a plan, left to right as the board shows them. Read-only."""
    planner.planner_id(plan_id, "plan_id")
    with graph_client() as client:
        buckets = client.list_planner_buckets(plan_id)
    return [{"id": b.get("id", ""), "name": b.get("name", ""), "planId": b.get("planId", "")}
            for b in planner.board_order(buckets)]


def planner_tasks(plan_id: str | None = None, bucket_id: str | None = None, count: int = 50) -> list[JsonDict]:
    """List Planner tasks of a plan or a bucket. With neither, list tasks assigned to you. Read-only."""
    for value, name in ((plan_id, "plan_id"), (bucket_id, "bucket_id")):
        if value is not None:
            planner.planner_id(value, name)
    with graph_client() as client:
        tasks = client.list_planner_tasks(plan_id=plan_id, bucket_id=bucket_id, top=clamp(count))
    return [planner.task_summary(t) for t in tasks]


def planner_task_get(task_id: str) -> JsonDict:
    """Get one Planner task with its description and checklist. Read-only."""
    planner.planner_id(task_id, "task_id")
    with graph_client() as client:
        task = client.get_planner_task(task_id)
        details = client.get_planner_task_details(task_id)
    checklist = [{"title": i.get("title", ""), "checked": bool(i.get("isChecked"))}
                 for i in (details.get("checklist") or {}).values()]
    return {**planner.task_summary(task), "description": details.get("description", ""), "checklist": checklist}


def planner_task_create(plan_id: str, bucket_id: str, title: str, assign: list[str] | None = None,
                        due: str | None = None, priority: str | None = None, description: str | None = None,
                        confirm: bool = False) -> JsonDict:
    """Preview a new Planner task. assign takes emails or user IDs and notifies them. priority is urgent, important, medium or low. Due dates use the configured local timezone. Set confirm=true only after approval. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.create_task(client, plan_id=plan_id, bucket_id=bucket_id, title=title, assign=assign,
                                   due=due, priority=priority, description=description, confirm=confirm)


def planner_task_update(task_id: str, title: str | None = None, bucket_id: str | None = None,
                        assign: list[str] | None = None, unassign: list[str] | None = None,
                        due: str | None = None, clear_due: bool = False, priority: str | None = None,
                        percent_complete: int | None = None, description: str | None = None,
                        confirm: bool = False) -> JsonDict:
    """Preview Planner task edits. Omitted fields stay unchanged. Move a task with bucket_id. percent_complete 100 completes it. A concurrent edit makes the write fail, not overwrite. Set confirm=true only after approval. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.update_task(client, task_id, assign=assign, unassign=unassign, description=description,
                                   confirm=confirm, title=title, bucket_id=bucket_id, due=due,
                                   clear_due=clear_due, priority=priority, percent_complete=percent_complete)


def planner_task_delete(task_id: str, confirm: bool = False) -> JsonDict:
    """Preview deletion of a Planner task. Set confirm=true only after approval of this exact task. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.delete_task(client, task_id, confirm=confirm)


def planner_bucket_update(bucket_id: str, name: str | None = None, after_bucket_id: str | None = None,
                          first: bool = False, confirm: bool = False) -> JsonDict:
    """Preview renaming a bucket or moving it within its plan. Use after_bucket_id to place it right after (to the right of) another bucket of the same plan, or first=true to place it first (leftmost) on the board. The preview lists the bucket order after the change. Set confirm=true only after approval. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.update_bucket(client, bucket_id, name=name, after=after_bucket_id, first=first, confirm=confirm)


def planner_bucket_delete(bucket_id: str, confirm: bool = False) -> JsonDict:
    """Preview deleting an empty bucket. A bucket that still holds tasks is refused, because the API would leave those tasks orphaned and hidden: move or delete them first (the preview shows their count and titles). Set confirm=true only after approval of this exact bucket. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.delete_bucket(client, bucket_id, confirm=confirm)


def planner_members(plan_id: str) -> JsonDict:
    """List who can work on a plan: the members of its owning Microsoft 365 group. Read-only."""
    with graph_client() as client:
        return planner.list_members(client, plan_id)


def planner_member_add(plan_id: str, people: list[str], confirm: bool = False) -> JsonDict:
    """Preview adding people (emails or user IDs) to a plan so they can work on it. Planner access comes from the owning Microsoft 365 group, so this adds each person to the group: they also get the group mailbox and SharePoint site, and the service emails them an invitation. You must own the group. Set confirm=true only after approval of the exact list of people. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.add_members(client, plan_id, people, confirm=confirm)


def planner_member_remove(plan_id: str, people: list[str], confirm: bool = False) -> JsonDict:
    """Preview removing people (emails or user IDs) from a plan's Microsoft 365 group. They lose access to the plan, the group mailbox and the SharePoint site. Tasks assigned to them stay assigned: the preview counts those tasks. Group owners and the last member are refused. You must own the group. Set confirm=true only after approval of the exact list of people. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.remove_members(client, plan_id, people, confirm=confirm)


def planner_plan_unpin(plan_id: str, confirm: bool = False) -> JsonDict:
    """Preview removing a plan from your My plans list. It changes only your own Planner list, not the plan. Set confirm=true only after approval. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.unpin_plan(client, plan_id, confirm=confirm)


def planner_plan_pin(plan_id: str, confirm: bool = False) -> JsonDict:
    """Preview pinning a plan to your My plans list. It changes only your own Planner list, not the plan. Set confirm=true only after approval. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.pin_plan(client, plan_id, confirm=confirm)


def planner_bucket_create(plan_id: str, name: str, confirm: bool = False) -> JsonDict:
    """Preview a new bucket in a plan. It is added at the end (right) of the board. Set confirm=true only after approval. Never retry uncertain writes automatically."""
    with graph_client() as client:
        return planner.create_bucket(client, plan_id, name, confirm=confirm)


TOOLS = [planner_plans, planner_groups, planner_buckets, planner_tasks, planner_task_get, planner_task_create,
         planner_task_update, planner_task_delete, planner_bucket_create, planner_bucket_update,
         planner_bucket_delete, planner_plan_pin, planner_members, planner_member_add,
         planner_member_remove, planner_plan_unpin]
