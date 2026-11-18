"""Microsoft Planner plans, buckets and tasks through Microsoft Graph.

The client mixin holds the HTTP calls. The service functions below it hold the
preview and confirm logic that the CLI and the MCP server share.
"""
from __future__ import annotations

import re
from datetime import UTC
from typing import Any
from urllib.parse import quote

from . import config
from .calendar_time import parse_local_datetime
from .errors import OutlookAPIError

JsonDict = dict[str, Any]

PRIORITIES = {"urgent": 1, "important": 3, "medium": 5, "low": 9}
ORDER_HINT = " !"
BETA_PLANNER_USER = "https://graph.microsoft.com/beta/me/planner"
_GUID = re.compile(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}")
_ID = re.compile(r"[A-Za-z0-9_-]{10,64}")


def planner_id(value: str, name: str) -> str:
    """Require a Planner ID (URL-safe characters). Cached CLI indexes are not accepted."""
    if not isinstance(value, str) or not _ID.fullmatch(value):
        raise ValueError(f"{name} must be a full Planner ID from the list commands.")
    return value


def priority_name(value: int | None) -> str:
    """Name the Planner priority band for a 0 to 10 value."""
    if value is None:
        return ""
    return "urgent" if value <= 1 else "important" if value <= 4 else "medium" if value <= 7 else "low"


def status_name(percent: int | None) -> str:
    """Name the Planner progress state for a percent value."""
    return "NotStarted" if not percent else "Completed" if percent >= 100 else "InProgress"


def task_summary(task: JsonDict) -> JsonDict:
    """Reduce a Graph plannerTask to the fields that triage needs."""
    return {
        "id": task.get("id", ""),
        "title": task.get("title", ""),
        "planId": task.get("planId", ""),
        "bucketId": task.get("bucketId", ""),
        "percentComplete": task.get("percentComplete", 0),
        "status": status_name(task.get("percentComplete")),
        "priority": priority_name(task.get("priority")),
        "due": task.get("dueDateTime") or "",
        "created": task.get("createdDateTime", ""),
        "completed": task.get("completedDateTime") or "",
        "assignees": sorted((task.get("assignments") or {}).keys()),
        "hasDescription": bool(task.get("hasDescription")),
    }


def _due_utc(value: str) -> str:
    return parse_local_datetime(value).astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def task_changes(*, title: str | None = None, bucket_id: str | None = None, due: str | None = None,
                 clear_due: bool = False, priority: str | None = None,
                 percent_complete: int | None = None) -> JsonDict:
    """Validate task fields and map them to Graph property names."""
    changes: JsonDict = {}
    if title is not None:
        if not title.strip():
            raise ValueError("title must not be empty")
        changes["title"] = title.strip()
    if bucket_id is not None:
        changes["bucketId"] = planner_id(bucket_id, "bucket_id")
    if due is not None and clear_due:
        raise ValueError("Use either due or clear_due, not both.")
    if due is not None:
        changes["dueDateTime"] = _due_utc(due)
    if clear_due:
        changes["dueDateTime"] = None
    if priority is not None:
        if priority not in PRIORITIES:
            raise ValueError("priority must be urgent, important, medium or low")
        changes["priority"] = PRIORITIES[priority]
    if percent_complete is not None:
        if type(percent_complete) is not int or not 0 <= percent_complete <= 100:
            raise ValueError("percent_complete must be a whole number from 0 to 100")
        changes["percentComplete"] = percent_complete
    return changes


def _valid_hint(hint: str | None) -> bool:
    """A Planner order hint holds a space or an exclamation mark."""
    return isinstance(hint, str) and (" " in hint or "!" in hint)


def _assignment(add: bool) -> JsonDict | None:
    if not add:
        return None
    return {"@odata.type": "#microsoft.graph.plannerAssignment", "orderHint": ORDER_HINT}


class PlannerMixin:
    """Graph calls for Planner. Use with a client that has the Graph base URL."""

    def _planner_pages(self, path: str, top: int, params: dict[str, str] | None = None) -> list[JsonDict]:
        items: list[JsonDict] = []
        next_path: str | None = path
        while next_path and len(items) < top:
            data = self._request("GET", next_path, params=params).json()
            items.extend(data.get("value", []))
            next_path, params = data.get("@odata.nextLink"), None
        return items[:top]

    def list_planner_plans(self, *, group_id: str | None = None, top: int = 50) -> list[JsonDict]:
        path = f"/groups/{quote(group_id, safe='')}/planner/plans" if group_id else "/me/planner/plans"
        return self._planner_pages(path, top)

    def list_planner_buckets(self, plan_id: str) -> list[JsonDict]:
        return self._planner_pages(f"/planner/plans/{quote(plan_id, safe='')}/buckets", 500)

    def list_planner_tasks(self, *, plan_id: str | None = None, bucket_id: str | None = None,
                           top: int = 50) -> list[JsonDict]:
        if bucket_id:
            path = f"/planner/buckets/{quote(bucket_id, safe='')}/tasks"
        elif plan_id:
            path = f"/planner/plans/{quote(plan_id, safe='')}/tasks"
        else:
            path = "/me/planner/tasks"
        return self._planner_pages(path, top)

    def list_planner_groups(self, *, search: str | None = None, top: int = 100) -> list[JsonDict]:
        """List the Microsoft 365 groups you belong to. Only these groups can own a plan."""
        params = {"$select": "id,displayName,description,visibility,groupTypes", "$top": "100"}
        groups = self._planner_pages("/me/memberOf/microsoft.graph.group", 5000, params)
        groups = [g for g in groups if "Unified" in (g.get("groupTypes") or [])]
        if search:
            needle = search.lower()
            groups = [g for g in groups if needle in (g.get("displayName") or "").lower()]
        return sorted(groups, key=lambda g: (g.get("displayName") or "").lower())[:top]

    def get_planner_plan(self, plan_id: str) -> JsonDict:
        return self._request("GET", f"/planner/plans/{quote(plan_id, safe='')}").json()

    def get_planner_user(self) -> JsonDict:
        """Read your Planner preferences, including pinned plans. Only the beta API exposes them."""
        return self._request("GET", BETA_PLANNER_USER).json()

    def pin_planner_plan(self, plan_id: str, title: str, user: JsonDict) -> JsonDict:
        """Pin a plan, so it shows under My plans.

        Existing pins are sent back so none is lost. The service can return a pin with an order hint such as
        "000" that it then rejects as invalid, so such a hint is replaced. The service also applies pins with
        a short delay, so a read straight after a write can still show the old list.
        """
        pins = {key: {**ref, "orderHint": ref["orderHint"] if _valid_hint(ref.get("orderHint")) else ORDER_HINT}
                for key, ref in (user.get("favoritePlanReferences") or {}).items()}
        pins[plan_id] = {"@odata.type": "microsoft.graph.plannerFavoritePlanReference",
                         "orderHint": ORDER_HINT, "planTitle": title}
        return self._request("PATCH", BETA_PLANNER_USER, json_body={"favoritePlanReferences": pins},
                             extra_headers={"If-Match": user["@odata.etag"], "Prefer": "return=representation"},
                             max_retries=0).json()

    def list_planner_members(self, group_id: str) -> list[JsonDict]:
        """List the members of the Microsoft 365 group that owns a plan."""
        params = {"$select": "id,displayName,mail,userPrincipalName", "$top": "100"}
        return self._planner_pages(f"/groups/{quote(group_id, safe='')}/members", 5000, params)

    def add_planner_member(self, group_id: str, user_id: str) -> None:
        """Add a user to the group. The service sends the person an invitation email. Never retried."""
        body = {"@odata.id": f"{config.GRAPH_BASE_URL}/directoryObjects/{user_id}"}
        self._request("POST", f"/groups/{quote(group_id, safe='')}/members/$ref", json_body=body, max_retries=0)

    def get_planner_task(self, task_id: str) -> JsonDict:
        return self._request("GET", f"/planner/tasks/{quote(task_id, safe='')}").json()

    def get_planner_task_details(self, task_id: str) -> JsonDict:
        return self._request("GET", f"/planner/tasks/{quote(task_id, safe='')}/details").json()

    def resolve_planner_user(self, reference: str) -> JsonDict:
        """Resolve an email address or a directory user ID to an ID and a display name."""
        ref = reference.strip()
        if _GUID.fullmatch(ref):
            return {"id": ref.lower(), "displayName": ""}
        if not re.fullmatch(r"[^@\s<>/]+@[^@\s<>/]+", ref):
            raise ValueError("Assignees must be an email address or a directory user ID.")
        user = self._request("GET", f"/users/{quote(ref, safe='@')}",
                             params={"$select": "id,displayName,mail"}).json()
        return {"id": user["id"], "displayName": user.get("displayName", "")}

    def get_planner_bucket(self, bucket_id: str) -> JsonDict:
        return self._request("GET", f"/planner/buckets/{quote(bucket_id, safe='')}").json()

    def update_planner_bucket(self, bucket_id: str, changes: JsonDict) -> JsonDict:
        """Patch a bucket with its current etag, so a concurrent edit fails instead of being overwritten."""
        path = f"/planner/buckets/{quote(bucket_id, safe='')}"
        etag = self.get_planner_bucket(bucket_id)["@odata.etag"]
        self._request("PATCH", path, json_body=changes, max_retries=0,
                      extra_headers={"If-Match": etag, "Prefer": "return=minimal"})
        return self.get_planner_bucket(bucket_id)

    def delete_planner_bucket(self, bucket_id: str) -> None:
        path = f"/planner/buckets/{quote(bucket_id, safe='')}"
        etag = self.get_planner_bucket(bucket_id)["@odata.etag"]
        self._request("DELETE", path, extra_headers={"If-Match": etag}, max_retries=0)

    def create_planner_bucket(self, plan_id: str, name: str) -> JsonDict:
        body = {"name": name, "planId": plan_id, "orderHint": ORDER_HINT}
        return self._request("POST", "/planner/buckets", json_body=body, max_retries=0).json()

    def create_planner_task(self, body: JsonDict, description: str | None = None) -> JsonDict:
        """Create a task. Writes never retry automatically. The description is a second call."""
        task = self._request("POST", "/planner/tasks", json_body=body, max_retries=0).json()
        if description:
            try:
                self.set_planner_description(task["id"], description)
            except OutlookAPIError as exc:
                raise OutlookAPIError(
                    exc.status, f"Task {task['id']} was created, but the description failed: {exc}",
                ) from exc
        return task

    def set_planner_description(self, task_id: str, description: str) -> None:
        path = f"/planner/tasks/{quote(task_id, safe='')}/details"
        etag = self._request("GET", path).json()["@odata.etag"]
        self._request("PATCH", path, json_body={"description": description},
                      extra_headers={"If-Match": etag, "Prefer": "return=minimal"}, max_retries=0)

    def update_planner_task(self, task_id: str, changes: JsonDict, description: str | None = None) -> JsonDict:
        """Patch a task with its current etag, so a concurrent edit fails instead of being overwritten."""
        path = f"/planner/tasks/{quote(task_id, safe='')}"
        task = self.get_planner_task(task_id)
        if changes:
            self._request("PATCH", path, json_body=changes, max_retries=0,
                          extra_headers={"If-Match": task["@odata.etag"], "Prefer": "return=minimal"})
        if description is not None:
            self.set_planner_description(task_id, description)
        return self.get_planner_task(task_id)

    def delete_planner_task(self, task_id: str) -> None:
        path = f"/planner/tasks/{quote(task_id, safe='')}"
        etag = self.get_planner_task(task_id)["@odata.etag"]
        self._request("DELETE", path, extra_headers={"If-Match": etag}, max_retries=0)


def _resolve_assignees(client: Any, assign: list[str] | None, unassign: list[str] | None) -> tuple[JsonDict, list[JsonDict]]:
    people = [client.resolve_planner_user(ref) for ref in [*(assign or []), *(unassign or [])]]
    added = people[: len(assign or [])]
    assignments = {person["id"]: _assignment(True) for person in added}
    assignments.update({person["id"]: _assignment(False) for person in people[len(added):]})
    return assignments, people


def create_task(client: Any, *, plan_id: str, bucket_id: str, title: str, assign: list[str] | None = None,
                due: str | None = None, priority: str | None = None, description: str | None = None,
                confirm: bool = False) -> JsonDict:
    """Preview a new task, or create it when confirm is true."""
    body = {"planId": planner_id(plan_id, "plan_id"),
            **task_changes(title=title, bucket_id=bucket_id, due=due, priority=priority)}
    assignments, people = _resolve_assignees(client, assign, None)
    if assignments:
        body["assignments"] = assignments
    result: JsonDict = {"task": body, "assignees": people, "description": description, "created": False}
    if confirm is True:
        created = client.create_planner_task(body, description)
        result.update(task=task_summary(created), created=True)
    return result


def update_task(client: Any, task_id: str, *, assign: list[str] | None = None, unassign: list[str] | None = None,
                description: str | None = None, confirm: bool = False, **fields: Any) -> JsonDict:
    """Preview task edits, or apply them when confirm is true. Omitted fields stay unchanged."""
    planner_id(task_id, "task_id")
    changes = task_changes(**fields)
    assignments, people = _resolve_assignees(client, assign, unassign)
    if assignments:
        changes["assignments"] = assignments
    if not changes and description is None:
        raise ValueError("Supply at least one task field to change.")
    result: JsonDict = {"id": task_id, "before": task_summary(client.get_planner_task(task_id)),
                        "changes": changes, "description": description, "assignees": people, "updated": False}
    if confirm is True:
        result.update(task=task_summary(client.update_planner_task(task_id, changes, description)), updated=True)
    return result


def delete_task(client: Any, task_id: str, *, confirm: bool = False) -> JsonDict:
    """Preview the exact task, or delete it when confirm is true."""
    planner_id(task_id, "task_id")
    result: JsonDict = {"id": task_id, "task": task_summary(client.get_planner_task(task_id)), "deleted": False}
    if confirm is True:
        client.delete_planner_task(task_id)
        result["deleted"] = True
    return result


def board_order(buckets: list[JsonDict]) -> list[JsonDict]:
    """Order buckets left to right as the Planner board shows them.

    The service lists buckets with the smallest order hint first, but the board shows the largest hint
    first. A new bucket gets the smallest hint, so it lands at the end of the board.
    """
    return sorted(buckets, key=lambda b: b.get("orderHint") or "", reverse=True)


def _move_hint(others: list[JsonDict], after: str | None) -> str:
    """Build the order hint that places a bucket first on the board, or right after the bucket `after`.

    `others` are the other buckets in board order. Planner reads " <smaller hint> <larger hint>!" as
    "between these two" and returns the real hint. First on the board means a hint larger than all.
    """
    if after is None:
        return f"{others[0]['orderHint']} !" if others else " !"
    index = next((i for i, b in enumerate(others) if b["id"] == after), None)
    if index is None:
        raise ValueError("after_bucket_id must be another bucket of the same plan.")
    smaller = others[index + 1]["orderHint"] if index + 1 < len(others) else ""
    return f"{smaller} {others[index]['orderHint']}!"


def update_bucket(client: Any, bucket_id: str, *, name: str | None = None, after: str | None = None,
                  first: bool = False, confirm: bool = False) -> JsonDict:
    """Preview renaming or moving a bucket, or apply it when confirm is true."""
    planner_id(bucket_id, "bucket_id")
    if after is not None:
        planner_id(after, "after_bucket_id")
        if after == bucket_id:
            raise ValueError("A bucket cannot be placed after itself.")
    if after is not None and first:
        raise ValueError("Use either after_bucket_id or first, not both.")
    if name is not None and not name.strip():
        raise ValueError("name must not be empty")
    if name is None and after is None and not first:
        raise ValueError("Supply a new name, after_bucket_id or first.")
    bucket = client.get_planner_bucket(bucket_id)
    siblings = board_order(client.list_planner_buckets(bucket["planId"]))
    others = [b for b in siblings if b["id"] != bucket_id]
    changes: JsonDict = {}
    if name is not None:
        changes["name"] = name.strip()
    if after is not None or first:
        changes["orderHint"] = _move_hint(others, after)
    names = {b["id"]: b.get("name", "") for b in siblings}
    names[bucket_id] = changes.get("name", names.get(bucket_id, ""))
    order = [b["id"] for b in others]
    if "orderHint" in changes:
        order.insert(0 if first else order.index(after) + 1, bucket_id)
    else:
        order = [b["id"] for b in siblings]
    result: JsonDict = {"id": bucket_id, "before": {"name": bucket.get("name", "")}, "changes": changes,
                        "orderAfter": [names[i] for i in order], "updated": False}
    if confirm is True:
        updated = client.update_planner_bucket(bucket_id, changes)
        result.update(bucket={"id": bucket_id, "name": updated.get("name", "")}, updated=True)
    return result


def delete_bucket(client: Any, bucket_id: str, *, confirm: bool = False) -> JsonDict:
    """Preview deleting an empty bucket, or delete it when confirm is true.

    The Graph API does not delete the tasks of a deleted bucket. It leaves them attached to a bucket that
    no longer exists, and Planner hides them. So a bucket that still holds tasks is never deleted.
    """
    planner_id(bucket_id, "bucket_id")
    bucket = client.get_planner_bucket(bucket_id)
    tasks = client.list_planner_tasks(bucket_id=bucket_id, top=1000)
    result: JsonDict = {"bucket": {"id": bucket_id, "name": bucket.get("name", ""), "planId": bucket.get("planId", "")},
                        "tasksInBucket": len(tasks), "taskTitles": [t.get("title", "") for t in tasks[:10]],
                        "blocked": bool(tasks), "deleted": False}
    if tasks and confirm is True:
        raise ValueError(f"The bucket still holds {len(tasks)} task(s). Move or delete them first, "
                         "because the API would leave them orphaned and hidden.")
    if confirm is True:
        client.delete_planner_bucket(bucket_id)
        result["deleted"] = True
    return result


def pin_plan(client: Any, plan_id: str, *, confirm: bool = False) -> JsonDict:
    """Preview pinning a plan to My plans, or pin it when confirm is true. It changes only your own list."""
    planner_id(plan_id, "plan_id")
    plan = client.get_planner_plan(plan_id)
    user = client.get_planner_user()
    already = plan_id in (user.get("favoritePlanReferences") or {})
    result: JsonDict = {"plan": {"id": plan_id, "title": plan.get("title", "")}, "alreadyPinned": already,
                        "pinned": already}
    if confirm is True and not already:
        client.pin_planner_plan(plan_id, plan.get("title", ""), user)
        result["pinned"] = True
    return result


def _person(user: JsonDict) -> JsonDict:
    return {"id": user.get("id", ""), "displayName": user.get("displayName") or "",
            "mail": user.get("mail") or user.get("userPrincipalName") or ""}


def _plan_group(client: Any, plan_id: str) -> tuple[JsonDict, str]:
    plan = client.get_planner_plan(plan_id)
    group_id = plan.get("owner") or ""
    if not group_id:
        raise ValueError("The plan has no owning group, so it has no member list to change.")
    return {"id": plan_id, "title": plan.get("title", "")}, group_id


def list_members(client: Any, plan_id: str) -> JsonDict:
    """List who can work on a plan: the members of its owning Microsoft 365 group."""
    planner_id(plan_id, "plan_id")
    plan, group_id = _plan_group(client, plan_id)
    return {"plan": plan, "groupId": group_id, "members": [_person(m) for m in client.list_planner_members(group_id)]}


def add_members(client: Any, plan_id: str, people: list[str], *, confirm: bool = False) -> JsonDict:
    """Preview adding people to a plan's group, or add them when confirm is true.

    Planner has no member list of its own. Access comes from the owning group, so this adds each person to
    the group. That also gives them the group mailbox and SharePoint site, and the service emails them.
    """
    planner_id(plan_id, "plan_id")
    if not people:
        raise ValueError("Supply at least one person (email address or directory user ID).")
    plan, group_id = _plan_group(client, plan_id)
    current = {m["id"].lower() for m in client.list_planner_members(group_id)}
    resolved = [client.resolve_planner_user(ref) for ref in people]
    unique = {p["id"]: p for p in resolved}.values()
    to_add = [p for p in unique if p["id"].lower() not in current]
    result: JsonDict = {"plan": plan, "groupId": group_id, "toAdd": to_add,
                        "alreadyMembers": [p for p in unique if p["id"].lower() in current], "added": [], "applied": False}
    if confirm is True:
        for person in to_add:
            try:
                client.add_planner_member(group_id, person["id"])
            except OutlookAPIError as exc:
                done = ", ".join(p["displayName"] or p["id"] for p in result["added"]) or "nobody"
                raise OutlookAPIError(exc.status, f"Added {done} before failing on {person['displayName'] or person['id']}: {exc}") from exc
            result["added"].append(person)
        result["applied"] = True
    return result


def create_bucket(client: Any, plan_id: str, name: str, *, confirm: bool = False) -> JsonDict:
    """Preview a new bucket, or create it when confirm is true."""
    planner_id(plan_id, "plan_id")
    if not name.strip():
        raise ValueError("name must not be empty")
    result: JsonDict = {"planId": plan_id, "name": name.strip(), "created": False}
    if confirm is True:
        bucket = client.create_planner_bucket(plan_id, name.strip())
        result.update(id=bucket.get("id", ""), created=True)
    return result
