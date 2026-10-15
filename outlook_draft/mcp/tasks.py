"""Task tools."""

from __future__ import annotations

from .clients import clamp, outlook_client
from .shaping import JsonDict


def _task(task: JsonDict) -> JsonDict:
    due = task.get("DueDateTime")
    return {
        "id": task.get("Id", ""),
        "subject": task.get("Subject", ""),
        "status": task.get("Status", ""),
        "importance": task.get("Importance", ""),
        "due": due.get("DateTime", "") if isinstance(due, dict) else "",
    }


def task_list(count: int = 50) -> list[JsonDict]:
    """List incomplete tasks. Read-only."""
    with outlook_client() as client:
        tasks = client.list_tasks(top=clamp(count, high=200))
    return [_task(t) for t in tasks]


def task_create(subject: str) -> JsonDict:
    """Create a task in Outlook Tasks. Changes mailbox data."""
    if not subject.strip():
        raise ValueError("subject must not be empty")
    with outlook_client() as client:
        return _task(client.create_task(subject.strip()))


def task_complete(task_id: str) -> JsonDict:
    """Mark a task as completed. Changes mailbox data."""
    with outlook_client() as client:
        client.complete_task(task_id)
    return {"id": task_id, "status": "Completed"}


TOOLS = [task_list, task_create, task_complete]
