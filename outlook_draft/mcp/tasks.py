"""Task tools."""

from __future__ import annotations

from dateutil.parser import parse as parse_datetime

from ..calendar_time import outlook_datetime
from .clients import clamp, outlook_client
from .shaping import JsonDict
from .validation import require_id


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
        tasks = client.list_tasks(top=clamp(count))
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


def task_update(task_id: str, subject: str | None = None, due: str | None = None,
                importance: str | None = None, status: str | None = None, confirm: bool = False) -> JsonDict:
    """Preview task edits. Omitted fields stay unchanged. Due dates use the configured local timezone. Set confirm=true only after approval. Never retry uncertain writes automatically."""
    require_id(task_id, "task_id")
    if importance is not None and importance not in {"Low", "Normal", "High"}:
        raise ValueError("importance must be Low, Normal or High")
    if status is not None and status not in {"NotStarted", "InProgress", "Completed", "WaitingOnOthers", "Deferred"}:
        raise ValueError("Invalid task status")
    fields = {key: value for key, value in dict(subject=subject, importance=importance, status=status).items() if value is not None}
    if due is not None:
        fields["due_dt"] = outlook_datetime(parse_datetime(due))
    if not fields:
        raise ValueError("Supply at least one task field to change.")
    with outlook_client() as client:
        result = {"id": task_id, "before": _task(client.get_task(task_id)), "fields": fields, "updated": False}
        if confirm is True:
            result.update(task=_task(client.update_task(task_id, **fields)), updated=True)
    return result


def task_delete(task_id: str, confirm: bool = False) -> JsonDict:
    """Preview task deletion. Set confirm=true only after approval of this exact task. Never retry uncertain writes automatically."""
    require_id(task_id, "task_id")
    with outlook_client() as client:
        result = {"id": task_id, "task": _task(client.get_task(task_id)), "deleted": False}
        if confirm is True:
            client.delete_task(task_id)
            result["deleted"] = True
    return result


TOOLS = [task_list, task_create, task_complete, task_update, task_delete]
