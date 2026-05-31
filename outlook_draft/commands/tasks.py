from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime

import dateutil.parser
from typing import Any, Callable

from rich.console import Console
from rich.table import Table

from ..calendar_time import outlook_datetime
from ..cache import TASK_CACHE, save_cache
from ..errors import OutlookAPIError


def build_ctx(
    *,
    console: Console,
    get_client: Callable[[], Any],
    resolve_task_id: Callable[[Any, str], str],
) -> dict[str, Any]:
    return {
        "console": console,
        "get_client": get_client,
        "resolve_task_id": resolve_task_id,
    }


def _ctx(args: argparse.Namespace) -> dict[str, Any]:
    return args._tasks_ctx


def _console(args: argparse.Namespace) -> Console:
    return _ctx(args)["console"]


def _get_client(args: argparse.Namespace):
    return _ctx(args)["get_client"]()


def _resolve_task_id(args: argparse.Namespace, client: Any, task_ref: str) -> str:
    return _ctx(args)["resolve_task_id"](client, task_ref)


def cmd_task_list(args: argparse.Namespace) -> None:
    """List incomplete tasks."""
    client = _get_client(args)
    try:
        tasks = client.list_tasks(top=args.count)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to list tasks: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not tasks:
        if args.json:
            print(json.dumps([]))
            return
        _console(args).print("[dim]No active tasks.[/]")
        return

    save_cache(TASK_CACHE, tasks)

    if args.json:
        print(json.dumps([
            {
                "index": i,
                "id": task.get("Id", ""),
                "subject": task.get("Subject", ""),
                "status": task.get("Status", ""),
                "importance": task.get("Importance", ""),
                "dueDateTime": task.get("DueDateTime", {}),
            }
            for i, task in enumerate(tasks, 1)
        ], indent=2, ensure_ascii=False))
        return

    table = Table(title=f"Active Tasks ({len(tasks)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Task", ratio=1)
    table.add_column("Due", width=15)
    table.add_column("Status", width=12)

    for i, t in enumerate(tasks, 1):
        due_dt = t.get("DueDateTime", {})
        due = due_dt.get("DateTime", "") if isinstance(due_dt, dict) else ""
        if due:
            try:
                parsed = datetime.fromisoformat(due[:23])
                due = parsed.strftime("%b %d, %Y")
            except Exception:
                due = due[:10]

        subj = t.get("Subject", "(no subject)")
        status = t.get("Status", "")
        table.add_row(str(i), subj, due, status)

    _console(args).print(table)
    _console(args).print("[dim]Use 'outlook-cli task complete <n>' to mark a task as done.[/]")


def cmd_task_create(args: argparse.Namespace) -> None:
    """Create a new task."""
    client = _get_client(args)
    try:
        client.create_task(args.subject)
        _console(args).print(f"[green]Created task:[/] {args.subject}")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to create task: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_task_update(args: argparse.Namespace) -> None:
    """Update a task."""
    if not any([args.subject, args.due, args.importance, args.status]):
        _console(args).print("[red]Provide at least one field to update.[/]")
        sys.exit(1)
    due_dt = None
    if args.due:
        try:
            due_dt = outlook_datetime(dateutil.parser.parse(args.due))
        except Exception as e:
            _console(args).print(f"[red]Failed to parse due date: {e}[/]")
            sys.exit(1)
    client = _get_client(args)
    task_id = _resolve_task_id(args, client, args.task_id)
    try:
        client.update_task(
            task_id,
            subject=args.subject,
            due_dt=due_dt,
            importance=args.importance,
            status=args.status,
        )
        _console(args).print("[green]Task updated.[/]")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to update task: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_task_complete(args: argparse.Namespace) -> None:
    """Mark a task as complete."""
    client = _get_client(args)
    task_id = _resolve_task_id(args, client, args.task_id)
    try:
        client.complete_task(task_id)
        _console(args).print("[green]Task marked as complete.[/]")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to complete task: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_task_delete(args: argparse.Namespace) -> None:
    """Delete a task."""
    client = _get_client(args)
    task_id = _resolve_task_id(args, client, args.task_id)
    try:
        client.delete_task(task_id)
        _console(args).print("[green]Task deleted.[/]")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to delete task: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
