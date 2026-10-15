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
from ..progress import spinner
from .. import runtime
from ..formatting import html_to_text, format_datetime
from ..references import _resolve_message_id, _resolve_cal_id, _resolve_task_id, _resolve_contact_id


def cmd_task_list(args: argparse.Namespace) -> None:
    """List incomplete tasks."""
    client = runtime.get_client()
    try:
        with spinner(args, "Loading tasks..."):
            tasks = client.list_tasks(top=args.count)
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to list tasks: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not tasks:
        if args.json:
            print(json.dumps([]))
            return
        runtime.console.print("[dim]No active tasks.[/]")
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

    runtime.console.print(table)
    runtime.console.print("[dim]Use 'outlook-cli task complete <n>' to mark a task as done.[/]")


def cmd_task_create(args: argparse.Namespace) -> None:
    """Create a new task."""
    client = runtime.get_client()
    try:
        with spinner(args, "Creating task..."):
            client.create_task(args.subject)
        runtime.console.print(f"[green]Created task:[/] {args.subject}")
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to create task: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_task_update(args: argparse.Namespace) -> None:
    """Update a task."""
    if not any([args.subject, args.due, args.importance, args.status]):
        runtime.console.print("[red]Provide at least one field to update.[/]")
        sys.exit(1)
    due_dt = None
    if args.due:
        try:
            due_dt = outlook_datetime(dateutil.parser.parse(args.due))
        except Exception as e:
            runtime.console.print(f"[red]Failed to parse due date: {e}[/]")
            sys.exit(1)
    client = runtime.get_client()
    task_id = _resolve_task_id(client, args.task_id)
    try:
        with spinner(args, "Updating task..."):
            client.update_task(
                task_id,
                subject=args.subject,
                due_dt=due_dt,
                importance=args.importance,
                status=args.status,
            )
        runtime.console.print("[green]Task updated.[/]")
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to update task: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_task_complete(args: argparse.Namespace) -> None:
    """Mark a task as complete."""
    client = runtime.get_client()
    task_id = _resolve_task_id(client, args.task_id)
    try:
        with spinner(args, "Completing task..."):
            client.complete_task(task_id)
        runtime.console.print("[green]Task marked as complete.[/]")
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to complete task: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_task_delete(args: argparse.Namespace) -> None:
    """Delete a task."""
    client = runtime.get_client()
    task_id = _resolve_task_id(client, args.task_id)
    try:
        with spinner(args, "Deleting task..."):
            client.delete_task(task_id)
        runtime.console.print("[green]Task deleted.[/]")
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to delete task: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
