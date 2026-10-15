from __future__ import annotations

import logging

from ..errors import CliError

from ..parser import add_output_args, _add_yes_arg


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


log = logging.getLogger(__name__)


def cmd_task_list(args: argparse.Namespace) -> None:
    """List incomplete tasks."""
    client = runtime.get_client()
    try:
        with spinner(args, "Loading tasks..."):
            tasks = client.list_tasks(top=args.count)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to list tasks: {e}[/]")
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
                log.debug("Best-effort operation failed", exc_info=True)
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
        raise CliError(f"[red]Failed to create task: {e}[/]")
    finally:
        client.close()


def cmd_task_update(args: argparse.Namespace) -> None:
    """Update a task."""
    if not any([args.subject, args.due, args.importance, args.status]):
        raise CliError("[red]Provide at least one field to update.[/]")
    due_dt = None
    if args.due:
        try:
            due_dt = outlook_datetime(dateutil.parser.parse(args.due))
        except Exception as e:
            log.debug("Best-effort operation failed", exc_info=True)
            raise CliError(f"[red]Failed to parse due date: {e}[/]")
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
        raise CliError(f"[red]Failed to update task: {e}[/]")
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
        raise CliError(f"[red]Failed to complete task: {e}[/]")
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
        raise CliError(f"[red]Failed to delete task: {e}[/]")
    finally:
        client.close()


def register(sub) -> None:
    p_task = sub.add_parser("task", help="Manage tasks")
    add_output_args(p_task)
    task_sub = p_task.add_subparsers(dest="command", required=True)

    cmd_task_l = task_sub.add_parser("list", help="List active tasks")
    add_output_args(cmd_task_l)
    cmd_task_l.add_argument("--count", "-n", type=int, default=50, help="Max tasks")
    cmd_task_l.set_defaults(func=cmd_task_list)

    cmd_task_c = task_sub.add_parser("create", help="Create a task")
    cmd_task_c.add_argument("subject", help="Task description")
    cmd_task_c.set_defaults(func=cmd_task_create)

    cmd_task_update_cmd = task_sub.add_parser("update", help="Update a task")
    cmd_task_update_cmd.add_argument("task_id", help="Task index or full ID")
    cmd_task_update_cmd.add_argument("--subject", help="New task subject")
    cmd_task_update_cmd.add_argument("--due", help="Due date/time")
    cmd_task_update_cmd.add_argument("--importance", choices=["Low", "Normal", "High"], help="Task importance")
    cmd_task_update_cmd.add_argument("--status", choices=["NotStarted", "InProgress", "Completed", "WaitingOnOthers", "Deferred"], help="Task status")
    cmd_task_update_cmd.set_defaults(func=cmd_task_update)

    cmd_task_complete_cmd = task_sub.add_parser("complete", help="Mark task complete")
    cmd_task_complete_cmd.add_argument("task_id", help="Task index or full ID")
    cmd_task_complete_cmd.set_defaults(func=cmd_task_complete)

    cmd_task_delete_cmd = task_sub.add_parser("delete", help="Delete a task")
    cmd_task_delete_cmd.add_argument("task_id", help="Task index or full ID")
    cmd_task_delete_cmd.set_defaults(func=cmd_task_delete)
