from __future__ import annotations

import argparse
import json

from rich.table import Table

from .. import planner, runtime
from ..errors import CliError, OutlookAPIError, OutlookDraftError
from ..parser import add_output_args
from ..progress import spinner


def _run(args: argparse.Namespace, label: str, action):
    """Run one Graph action with a Graph client. Convert API and validation errors to CLI errors."""
    client = runtime.get_graph_client()
    try:
        with spinner(args, label):
            return action(client)
    except (OutlookAPIError, OutlookDraftError, ValueError) as exc:
        raise CliError(f"[red]{label.rstrip('.')} failed: {exc}[/]") from exc
    finally:
        client.close()


def _table(title: str, columns: list[str], rows: list[list[str]]) -> None:
    table = Table(title=title)
    for column in columns:
        table.add_column(column, overflow="fold")
    for row in rows:
        table.add_row(*row)
    runtime.console.print(table)


def cmd_planner_plans(args: argparse.Namespace) -> None:
    plans = _run(args, "Loading plans...", lambda c: c.list_planner_plans(group_id=args.group, top=args.count))
    if args.json:
        return runtime.print_json(plans)
    _table(f"Planner plans ({len(plans)})", ["Title", "ID"], [[p.get("title", ""), p.get("id", "")] for p in plans])


def cmd_planner_groups(args: argparse.Namespace) -> None:
    groups = _run(args, "Loading groups...", lambda c: c.list_planner_groups(search=args.search, top=args.count))
    if args.json:
        return runtime.print_json(groups)
    _table(f"Groups ({len(groups)})", ["Name", "ID"], [[g.get("displayName", ""), g.get("id", "")] for g in groups])


def cmd_planner_members(args: argparse.Namespace) -> None:
    result = _run(args, "Loading members...", lambda c: planner.list_members(c, args.plan_id))
    if args.json:
        return runtime.print_json(result)
    _table(f"Members of {result['plan']['title']} ({len(result['members'])})", ["Name", "Email"],
           [[m["displayName"], m["mail"]] for m in result["members"]])


def cmd_planner_member_add(args: argparse.Namespace) -> None:
    result = _run(args, "Adding members...", lambda c: planner.add_members(
        c, args.plan_id, args.people, confirm=args.yes))
    if args.json:
        return runtime.print_json(result)
    label = lambda people: ", ".join(p["displayName"] or p["id"] for p in people) or "nobody"
    runtime.console.print(f"Already members: {label(result['alreadyMembers'])}")
    if result["applied"]:
        runtime.console.print(f"[green]Added: {label(result['added'])}[/]")
    else:
        runtime.console.print(f"[yellow]Preview only. Nothing changed. Would add: {label(result['toAdd'])}. "
                              "The service emails each new member and gives access to the group mailbox and "
                              "SharePoint site. Add --yes to apply.[/]")


def cmd_planner_pin(args: argparse.Namespace) -> None:
    result = _run(args, "Pinning plan...", lambda c: planner.pin_plan(c, args.plan_id, confirm=args.yes))
    if args.json:
        return runtime.print_json(result)
    title = result["plan"]["title"]
    if result["alreadyPinned"]:
        runtime.console.print(f"[green]{title} is already pinned.[/]")
    elif result["pinned"]:
        runtime.console.print(f"[green]Pinned {title} to My plans.[/]")
    else:
        runtime.console.print(f"[yellow]Preview only. Nothing changed. Add --yes to pin {title}.[/]")


def cmd_planner_buckets(args: argparse.Namespace) -> None:
    buckets = planner.board_order(_run(args, "Loading buckets...", lambda c: c.list_planner_buckets(
        planner.planner_id(args.plan_id, "plan_id"))))
    if args.json:
        return runtime.print_json(buckets)
    _table(f"Buckets ({len(buckets)})", ["Name", "ID"], [[b.get("name", ""), b.get("id", "")] for b in buckets])


def cmd_planner_tasks(args: argparse.Namespace) -> None:
    tasks = _run(args, "Loading tasks...", lambda c: c.list_planner_tasks(
        plan_id=args.plan, bucket_id=args.bucket, top=args.count))
    summaries = [planner.task_summary(t) for t in tasks]
    if args.json:
        return runtime.print_json(summaries)
    _table(f"Planner tasks ({len(summaries)})", ["Title", "Status", "Priority", "Due", "ID"],
           [[t["title"], t["status"], t["priority"], t["due"][:10], t["id"]] for t in summaries])


def cmd_planner_show(args: argparse.Namespace) -> None:
    def load(client):
        planner.planner_id(args.task_id, "task_id")
        return client.get_planner_task(args.task_id), client.get_planner_task_details(args.task_id)

    task, details = _run(args, "Loading task...", load)
    summary = {**planner.task_summary(task), "description": details.get("description", ""),
               "checklist": [{"title": i.get("title", ""), "checked": bool(i.get("isChecked"))}
                             for i in (details.get("checklist") or {}).values()]}
    if args.json:
        return runtime.print_json(summary)
    for key, value in summary.items():
        runtime.console.print(f"[bold]{key}:[/] {value}")


def _report(args: argparse.Namespace, result: dict, done_key: str, done_text: str) -> None:
    """Print a write result. Without --yes the result is a preview."""
    if args.json:
        return runtime.print_json(result)
    if result.get(done_key):
        runtime.console.print(f"[green]{done_text}[/]")
    else:
        runtime.console.print("[yellow]Preview only. Nothing changed. Add --yes to apply.[/]")
    runtime.console.print(json.dumps(result, indent=2, ensure_ascii=False))


def cmd_planner_create(args: argparse.Namespace) -> None:
    result = _run(args, "Creating task...", lambda c: planner.create_task(
        c, plan_id=args.plan_id, bucket_id=args.bucket_id, title=args.title, assign=args.assign,
        due=args.due, priority=args.priority, description=args.description, confirm=args.yes))
    _report(args, result, "created", "Task created.")


def cmd_planner_update(args: argparse.Namespace) -> None:
    result = _run(args, "Updating task...", lambda c: planner.update_task(
        c, args.task_id, assign=args.assign, unassign=args.unassign, description=args.description,
        confirm=args.yes, title=args.title, bucket_id=args.bucket, due=args.due, clear_due=args.clear_due,
        priority=args.priority, percent_complete=args.percent))
    _report(args, result, "updated", "Task updated.")


def cmd_planner_delete(args: argparse.Namespace) -> None:
    result = _run(args, "Deleting task...", lambda c: planner.delete_task(c, args.task_id, confirm=args.yes))
    _report(args, result, "deleted", "Task deleted.")


def cmd_planner_bucket_update(args: argparse.Namespace) -> None:
    result = _run(args, "Updating bucket...", lambda c: planner.update_bucket(
        c, args.bucket_id, name=args.name, after=args.after, first=args.first, confirm=args.yes))
    _report(args, result, "updated", "Bucket updated.")


def cmd_planner_bucket_delete(args: argparse.Namespace) -> None:
    result = _run(args, "Deleting bucket...", lambda c: planner.delete_bucket(c, args.bucket_id, confirm=args.yes))
    if not args.json and result["blocked"]:
        runtime.console.print(f"[red]Blocked: the bucket holds {result['tasksInBucket']} task(s). Move or delete them first.[/]")
    _report(args, result, "deleted", "Bucket deleted.")


def cmd_planner_bucket_create(args: argparse.Namespace) -> None:
    result = _run(args, "Creating bucket...", lambda c: planner.create_bucket(
        c, args.plan_id, args.name, confirm=args.yes))
    _report(args, result, "created", "Bucket created.")


def _yes(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--yes", action="store_true", help="Apply the change. Without it the command only previews.")


def register(sub) -> None:
    p = sub.add_parser("planner", help="Microsoft Planner plans, buckets and tasks")
    add_output_args(p)
    p.epilog = "IDs can start with a dash. Then put -- before the positional arguments, for example: planner show -- <task-id>"
    ps = p.add_subparsers(dest="command", required=True)

    c = ps.add_parser("plans", help="List plans you recently used or pinned, or the plans of one group")
    add_output_args(c)
    c.add_argument("--group", help="Group ID. Default: plans you recently used or pinned")
    c.add_argument("--count", "-n", type=int, default=50, help="Max plans")
    c.set_defaults(func=cmd_planner_plans)

    c = ps.add_parser("groups", help="List the Microsoft 365 groups you belong to (they own plans)")
    add_output_args(c)
    c.add_argument("--search", help="Only groups whose name contains this text")
    c.add_argument("--count", "-n", type=int, default=100, help="Max groups")
    c.set_defaults(func=cmd_planner_groups)

    c = ps.add_parser("members", help="List who can work on a plan (the members of its group)")
    add_output_args(c)
    c.add_argument("plan_id", help="Full plan ID")
    c.set_defaults(func=cmd_planner_members)

    c = ps.add_parser("member-add", help="Add people to a plan's group (preview unless --yes)")
    add_output_args(c)
    c.add_argument("plan_id", help="Full plan ID")
    c.add_argument("people", nargs="+", help="Email addresses or user IDs")
    _yes(c)
    c.set_defaults(func=cmd_planner_member_add)

    c = ps.add_parser("pin", help="Pin a plan to My plans (preview unless --yes)")
    add_output_args(c)
    c.add_argument("plan_id", help="Full plan ID")
    _yes(c)
    c.set_defaults(func=cmd_planner_pin)

    c = ps.add_parser("buckets", help="List the buckets of a plan, left to right as the board shows them")
    add_output_args(c)
    c.add_argument("plan_id", help="Full plan ID")
    c.set_defaults(func=cmd_planner_buckets)

    c = ps.add_parser("tasks", help="List tasks of a plan or bucket (default: assigned to you)")
    add_output_args(c)
    c.add_argument("--plan", help="Full plan ID")
    c.add_argument("--bucket", help="Full bucket ID")
    c.add_argument("--count", "-n", type=int, default=50, help="Max tasks")
    c.set_defaults(func=cmd_planner_tasks)

    c = ps.add_parser("show", help="Show a task with its description and checklist")
    add_output_args(c)
    c.add_argument("task_id", help="Full task ID")
    c.set_defaults(func=cmd_planner_show)

    c = ps.add_parser("create", help="Create a task (preview unless --yes)")
    add_output_args(c)
    c.add_argument("plan_id", help="Full plan ID")
    c.add_argument("bucket_id", help="Full bucket ID")
    c.add_argument("title", help="Task title")
    c.add_argument("--assign", action="append", help="Email or user ID. Repeatable. The person is notified")
    c.add_argument("--due", help="Due date/time in local time")
    c.add_argument("--priority", choices=list(planner.PRIORITIES))
    c.add_argument("--description", help="Task notes")
    _yes(c)
    c.set_defaults(func=cmd_planner_create)

    c = ps.add_parser("update", help="Edit or move a task (preview unless --yes)")
    add_output_args(c)
    c.add_argument("task_id", help="Full task ID")
    c.add_argument("--title")
    c.add_argument("--bucket", help="Move to this full bucket ID")
    c.add_argument("--assign", action="append", help="Email or user ID. Repeatable")
    c.add_argument("--unassign", action="append", help="Email or user ID. Repeatable")
    c.add_argument("--due", help="Due date/time in local time")
    c.add_argument("--clear-due", action="store_true", help="Remove the due date")
    c.add_argument("--priority", choices=list(planner.PRIORITIES))
    c.add_argument("--percent", type=int, help="Progress 0-100. 100 completes the task")
    c.add_argument("--description", help="Replace the task notes")
    _yes(c)
    c.set_defaults(func=cmd_planner_update)

    c = ps.add_parser("bucket-update", help="Rename or move a bucket (preview unless --yes)")
    add_output_args(c)
    c.add_argument("bucket_id", help="Full bucket ID")
    c.add_argument("--name", help="New name")
    c.add_argument("--after", help="Place the bucket right after (to the right of) this bucket ID (same plan)")
    c.add_argument("--first", action="store_true", help="Place the bucket first (leftmost on the board)")
    _yes(c)
    c.set_defaults(func=cmd_planner_bucket_update)

    c = ps.add_parser("bucket-delete", help="Delete an empty bucket (refused while it holds tasks; preview unless --yes)")
    add_output_args(c)
    c.add_argument("bucket_id", help="Full bucket ID")
    _yes(c)
    c.set_defaults(func=cmd_planner_bucket_delete)

    c = ps.add_parser("delete", help="Delete a task (preview unless --yes)")
    add_output_args(c)
    c.add_argument("task_id", help="Full task ID")
    _yes(c)
    c.set_defaults(func=cmd_planner_delete)

    c = ps.add_parser("bucket-create", help="Create a bucket (preview unless --yes)")
    add_output_args(c)
    c.add_argument("plan_id", help="Full plan ID")
    c.add_argument("name", help="Bucket name")
    _yes(c)
    c.set_defaults(func=cmd_planner_bucket_create)
