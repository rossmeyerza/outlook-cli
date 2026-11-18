"""Commands for Microsoft 365 group inboxes."""
from __future__ import annotations

import argparse
import json
import time
from datetime import UTC, datetime, timedelta

from rich.table import Table
from rich.text import Text

from .. import group_mail, runtime
from ..errors import CliError, OutlookAPIError
from ..formatting import html_to_text
from ..parser import add_output_args
from ..progress import spinner

OVERLAP = timedelta(minutes=2)


def _positive(value: str) -> int:
    number = int(value)
    if number < 1:
        raise argparse.ArgumentTypeError("Enter a positive integer.")
    return number


def _after(value: str) -> datetime:
    try:
        return group_mail.parse_timestamp(value)
    except ValueError as error:
        raise argparse.ArgumentTypeError("Enter an ISO date or timestamp.") from error


def _list(args: argparse.Namespace, client) -> None:
    if args.command == "list":
        result = client.list_planner_groups(search=args.search, top=args.count)
        columns = {"displayName": "Group", "id": "Group ID"}
    else:
        result = sorted(group_mail.threads(client, args.group_id), key=group_mail.thread_time, reverse=True)
        result = result[:args.count]
        columns = {"lastDeliveredDateTime": "Last delivery", "topic": "Subject", "id": "Thread ID"}
    if args.json:
        runtime.print_json(result)
        return
    table = Table()
    for title in columns.values():
        table.add_column(title, overflow="fold")
    for item in result:
        table.add_row(*(Text(item.get(key) or "") for key in columns))
    runtime.console.print(table)


def _posts(args: argparse.Namespace, client) -> None:
    posts = sorted(
        group_mail.posts(client, args.group_id, args.thread_id), key=group_mail.post_time, reverse=True,
    )[:args.count]
    if args.json:
        runtime.print_json(posts)
        return
    for post in posts:
        body = post.get("body") or {}
        content = body.get("content") or ""
        if (body.get("contentType") or "").lower() == "html":
            content = html_to_text(content)
        label = f"{group_mail.post_time(post).isoformat()}  {group_mail.sender(post)}"
        runtime.console.print(Text(label, style="bold"))
        runtime.console.print(Text(content + "\n"))


def _watch(args: argparse.Namespace, client) -> None:
    start = args.after or datetime.now(UTC)
    after = start
    seen = {}
    runtime.error_console.print("Watching the group inbox. Press Ctrl+C to stop.")
    while True:
        poll_start = datetime.now(UTC)
        try:
            posts = group_mail.recent_posts(client, args.group_id, after)
        except OutlookAPIError as error:
            if error.status not in {0, 429, 500, 502, 503, 504}:
                raise
            runtime.error_console.print(f"{error}. Retrying in {args.poll} seconds.", markup=False)
        else:
            for post in posts:
                key = post["threadId"], post["postId"]
                if key in seen:
                    continue
                if args.json:
                    print(json.dumps(post, ensure_ascii=False), flush=True)
                else:
                    runtime.console.print(
                        f"{post['receivedDateTime']}  {post['sender']}  {post['subject']}",
                        markup=False, highlight=False,
                    )
                seen[key] = group_mail.parse_timestamp(post["receivedDateTime"])
            # Keep a short overlap for posts that take time to appear in Graph.
            after = max(after, poll_start - OVERLAP)
            seen = {key: when for key, when in seen.items() if when >= after}
        time.sleep(args.poll)


def _run(args: argparse.Namespace) -> None:
    client = runtime.get_graph_client()
    try:
        if args.command == "watch":
            _watch(args, client)
        else:
            with spinner(args, "Reading Microsoft 365 groups..."):
                if args.command == "posts":
                    _posts(args, client)
                else:
                    _list(args, client)
    except KeyboardInterrupt:
        if args.command != "watch":
            raise
    except (OutlookAPIError, ValueError) as error:
        raise CliError(str(error)) from error
    finally:
        client.close()


def register(sub) -> None:
    groups = sub.add_parser("groups", help="Read and watch Microsoft 365 group inboxes")
    add_output_args(groups)
    commands = groups.add_subparsers(dest="command", required=True)
    listing = commands.add_parser("list", help="List your Microsoft 365 groups")
    listing.add_argument("--search", help="Filter by group name")
    threads = commands.add_parser("threads", help="List group inbox threads")
    posts = commands.add_parser("posts", help="Read messages in a group thread")
    watch = commands.add_parser("watch", help="Watch for new messages and replies")
    watch.add_argument("--poll", type=_positive, default=60, help="Seconds between checks (default 60)")
    watch.add_argument("--after", type=_after, help="Start at this ISO date or timestamp (default now)")
    for command in (listing, threads, posts, watch):
        add_output_args(command)
        command.set_defaults(func=_run)
        if command is not watch:
            command.add_argument("-n", "--count", type=_positive, default=20, help="Maximum results (default 20)")
    for command in (threads, posts, watch):
        command.add_argument("group_id", help="Full group ID from groups list")
    posts.add_argument("thread_id", help="Full thread ID from groups threads")
