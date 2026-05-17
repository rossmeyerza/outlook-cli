from __future__ import annotations

import argparse
import sys
from typing import Any, Callable

from rich.panel import Panel
from rich.table import Table

from .. import config
from ..cache import TEAMS_CACHE, load_cache, save_cache
from ..errors import OutlookAPIError


JsonDict = dict[str, Any]


def _ctx(args: argparse.Namespace) -> dict[str, Any]:
    return args._teams_ctx


def _console(args: argparse.Namespace):
    return _ctx(args)["console"]


def _get_graph_client(args: argparse.Namespace):
    return _ctx(args)["get_graph_client"]()


def _format_datetime(args: argparse.Namespace, value: str) -> str:
    return _ctx(args)["format_datetime"](value)


def _html_to_text(args: argparse.Namespace, html: str) -> str:
    return _ctx(args)["html_to_text"](html)


def build_ctx(
    *,
    console: Any,
    get_graph_client: Callable[[], Any],
    format_datetime: Callable[[str], str],
    html_to_text: Callable[[str], str],
) -> dict[str, Any]:
    return {
        "console": console,
        "get_graph_client": get_graph_client,
        "format_datetime": format_datetime,
        "html_to_text": html_to_text,
    }


def _member_label(member: JsonDict) -> str:
    return member.get("displayName") or member.get("email") or member.get("userId") or member.get("id", "")


def _chat_title(chat: JsonDict, members: list[JsonDict] | None = None) -> str:
    if chat.get("topic"):
        return chat["topic"]

    names: list[str] = []
    self_email = config.MS_EMAIL.lower()
    for member in members or []:
        email = (member.get("email") or "").lower()
        label = _member_label(member)
        if label and email != self_email and label not in names:
            names.append(label)

    if names:
        suffix = f" +{len(names) - 3}" if len(names) > 3 else ""
        return ", ".join(names[:3]) + suffix

    chat_type = chat.get("chatType", "chat")
    short_id = chat.get("id", "")[-8:]
    return f"{chat_type} {short_id}".strip()


def _resolve_teams_chat_id(args: argparse.Namespace, ref: str) -> str:
    cached = load_cache(TEAMS_CACHE)
    console = _console(args)

    if ref.isdigit():
        idx = int(ref)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print("[red]No Teams chats cached. Run 'outlook-cli teams list' first.[/]")
        else:
            console.print(f"[red]Index {idx} out of range. Only {len(cached)} chats cached.[/]")
        sys.exit(1)

    if len(ref) < 40 and cached:
        matches = [item for item in cached if item["Id"].endswith(ref)]
        if len(matches) == 1:
            return matches[0]["Id"]
        if len(matches) > 1:
            console.print(f"[red]Ambiguous Teams chat ID suffix '{ref}'. Use a longer ID.[/]")
            sys.exit(1)

    return ref


def _teams_sender(message: JsonDict) -> str:
    from_obj = message.get("from", {})
    for key in ("user", "application", "device"):
        entity = from_obj.get(key) or {}
        label = entity.get("displayName") or entity.get("id")
        if label:
            return label
    return "Unknown"


def _teams_body(args: argparse.Namespace, message: JsonDict) -> str:
    if message.get("deletedDateTime"):
        return "[deleted message]"
    if message.get("messageType") == "systemEventMessage":
        return "[system event]"

    body = message.get("body", {})
    content = body.get("content", "")
    if body.get("contentType") == "html":
        content = _html_to_text(args, content)
    content = content.strip()

    if content:
        return content
    if message.get("attachments") or message.get("hostedContents"):
        return "[attachment or rich content]"
    return "[empty message]"


def cmd_teams_list(args: argparse.Namespace) -> None:
    console = _console(args)
    client = _get_graph_client(args)
    try:
        chats = client.list_teams_chats(top=args.count)
        save_cache(TEAMS_CACHE, chats, id_key="id")
    except OutlookAPIError as e:
        console.print(f"[red]Failed to list Teams chats: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not chats:
        console.print("[dim]No Teams chats found.[/]")
        return

    table = Table(title=f"Teams chats ({len(chats)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Updated", style="dim", width=12)
    table.add_column("Type", width=12)
    table.add_column("Chat", ratio=1)
    table.add_column("ID", style="dim", width=16, no_wrap=True)

    client = _get_graph_client(args)
    try:
        for i, chat in enumerate(chats, 1):
            members = []
            if not chat.get("topic"):
                try:
                    members = client.list_teams_chat_members(chat["id"], top=10)
                except OutlookAPIError:
                    members = []
            table.add_row(
                str(i),
                _format_datetime(args, chat.get("lastUpdatedDateTime") or chat.get("createdDateTime", "")),
                chat.get("chatType", ""),
                _chat_title(chat, members),
                chat.get("id", "")[-16:],
            )
    finally:
        client.close()

    console.print(table)
    console.print("[dim]Use 'outlook-cli teams show <n>' or 'outlook-cli teams messages <n>' to inspect a chat.[/]")


def cmd_teams_show(args: argparse.Namespace) -> None:
    console = _console(args)
    client = _get_graph_client(args)
    chat_id = _resolve_teams_chat_id(args, args.chat_id)
    try:
        chat = client.get_teams_chat(chat_id)
        members = client.list_teams_chat_members(chat_id, top=50)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to fetch Teams chat: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    member_list = ", ".join(_member_label(member) for member in members if _member_label(member))
    lines = [
        f"[bold]Title:[/] {_chat_title(chat, members)}",
        f"[bold]Type:[/] {chat.get('chatType', '')}",
        f"[bold]Created:[/] {chat.get('createdDateTime', '')}",
        f"[bold]Updated:[/] {chat.get('lastUpdatedDateTime', '')}",
        f"[bold]Web URL:[/] {chat.get('webUrl', '')}",
        f"[bold]ID:[/] [dim]{chat.get('id', '')}[/]",
    ]
    if member_list:
        lines.append(f"[bold]Members:[/] {member_list}")

    console.print(Panel("\n".join(lines), title="Teams chat", border_style="blue"))


def cmd_teams_send(args: argparse.Namespace) -> None:
    console = _console(args)
    client = _get_graph_client(args)
    chat_id = _resolve_teams_chat_id(args, args.chat_id)
    try:
        client.send_teams_message(
            chat_id,
            args.message,
            content_type="html" if args.html else "text",
        )
        console.print("[green]Teams message sent.[/]")
    except OutlookAPIError as e:
        console.print(f"[red]Failed to send Teams message: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_teams_messages(args: argparse.Namespace) -> None:
    console = _console(args)
    client = _get_graph_client(args)
    chat_id = _resolve_teams_chat_id(args, args.chat_id)
    try:
        chat = client.get_teams_chat(chat_id)
        members = client.list_teams_chat_members(chat_id, top=20)
        messages = client.list_teams_messages(chat_id, top=args.count)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to fetch Teams messages: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not messages:
        console.print("[dim]No messages found in this Teams chat.[/]")
        return

    messages.sort(key=lambda item: item.get("createdDateTime", ""))
    table = Table(title=f"Messages, {_chat_title(chat, members)}")
    table.add_column("Time", style="dim", width=12)
    table.add_column("Sender", width=24)
    table.add_column("Message", ratio=1)

    for message in messages:
        table.add_row(
            _format_datetime(args, message.get("createdDateTime", "")),
            _teams_sender(message),
            _teams_body(args, message),
        )

    console.print(table)
