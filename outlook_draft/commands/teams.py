from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Callable

from rich.panel import Panel
from rich.table import Table

from .. import config
from ..cache import TEAMS_CACHE, load_cache, save_cache
from ..errors import OutlookAPIError
from ..links import extract_links_from_html, looks_like_share_url
from ..progress import spinner


JsonDict = dict[str, Any]
SELF_CHAT_ID = "48:notes"


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


def _normalize_search_text(value: str) -> str:
    return " ".join(value.casefold().split())


def _matches_query(values: list[str], query: str) -> bool:
    normalized_query = _normalize_search_text(query)
    if not normalized_query:
        return False

    haystack = _normalize_search_text(" ".join(value for value in values if value))
    if not haystack:
        return False
    if normalized_query in haystack:
        return True

    terms = normalized_query.split()
    return bool(terms) and all(term in haystack for term in terms)


def _member_search_values(member: JsonDict) -> list[str]:
    return [
        member.get("displayName") or "",
        member.get("email") or "",
        member.get("userId") or "",
        member.get("id") or "",
    ]


def _chat_match_reasons(chat: JsonDict, members: list[JsonDict], query: str) -> list[str]:
    reasons: list[str] = []
    topic = chat.get("topic") or ""
    if topic and _matches_query([topic], query):
        reasons.append("topic")

    member_matches: list[str] = []
    for member in members:
        if _matches_query(_member_search_values(member), query):
            label = _member_label(member)
            if label and label not in member_matches:
                member_matches.append(label)

    if member_matches:
        suffix = f" +{len(member_matches) - 2}" if len(member_matches) > 2 else ""
        reasons.append("member: " + ", ".join(member_matches[:2]) + suffix)

    return reasons


def _chat_title(chat: JsonDict, members: list[JsonDict] | None = None) -> str:
    if chat.get("id") == SELF_CHAT_ID:
        return "Self chat"
    if chat.get("topic"):
        return chat["topic"]

    names: list[str] = []
    self_email = config.MS_EMAIL.lower()
    self_only = False
    if members:
        member_emails = [
            (member.get("email") or "").lower()
            for member in members
            if member.get("email")
        ]
        self_only = bool(member_emails) and all(email == self_email for email in member_emails)
    if self_only:
        return "Self chat"

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
            console.print("[red]No Teams chats cached. Run 'outlook-cli teams list' or 'outlook-cli teams search <query>' first.[/]")
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


def find_self_chat(client: Any, *, top: int = 500) -> tuple[JsonDict, list[JsonDict]] | None:
    """Find the Teams chat whose membership is only the current user."""
    current_user = client.get_current_user()
    self_id = current_user.get("id", "")
    self_email = (
        current_user.get("mail")
        or current_user.get("userPrincipalName")
        or config.MS_EMAIL
        or ""
    ).lower()
    self_member = {
        "displayName": current_user.get("displayName") or "Self",
        "email": self_email,
        "userId": self_id,
    }

    try:
        messages = client.list_teams_message_metadata(SELF_CHAT_ID, top=1)
        latest = messages[0].get("createdDateTime", "") if messages else ""
        return (
            {
                "id": SELF_CHAT_ID,
                "topic": "Self chat",
                "chatType": "oneOnOne",
                "lastUpdatedDateTime": latest,
                "webUrl": "",
            },
            [self_member],
        )
    except OutlookAPIError:
        pass

    chats = client.list_teams_chats(top=top)
    chats.sort(key=lambda chat: chat.get("lastUpdatedDateTime") or "", reverse=True)
    for chat in chats:
        if chat.get("chatType") != "oneOnOne":
            continue
        try:
            members = client.list_teams_chat_members(chat["id"], top=10)
        except OutlookAPIError:
            continue
        if not members:
            continue

        member_ids = [member.get("userId") for member in members if member.get("userId")]
        member_emails = [
            (member.get("email") or "").lower()
            for member in members
            if member.get("email")
        ]
        ids_match = bool(self_id and member_ids) and all(user_id == self_id for user_id in member_ids)
        emails_match = bool(self_email and member_emails) and all(email == self_email for email in member_emails)
        if ids_match or emails_match:
            return chat, members
    return None


def _teams_sender(message: JsonDict) -> str:
    from_obj = message.get("from") or {}
    for key in ("user", "application", "device"):
        entity = from_obj.get(key) or {}
        label = entity.get("displayName") or entity.get("id")
        if label:
            return label
    return "Unknown"


def _message_user_id(message: JsonDict) -> str:
    from_obj = message.get("from") or {}
    user = from_obj.get("user") or {}
    return user.get("id", "")


def _is_received_user_message(message: JsonDict, self_user_id: str) -> bool:
    if message.get("deletedDateTime"):
        return False
    if message.get("messageType") != "message":
        return False
    sender_id = _message_user_id(message)
    return not sender_id or sender_id != self_user_id


def _latest_received_message_time(
    client: Any, chat: dict, self_user_id: str
) -> str:
    """Return the timestamp of the latest message not sent by self.

    Uses lastMessagePreview from the chat object to avoid an extra API call
    in the common case. Only falls back to fetching messages if the last
    message was sent by the current user.
    """
    preview = chat.get("lastMessagePreview") or {}
    preview_time = preview.get("createdDateTime", "")
    preview_sender_id = (
        (preview.get("from") or {}).get("user") or {}
    ).get("id", "")

    # If last message wasn't from self, we can use the preview timestamp directly
    if preview_time and preview_sender_id and preview_sender_id != self_user_id:
        return preview_time

    # If no preview or last message was from self, fetch recent messages
    try:
        messages = client.list_teams_message_metadata(chat["id"], top=50)
    except OutlookAPIError:
        return preview_time  # fall back to preview time rather than empty
    received = [
        message.get("createdDateTime", "")
        for message in messages
        if _is_received_user_message(message, self_user_id)
    ]
    return max(received) if received else ""


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
        with spinner(args, "Loading Teams chats..."):
            chats = client.list_teams_chats(top=args.count)
            if getattr(args, 'sort_received', False):
                # Slow path: fetch messages per chat to find last received
                try:
                    current_user = client.get_current_user()
                    self_user_id = current_user.get("id", "")
                except OutlookAPIError:
                    self_user_id = ""
                for chat in chats:
                    chat["_lastReceivedMessageDateTime"] = _latest_received_message_time(
                        client, chat, self_user_id,
                    )
                chats.sort(
                    key=lambda chat: chat.get("_lastReceivedMessageDateTime") or "",
                    reverse=True,
                )
            else:
                # Fast path: sort by lastUpdatedDateTime (already in response)
                chats.sort(
                    key=lambda chat: chat.get("lastUpdatedDateTime") or "",
                    reverse=True,
                )
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
    table.add_column("Last received", style="dim", width=12)
    table.add_column("Type", width=12)
    table.add_column("Chat", ratio=1)
    table.add_column("ID", style="dim", width=16, no_wrap=True)

    client = _get_graph_client(args)
    try:
        with spinner(args, "Resolving Teams chat participants..."):
            for i, chat in enumerate(chats, 1):
                members = []
                if not chat.get("topic"):
                    try:
                        members = client.list_teams_chat_members(chat["id"], top=10)
                    except OutlookAPIError:
                        members = []
                table.add_row(
                    str(i),
                    _format_datetime(args, chat.get("_lastReceivedMessageDateTime") or chat.get("lastUpdatedDateTime", "")),
                    chat.get("chatType", ""),
                    _chat_title(chat, members),
                    chat.get("id", "")[-16:],
                )
    finally:
        client.close()

    console.print(table)
    console.print("[dim]Use 'outlook-cli teams show <n>' or 'outlook-cli teams messages <n>' to inspect a chat.[/]")


def cmd_teams_search(args: argparse.Namespace) -> None:
    console = _console(args)
    client = _get_graph_client(args)
    query = " ".join(args.query).strip() if isinstance(args.query, list) else args.query.strip()
    if not query:
        console.print("[red]Provide a search query.[/]")
        sys.exit(1)

    scan_count = max(args.scan, 1)
    result_count = max(args.count, 1)
    member_count = max(args.member_count, 1)
    results: list[tuple[JsonDict, list[JsonDict], list[str]]] = []
    scanned = 0
    member_errors = 0

    try:
        with spinner(args, f"Scanning up to {scan_count} Teams chats by topic and participants..."):
            chats = client.list_teams_chats(top=scan_count)
            chats.sort(
                key=lambda chat: chat.get("lastUpdatedDateTime") or "",
                reverse=True,
            )

            for chat in chats:
                scanned += 1
                members: list[JsonDict] = []
                topic = chat.get("topic") or ""
                topic_matches = _matches_query([topic], query)
                if not topic_matches:
                    try:
                        members = client.list_teams_chat_members(chat["id"], top=member_count)
                    except OutlookAPIError:
                        member_errors += 1
                        members = []
                reasons = _chat_match_reasons(chat, members, query)
                if reasons:
                    results.append((chat, members, reasons))
                    if len(results) >= result_count:
                        break
    except OutlookAPIError as e:
        console.print(f"[red]Failed to search Teams chats: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    matched_chats = [chat for chat, _, _ in results]
    save_cache(TEAMS_CACHE, matched_chats, id_key="id")

    if not results:
        console.print(f"[dim]No Teams chats matched '{query}' in {scanned} scanned chats.[/]")
        console.print("[dim]Increase --scan to search more chats.[/]")
        return

    table = Table(title=f"Teams chats matching '{query}' ({len(results)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Updated", style="dim", width=12)
    table.add_column("Type", width=12)
    table.add_column("Match", ratio=1, min_width=18)
    table.add_column("Chat", ratio=2, min_width=24)
    table.add_column("ID", style="dim", width=16, no_wrap=True)

    for i, (chat, members, reasons) in enumerate(results, 1):
        table.add_row(
            str(i),
            _format_datetime(args, chat.get("lastUpdatedDateTime", "")),
            chat.get("chatType", ""),
            "; ".join(reasons),
            _chat_title(chat, members),
            chat.get("id", "")[-16:],
        )

    console.print(table)
    console.print(
        f"[dim]Scanned {scanned} chats. Use 'outlook-cli teams show <n>' or "
        "'outlook-cli teams messages <n>' to inspect a result.[/]"
    )
    if member_errors:
        console.print(f"[dim]Could not read members for {member_errors} chats.[/]")


def cmd_teams_show(args: argparse.Namespace) -> None:
    console = _console(args)
    client = _get_graph_client(args)
    chat_id = _resolve_teams_chat_id(args, args.chat_id)
    try:
        if chat_id == SELF_CHAT_ID:
            found = find_self_chat(client)
            if not found:
                raise OutlookAPIError(404, "Teams self-chat not found")
            chat, members = found
        else:
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


def cmd_teams_self(args: argparse.Namespace) -> None:
    console = _console(args)
    client = _get_graph_client(args)
    try:
        found = find_self_chat(client)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to read Teams self-chat: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not found:
        console.print("[yellow]No Teams self-chat found in Microsoft Graph.[/]")
        return

    chat, members = found
    member_list = ", ".join(_member_label(member) for member in members if _member_label(member))
    lines = [
        f"[bold]Title:[/] {_chat_title(chat, members)}",
        f"[bold]Type:[/] {chat.get('chatType', '')}",
        f"[bold]Updated:[/] {chat.get('lastUpdatedDateTime', '')}",
        f"[bold]Web URL:[/] {chat.get('webUrl', '')}",
        f"[bold]ID:[/] [dim]{chat.get('id', '')}[/]",
    ]
    if member_list:
        lines.append(f"[bold]Members:[/] {member_list}")
    console.print(Panel("\n".join(lines), title="Teams self-chat", border_style="green"))


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


def _collect_message_links_and_attachments(messages: list[JsonDict]) -> list[JsonDict]:
    items: list[JsonDict] = []
    for message in messages:
        if message.get("messageType") != "message":
            continue
        if message.get("deletedDateTime"):
            continue
        sender = _teams_sender(message)
        created = message.get("createdDateTime", "")
        for attachment in message.get("attachments") or []:
            if attachment.get("contentType") != "reference":
                continue
            url = attachment.get("contentUrl") or ""
            if not url:
                continue
            items.append({
                "created": created,
                "sender": sender,
                "name": attachment.get("name") or url,
                "url": url,
                "source": "attachment",
            })
        body = message.get("body") or {}
        if (body.get("contentType") or "").lower() != "html":
            continue
        for link in extract_links_from_html(body.get("content") or ""):
            if not looks_like_share_url(link["url"]):
                continue
            items.append({
                "created": created,
                "sender": sender,
                "name": link.get("label") or link["url"],
                "url": link["url"],
                "source": "body",
            })
    items.sort(key=lambda item: item["created"], reverse=True)
    return items


def _resolve_chat_attachment_index(items: list[JsonDict], ref: str | None) -> list[JsonDict]:
    if not ref:
        return items
    if ref.isdigit():
        idx = int(ref)
        if 1 <= idx <= len(items):
            return [items[idx - 1]]
    matches = [item for item in items if item["url"] == ref]
    if matches:
        return matches
    return [item for item in items if ref.lower() in (item["name"] or "").lower()]


def cmd_teams_attachments(args: argparse.Namespace) -> None:
    """List Teams attachments and SharePoint/OneDrive links in a chat."""
    console = _console(args)
    client = _get_graph_client(args)
    chat_id = _resolve_teams_chat_id(args, args.chat_id)
    try:
        messages = client.list_teams_messages(chat_id, top=args.scan)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to fetch Teams messages: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    items = _collect_message_links_and_attachments(messages)
    if not items:
        console.print("[dim]No attachments or share links found in this chat.[/]")
        return

    table = Table(title=f"Teams attachments ({len(items)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("When", style="dim", width=12)
    table.add_column("From", width=20)
    table.add_column("Source", width=10)
    table.add_column("Name", ratio=1, overflow="fold")
    table.add_column("URL", ratio=2, overflow="fold")
    for i, item in enumerate(items, 1):
        table.add_row(
            str(i),
            _format_datetime(args, item["created"]),
            item["sender"],
            item["source"],
            item["name"],
            item["url"],
        )
    console.print(table)


def cmd_teams_download_attachments(args: argparse.Namespace) -> None:
    """Download Teams attachments/share links via Graph."""
    console = _console(args)
    client = _get_graph_client(args)
    chat_id = _resolve_teams_chat_id(args, args.chat_id)
    try:
        messages = client.list_teams_messages(chat_id, top=args.scan)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to fetch Teams messages: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    items = _collect_message_links_and_attachments(messages)
    selected = _resolve_chat_attachment_index(items, args.attachment)
    if not selected:
        console.print("[dim]No attachments matched.[/]")
        return

    out_dir = Path(args.dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)

    graph = _get_graph_client(args)
    downloaded: list[Path] = []
    try:
        for item in selected:
            url = item["url"]
            try:
                metadata = graph.get_share_drive_item(url)
                content = graph.download_share_url(url)
            except OutlookAPIError as e:
                console.print(f"[yellow]Skipped {url}: {e}[/]")
                continue
            filename = (metadata.get("name") or item["name"] or "download").strip()
            filename = filename.replace("/", "_").replace("\\", "_") or "download"
            path = out_dir / filename
            if path.exists() and not args.overwrite:
                stem, suffix = path.stem, path.suffix
                n = 2
                while (out_dir / f"{stem}-{n}{suffix}").exists():
                    n += 1
                path = out_dir / f"{stem}-{n}{suffix}"
            path.write_bytes(content)
            downloaded.append(path)
    finally:
        graph.close()

    if not downloaded:
        console.print("[dim]Nothing downloaded.[/]")
        return
    for path in downloaded:
        console.print(f"[green]Downloaded:[/] {path}")


def cmd_teams_messages(args: argparse.Namespace) -> None:
    console = _console(args)
    client = _get_graph_client(args)
    chat_id = _resolve_teams_chat_id(args, args.chat_id)
    try:
        if chat_id == SELF_CHAT_ID:
            chat = {"id": SELF_CHAT_ID, "topic": "Self chat", "chatType": "oneOnOne"}
            members = []
        else:
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
