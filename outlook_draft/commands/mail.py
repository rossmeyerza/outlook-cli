from __future__ import annotations

import argparse
import base64
import re
from pathlib import Path
from typing import Any

from rich.panel import Panel
from rich.table import Table

from .. import runtime
from ..cache import MAIL_CACHE, MAIL_FOLDER_CACHE, load_cache, save_cache
from ..errors import CliError, OutlookAPIError
from ..formatting import format_datetime, html_to_text
from ..links import extract_links_from_html, filter_share_links, looks_like_share_url
from ..parser import _add_yes_arg, add_output_args
from ..progress import spinner
from ..references import _resolve_message_id
from . import drafts

JsonDict = dict[str, Any]


def _message_summary(args: argparse.Namespace, msg: JsonDict, index: int | None = None) -> JsonDict:
    from_obj = msg.get("From", {}).get("EmailAddress", {})
    return {
        "index": index,
        "id": msg.get("Id", ""),
        "subject": msg.get("Subject", ""),
        "from": {
            "name": from_obj.get("Name", ""),
            "address": from_obj.get("Address", ""),
        },
        "receivedDateTime": msg.get("ReceivedDateTime", ""),
        "received": format_datetime(msg.get("ReceivedDateTime", "")),
        "isRead": msg.get("IsRead", True),
        "hasAttachments": msg.get("HasAttachments", False),
        "importance": msg.get("Importance", ""),
    }


def _render_message_table(args: argparse.Namespace, messages: list[JsonDict], title: str) -> None:
    """Render a table of email messages and cache results for read."""
    save_cache(MAIL_CACHE, messages)

    table = Table(title=title)
    table.add_column("#", style="dim", width=3)
    table.add_column("Date", style="dim", width=12)
    table.add_column("From", width=25)
    table.add_column("Subject", ratio=1)
    table.add_column("", width=3)

    for i, msg in enumerate(messages, 1):
        from_obj = msg.get("From", {}).get("EmailAddress", {})
        from_name = from_obj.get("Name", from_obj.get("Address", ""))
        subject = msg.get("Subject", "(no subject)")
        read = msg.get("IsRead", True)
        attach = "📎" if msg.get("HasAttachments") else ""
        style = "" if read else "bold"
        table.add_row(
            str(i),
            format_datetime(msg.get("ReceivedDateTime", "")),
            from_name,
            subject,
            attach,
            style=style,
        )

    runtime.console.print(table)
    runtime.console.print("[dim]Use 'outlook-cli mail read <n>' to read an email.[/]")


def _resolve_folder_id(args: argparse.Namespace, ref: str) -> str:
    if ref in {"inbox", "archive", "drafts", "sentitems", "deleteditems", "junkemail"}:
        return ref
    cached = load_cache(MAIL_FOLDER_CACHE)
    if ref.isdigit():
        idx = int(ref)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        raise CliError("[red]No matching cached folder. Run 'outlook-cli mail folders' first.[/]")
    matches = [item for item in cached if item.get("DisplayName", "").lower() == ref.lower()]
    if len(matches) == 1:
        return matches[0]["Id"]
    if len(matches) > 1:
        raise CliError(f"[red]Ambiguous folder name '{ref}'. Use folder index or ID.[/]")
    return ref


def _safe_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name).strip()
    return name or "attachment"


def cmd_unread(args: argparse.Namespace) -> None:
    """List unread emails from the Inbox."""
    client = runtime.get_client()
    try:
        with spinner(args, "Loading unread emails..."):
            messages = client.list_unread(top=args.count)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to list unread emails: {e}[/]") from e
    finally:
        client.close()

    if not messages:
        if args.json:
            runtime.print_json([])
            return
        runtime.console.print("[dim]No unread emails.[/]")
        return

    if args.json:
        runtime.print_json([_message_summary(args, msg, i) for i, msg in enumerate(messages, 1)])
        return

    _render_message_table(args, messages, f"Unread emails ({len(messages)})")


def cmd_mail(args: argparse.Namespace) -> None:
    """Search emails by keyword."""
    client = runtime.get_client()
    try:
        with spinner(args, f"Searching emails for '{args.query}'..."):
            messages = client.search_messages(args.query, top=args.count)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to search emails: {e}[/]") from e
    finally:
        client.close()

    if not messages:
        if args.json:
            runtime.print_json([])
            return
        runtime.console.print(f"[dim]No emails found for '{args.query}'.[/]")
        return

    if args.json:
        runtime.print_json([_message_summary(args, msg, i) for i, msg in enumerate(messages, 1)])
        return

    _render_message_table(args, messages, f"Emails matching '{args.query}' ({len(messages)})")


def cmd_read(args: argparse.Namespace) -> None:
    """Read a specific email by ID or search index."""
    client = runtime.get_client()
    message_id = _resolve_message_id(client, args.message_id)
    try:
        with spinner(args, "Loading email..."):
            msg = client.get_message(message_id)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to read email: {e}[/]") from e
    finally:
        client.close()

    if args.json:
        data = dict(msg)
        body_obj = data.get("Body", {})
        if body_obj.get("ContentType") == "HTML":
            body_obj = dict(body_obj)
            body_obj["Text"] = html_to_text(body_obj.get("Content", ""))
            data["Body"] = body_obj
        runtime.print_json(data)
        return

    from_obj = msg.get("From", {}).get("EmailAddress", {})
    from_str = f"{from_obj.get('Name', '')} <{from_obj.get('Address', '')}>"
    to_str = ", ".join(
        r.get("EmailAddress", {}).get("Address", "")
        for r in msg.get("ToRecipients", [])
    )
    cc_str = ", ".join(
        r.get("EmailAddress", {}).get("Address", "")
        for r in msg.get("CcRecipients", [])
    )

    lines = [
        f"[bold]Subject:[/] {msg.get('Subject', '')}",
        f"[bold]From:[/] {from_str}",
        f"[bold]To:[/] {to_str}",
    ]
    if cc_str:
        lines.append(f"[bold]CC:[/] {cc_str}")
    lines.append(f"[bold]Date:[/] {msg.get('ReceivedDateTime', '')}")
    lines.append(f"[bold]Importance:[/] {msg.get('Importance', 'Normal')}")
    if msg.get("HasAttachments"):
        lines.append("[bold]Attachments:[/] Yes")

    body_obj = msg.get("Body", {})
    body_content = body_obj.get("Content", "")
    if body_obj.get("ContentType") == "HTML":
        body_content = html_to_text(body_content)

    lines.append("")
    lines.append(body_content)

    runtime.console.print(Panel("\n".join(lines), title="Email", border_style="cyan"))


def cmd_mark_read(args: argparse.Namespace) -> None:
    client = runtime.get_client()
    message_id = _resolve_message_id(client, args.message_id)
    try:
        client.update_message_read_state(message_id, is_read=True)
        runtime.console.print("[green]Marked email as read.[/]")
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to mark email read: {e}[/]") from e
    finally:
        client.close()


def cmd_mark_unread(args: argparse.Namespace) -> None:
    client = runtime.get_client()
    message_id = _resolve_message_id(client, args.message_id)
    try:
        client.update_message_read_state(message_id, is_read=False)
        runtime.console.print("[green]Marked email as unread.[/]")
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to mark email unread: {e}[/]") from e
    finally:
        client.close()


def cmd_archive(args: argparse.Namespace) -> None:
    client = runtime.get_client()
    message_id = _resolve_message_id(client, args.message_id)
    try:
        client.archive_message(message_id)
        runtime.console.print("[green]Archived email.[/]")
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to archive email: {e}[/]") from e
    finally:
        client.close()


def cmd_move(args: argparse.Namespace) -> None:
    client = runtime.get_client()
    message_id = _resolve_message_id(client, args.message_id)
    folder_id = _resolve_folder_id(args, args.folder)
    try:
        moved = client.move_message(message_id, folder_id)
        runtime.console.print(f"[green]Moved email:[/] {moved.get('Subject', message_id)}")
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to move email: {e}[/]") from e
    finally:
        client.close()


def cmd_folders(args: argparse.Namespace) -> None:
    client = runtime.get_client()
    try:
        with spinner(args, "Loading mail folders..."):
            folders = client.list_mail_folders(top=args.count)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to list folders: {e}[/]") from e
    finally:
        client.close()

    save_cache(MAIL_FOLDER_CACHE, folders)
    if args.json:
        runtime.print_json([
            {
                "index": i,
                "id": folder.get("Id", ""),
                "displayName": folder.get("DisplayName", ""),
                "unreadItemCount": folder.get("UnreadItemCount", 0),
                "totalItemCount": folder.get("TotalItemCount", 0),
            }
            for i, folder in enumerate(folders, 1)
        ])
        return

    table = Table(title=f"Mail folders ({len(folders)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Name", ratio=1)
    table.add_column("Unread", width=8)
    table.add_column("Total", width=8)
    table.add_column("ID", style="dim", width=16, no_wrap=True)
    for i, folder in enumerate(folders, 1):
        table.add_row(
            str(i),
            folder.get("DisplayName", ""),
            str(folder.get("UnreadItemCount", "")),
            str(folder.get("TotalItemCount", "")),
            folder.get("Id", "")[-16:],
        )
    runtime.console.print(table)


def cmd_attachments(args: argparse.Namespace) -> None:
    client = runtime.get_client()
    message_id = _resolve_message_id(client, args.message_id)
    try:
        with spinner(args, "Loading attachments..."):
            attachments = client.list_message_attachments(message_id)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to list attachments: {e}[/]") from e
    finally:
        client.close()

    if not attachments:
        if args.json:
            runtime.print_json([])
            return
        runtime.console.print("[dim]No attachments found.[/]")
        return

    if args.json:
        runtime.print_json([
            {
                "index": i,
                "id": attachment.get("Id", ""),
                "name": attachment.get("Name", ""),
                "contentType": attachment.get("ContentType", ""),
                "size": attachment.get("Size", 0),
                "isInline": attachment.get("IsInline", False),
            }
            for i, attachment in enumerate(attachments, 1)
        ])
        return

    table = Table(title=f"Attachments ({len(attachments)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Name", ratio=1)
    table.add_column("Type", width=28)
    table.add_column("Size", width=10)
    table.add_column("Inline", width=8)
    for i, attachment in enumerate(attachments, 1):
        table.add_row(
            str(i),
            attachment.get("Name", ""),
            attachment.get("ContentType", ""),
            str(attachment.get("Size", "")),
            "yes" if attachment.get("IsInline") else "no",
        )
    runtime.console.print(table)


def _unique_path(out_dir: Path, filename: str, *, overwrite: bool) -> Path:
    path = out_dir / filename
    if not path.exists() or overwrite:
        return path
    stem, suffix = path.stem, path.suffix
    n = 2
    while (out_dir / f"{stem}-{n}{suffix}").exists():
        n += 1
    return out_dir / f"{stem}-{n}{suffix}"


def cmd_download_attachments(args: argparse.Namespace) -> None:
    """Download all file attachments from a message, including inline ones."""
    client = runtime.get_client()
    message_id = _resolve_message_id(client, args.message_id)
    out_dir = Path(args.dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    skip_inline = not getattr(args, "include_inline", False)
    try:
        with spinner(args, "Downloading email attachments..."):
            attachments = client.list_message_attachments(message_id)
            downloaded: list[Path] = []
            for attachment in attachments:
                if skip_inline and attachment.get("IsInline"):
                    continue
                full = client.get_message_attachment(message_id, attachment.get("Id", ""))
                content = full.get("ContentBytes")
                if not content:
                    continue
                filename = _safe_filename(full.get("Name") or "attachment")
                path = _unique_path(out_dir, filename, overwrite=args.overwrite)
                path.write_bytes(base64.b64decode(content))
                downloaded.append(path)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to download attachments: {e}[/]") from e
    finally:
        client.close()

    if not downloaded:
        runtime.console.print("[dim]No downloadable file attachments found.[/]")
        return
    for path in downloaded:
        runtime.console.print(f"[green]Downloaded:[/] {path}")


def _collect_message_links(message: JsonDict, share_only: bool) -> list[JsonDict]:
    body = message.get("Body") or {}
    if (body.get("ContentType") or "").lower() != "html":
        return []
    links = extract_links_from_html(body.get("Content") or "")
    if share_only:
        links = filter_share_links(links)
    return links


def cmd_links(args: argparse.Namespace) -> None:
    """List shareable links and inline image URLs in an email body."""
    client = runtime.get_client()
    message_id = _resolve_message_id(client, args.message_id)
    try:
        with spinner(args, "Loading email links..."):
            message = client.get_message(message_id)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to read email: {e}[/]") from e
    finally:
        client.close()

    links = _collect_message_links(message, share_only=args.share_only)
    if not links:
        if args.json:
            runtime.print_json([])
            return
        runtime.console.print("[dim]No links found in this email body.[/]")
        return

    if args.json:
        runtime.print_json([
            {"index": i, **link}
            for i, link in enumerate(links, 1)
        ])
        return

    table = Table(title=f"Links ({len(links)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Kind", width=8)
    table.add_column("Label", ratio=1, overflow="fold")
    table.add_column("URL", ratio=2, overflow="fold")
    for i, link in enumerate(links, 1):
        table.add_row(str(i), link["kind"], link.get("label", "") or "", link["url"])
    runtime.console.print(table)


def cmd_download_links(args: argparse.Namespace) -> None:
    """Download SharePoint/OneDrive links from an email body via Graph."""
    client = runtime.get_client()
    message_id = _resolve_message_id(client, args.message_id)
    try:
        with spinner(args, "Loading email links..."):
            message = client.get_message(message_id)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to read email: {e}[/]") from e
    finally:
        client.close()

    candidates = _collect_message_links(message, share_only=True)
    if not candidates:
        runtime.console.print("[dim]No SharePoint/OneDrive links found in this email body.[/]")
        return

    out_dir = Path(args.dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)

    from ..cli import _get_graph_client  # avoid circular import at module load
    graph = _get_graph_client()
    downloaded: list[Path] = []
    try:
        with spinner(args, "Downloading linked files..."):
            for link in candidates:
                url = link["url"]
                if not looks_like_share_url(url):
                    continue
                try:
                    metadata = graph.get_share_drive_item(url)
                    content = graph.download_share_url(url)
                except OutlookAPIError as e:
                    runtime.console.print(f"[yellow]Skipped {url}: {e}[/]")
                    continue
                filename = _safe_filename(metadata.get("name") or link.get("label") or "download")
                path = _unique_path(out_dir, filename, overwrite=args.overwrite)
                path.write_bytes(content)
                downloaded.append(path)
    finally:
        graph.close()

    if not downloaded:
        runtime.console.print("[dim]No downloadable links resolved.[/]")
        return
    for path in downloaded:
        runtime.console.print(f"[green]Downloaded:[/] {path}")


def register(sub) -> None:
    p_mail = sub.add_parser("mail", help="Read and search emails")
    add_output_args(p_mail)
    mail_sub = p_mail.add_subparsers(dest="command", required=True)

    _register_cmd_mail_unread(mail_sub)
    _register_cmd_mail_search(mail_sub)
    _register_cmd_mail_read(mail_sub)
    _register_cmd_mail_mark_read(mail_sub)
    _register_cmd_mail_mark_unread(mail_sub)
    _register_cmd_mail_archive(mail_sub)
    _register_cmd_mail_move(mail_sub)
    _register_cmd_mail_folders(mail_sub)
    _register_cmd_mail_attachments(mail_sub)
    _register_cmd_mail_download_attachments(mail_sub)
    _register_cmd_mail_links(mail_sub)
    _register_cmd_mail_download_links(mail_sub)
    _register_cmd_mail_send(mail_sub)


def _register_cmd_mail_unread(mail_sub) -> None:
    cmd_mail_unread = mail_sub.add_parser("unread", help="List unread emails from the Inbox")
    add_output_args(cmd_mail_unread)
    cmd_mail_unread.add_argument("--count", "-n", type=int, default=50, help="Max results")
    cmd_mail_unread.set_defaults(func=cmd_unread)


def _register_cmd_mail_search(mail_sub) -> None:
    cmd_mail_search = mail_sub.add_parser("search", help="Search emails by keyword")
    add_output_args(cmd_mail_search)
    cmd_mail_search.add_argument("query", help="Search term")
    cmd_mail_search.add_argument("--count", "-n", type=int, default=20, help="Max results")
    cmd_mail_search.set_defaults(func=cmd_mail)


def _register_cmd_mail_read(mail_sub) -> None:
    cmd_mail_read = mail_sub.add_parser("read", help="Read an email by index or ID")
    add_output_args(cmd_mail_read)
    cmd_mail_read.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_read.set_defaults(func=cmd_read)


def _register_cmd_mail_mark_read(mail_sub) -> None:
    cmd_mail_mark_read = mail_sub.add_parser("mark-read", help="Mark an email as read")
    cmd_mail_mark_read.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_mark_read.set_defaults(func=cmd_mark_read)


def _register_cmd_mail_mark_unread(mail_sub) -> None:
    cmd_mail_mark_unread = mail_sub.add_parser("mark-unread", help="Mark an email as unread")
    cmd_mail_mark_unread.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_mark_unread.set_defaults(func=cmd_mark_unread)


def _register_cmd_mail_archive(mail_sub) -> None:
    cmd_mail_archive = mail_sub.add_parser("archive", help="Archive an email")
    cmd_mail_archive.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_archive.set_defaults(func=cmd_archive)


def _register_cmd_mail_move(mail_sub) -> None:
    cmd_mail_move = mail_sub.add_parser("move", help="Move an email to another folder")
    cmd_mail_move.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_move.add_argument("--folder", "-f", required=True, help="Folder index, well-known name, folder name, or folder ID")
    cmd_mail_move.set_defaults(func=cmd_move)


def _register_cmd_mail_folders(mail_sub) -> None:
    cmd_mail_folders = mail_sub.add_parser("folders", help="List mail folders")
    add_output_args(cmd_mail_folders)
    cmd_mail_folders.add_argument("--count", "-n", type=int, default=100, help="Max folders")
    cmd_mail_folders.set_defaults(func=cmd_folders)


def _register_cmd_mail_attachments(mail_sub) -> None:
    cmd_mail_attachments = mail_sub.add_parser("attachments", help="List message attachments")
    add_output_args(cmd_mail_attachments)
    cmd_mail_attachments.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_attachments.set_defaults(func=cmd_attachments)


def _register_cmd_mail_download_attachments(mail_sub) -> None:
    cmd_mail_download_attachments = mail_sub.add_parser("download-attachments", help="Download message attachments")
    cmd_mail_download_attachments.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_download_attachments.add_argument("--dir", "-d", default="attachments", help="Output directory")
    cmd_mail_download_attachments.add_argument("--overwrite", action="store_true", help="Overwrite files with matching names")
    cmd_mail_download_attachments.add_argument(
        "--include-inline",
        action="store_true",
        help="Also download attachments marked as inline (e.g. embedded images)",
    )
    cmd_mail_download_attachments.set_defaults(func=cmd_download_attachments)


def _register_cmd_mail_links(mail_sub) -> None:
    cmd_mail_links = mail_sub.add_parser("links", help="List links and inline image URLs in an email body")
    add_output_args(cmd_mail_links)
    cmd_mail_links.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_links.add_argument(
        "--share-only",
        action="store_true",
        help="Only show SharePoint/OneDrive style links",
    )
    cmd_mail_links.set_defaults(func=cmd_links)


def _register_cmd_mail_download_links(mail_sub) -> None:
    cmd_mail_download_links = mail_sub.add_parser(
        "download-links",
        help="Download SharePoint/OneDrive links from an email body via Graph",
    )
    cmd_mail_download_links.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_download_links.add_argument("--dir", "-d", default="attachments", help="Output directory")
    cmd_mail_download_links.add_argument("--overwrite", action="store_true", help="Overwrite files with matching names")
    cmd_mail_download_links.set_defaults(func=cmd_download_links)


def _register_cmd_mail_send(mail_sub) -> None:
    cmd_mail_send = mail_sub.add_parser("send", help="Send a new email immediately (requires --yes)")
    cmd_mail_send.add_argument("--to", required=True, action="append", help="Recipient (repeatable, comma-separated ok)")
    cmd_mail_send.add_argument("--cc", action="append", help="CC recipient")
    cmd_mail_send.add_argument("--bcc", action="append", help="BCC recipient")
    cmd_mail_send.add_argument("--subject", "-s", required=True, help="Subject line")
    cmd_mail_send.add_argument("--body", "-b", help="Body text (inline)")
    cmd_mail_send.add_argument("--body-file", "-f", help="Read body from file")
    cmd_mail_send.add_argument("--html", action="store_true", help="Treat body input as HTML")
    cmd_mail_send.add_argument("--signature", choices=["none", "new", "reply"], default="none", help="Append saved signature")
    cmd_mail_send.add_argument("--importance", choices=["Low", "Normal", "High"], default="Normal")
    cmd_mail_send.add_argument("--no-save", action="store_true", help="Do not save to Sent Items")
    _add_yes_arg(cmd_mail_send)
    cmd_mail_send.set_defaults(func=drafts.cmd_send_mail)
