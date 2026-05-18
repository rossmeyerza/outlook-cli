from __future__ import annotations

import argparse
import base64
import re
import sys
from pathlib import Path
from typing import Any, Callable

from rich.panel import Panel
from rich.table import Table

from ..cache import MAIL_CACHE, MAIL_FOLDER_CACHE, load_cache, save_cache
from ..errors import OutlookAPIError
from ..links import extract_links_from_html, filter_share_links, looks_like_share_url


JsonDict = dict[str, Any]


def build_ctx(
    *,
    console: Any,
    get_client: Callable[[], Any],
    html_to_text: Callable[[str], str],
    format_datetime: Callable[[str], str],
    resolve_message_id: Callable[[Any, str], str],
) -> dict[str, Any]:
    return {
        "console": console,
        "get_client": get_client,
        "html_to_text": html_to_text,
        "format_datetime": format_datetime,
        "resolve_message_id": resolve_message_id,
    }


def _ctx(args: argparse.Namespace) -> dict[str, Any]:
    return args._mail_ctx


def _console(args: argparse.Namespace):
    return _ctx(args)["console"]


def _get_client(args: argparse.Namespace):
    return _ctx(args)["get_client"]()


def _html_to_text(args: argparse.Namespace, html: str) -> str:
    return _ctx(args)["html_to_text"](html)


def _format_datetime(args: argparse.Namespace, value: str) -> str:
    return _ctx(args)["format_datetime"](value)


def _resolve_message_id(args: argparse.Namespace, client: Any, ref: str) -> str:
    return _ctx(args)["resolve_message_id"](client, ref)


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
            _format_datetime(args, msg.get("ReceivedDateTime", "")),
            from_name,
            subject,
            attach,
            style=style,
        )

    _console(args).print(table)
    _console(args).print("[dim]Use 'outlook-cli mail read <n>' to read an email.[/]")


def _resolve_folder_id(args: argparse.Namespace, ref: str) -> str:
    if ref in {"inbox", "archive", "drafts", "sentitems", "deleteditems", "junkemail"}:
        return ref
    cached = load_cache(MAIL_FOLDER_CACHE)
    if ref.isdigit():
        idx = int(ref)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        _console(args).print("[red]No matching cached folder. Run 'outlook-cli mail folders' first.[/]")
        sys.exit(1)
    matches = [item for item in cached if item.get("DisplayName", "").lower() == ref.lower()]
    if len(matches) == 1:
        return matches[0]["Id"]
    if len(matches) > 1:
        _console(args).print(f"[red]Ambiguous folder name '{ref}'. Use folder index or ID.[/]")
        sys.exit(1)
    return ref


def _safe_filename(name: str) -> str:
    name = re.sub(r"[\\/:*?\"<>|]+", "_", name).strip()
    return name or "attachment"


def cmd_unread(args: argparse.Namespace) -> None:
    """List unread emails from the Inbox."""
    client = _get_client(args)
    try:
        messages = client.list_unread(top=args.count)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to list unread emails: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not messages:
        _console(args).print("[dim]No unread emails.[/]")
        return

    _render_message_table(args, messages, f"Unread emails ({len(messages)})")


def cmd_mail(args: argparse.Namespace) -> None:
    """Search emails by keyword."""
    client = _get_client(args)
    try:
        messages = client.search_messages(args.query, top=args.count)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to search emails: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not messages:
        _console(args).print(f"[dim]No emails found for '{args.query}'.[/]")
        return

    _render_message_table(args, messages, f"Emails matching '{args.query}' ({len(messages)})")


def cmd_read(args: argparse.Namespace) -> None:
    """Read a specific email by ID or search index."""
    client = _get_client(args)
    message_id = _resolve_message_id(args, client, args.message_id)
    try:
        msg = client.get_message(message_id)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to read email: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

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
        body_content = _html_to_text(args, body_content)

    lines.append("")
    lines.append(body_content)

    _console(args).print(Panel("\n".join(lines), title="Email", border_style="cyan"))


def cmd_mark_read(args: argparse.Namespace) -> None:
    client = _get_client(args)
    message_id = _resolve_message_id(args, client, args.message_id)
    try:
        client.update_message_read_state(message_id, is_read=True)
        _console(args).print("[green]Marked email as read.[/]")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to mark email read: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_mark_unread(args: argparse.Namespace) -> None:
    client = _get_client(args)
    message_id = _resolve_message_id(args, client, args.message_id)
    try:
        client.update_message_read_state(message_id, is_read=False)
        _console(args).print("[green]Marked email as unread.[/]")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to mark email unread: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_archive(args: argparse.Namespace) -> None:
    client = _get_client(args)
    message_id = _resolve_message_id(args, client, args.message_id)
    try:
        client.archive_message(message_id)
        _console(args).print("[green]Archived email.[/]")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to archive email: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_move(args: argparse.Namespace) -> None:
    client = _get_client(args)
    message_id = _resolve_message_id(args, client, args.message_id)
    folder_id = _resolve_folder_id(args, args.folder)
    try:
        moved = client.move_message(message_id, folder_id)
        _console(args).print(f"[green]Moved email:[/] {moved.get('Subject', message_id)}")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to move email: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_folders(args: argparse.Namespace) -> None:
    client = _get_client(args)
    try:
        folders = client.list_mail_folders(top=args.count)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to list folders: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    save_cache(MAIL_FOLDER_CACHE, folders)
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
    _console(args).print(table)


def cmd_attachments(args: argparse.Namespace) -> None:
    client = _get_client(args)
    message_id = _resolve_message_id(args, client, args.message_id)
    try:
        attachments = client.list_message_attachments(message_id)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to list attachments: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not attachments:
        _console(args).print("[dim]No attachments found.[/]")
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
    _console(args).print(table)


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
    client = _get_client(args)
    message_id = _resolve_message_id(args, client, args.message_id)
    out_dir = Path(args.dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    skip_inline = not getattr(args, "include_inline", False)
    try:
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
        _console(args).print(f"[red]Failed to download attachments: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not downloaded:
        _console(args).print("[dim]No downloadable file attachments found.[/]")
        return
    for path in downloaded:
        _console(args).print(f"[green]Downloaded:[/] {path}")


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
    client = _get_client(args)
    message_id = _resolve_message_id(args, client, args.message_id)
    try:
        message = client.get_message(message_id)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to read email: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    links = _collect_message_links(message, share_only=args.share_only)
    if not links:
        _console(args).print("[dim]No links found in this email body.[/]")
        return

    table = Table(title=f"Links ({len(links)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Kind", width=8)
    table.add_column("Label", ratio=1, overflow="fold")
    table.add_column("URL", ratio=2, overflow="fold")
    for i, link in enumerate(links, 1):
        table.add_row(str(i), link["kind"], link.get("label", "") or "", link["url"])
    _console(args).print(table)


def cmd_download_links(args: argparse.Namespace) -> None:
    """Download SharePoint/OneDrive links from an email body via Graph."""
    client = _get_client(args)
    message_id = _resolve_message_id(args, client, args.message_id)
    try:
        message = client.get_message(message_id)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to read email: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    candidates = _collect_message_links(message, share_only=True)
    if not candidates:
        _console(args).print("[dim]No SharePoint/OneDrive links found in this email body.[/]")
        return

    out_dir = Path(args.dir).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)

    from ..cli import _get_graph_client  # avoid circular import at module load
    graph = _get_graph_client()
    downloaded: list[Path] = []
    try:
        for link in candidates:
            url = link["url"]
            if not looks_like_share_url(url):
                continue
            try:
                metadata = graph.get_share_drive_item(url)
                content = graph.download_share_url(url)
            except OutlookAPIError as e:
                _console(args).print(f"[yellow]Skipped {url}: {e}[/]")
                continue
            filename = _safe_filename(metadata.get("name") or link.get("label") or "download")
            path = _unique_path(out_dir, filename, overwrite=args.overwrite)
            path.write_bytes(content)
            downloaded.append(path)
    finally:
        graph.close()

    if not downloaded:
        _console(args).print("[dim]No downloadable links resolved.[/]")
        return
    for path in downloaded:
        _console(args).print(f"[green]Downloaded:[/] {path}")
