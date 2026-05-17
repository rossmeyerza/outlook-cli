from __future__ import annotations

import argparse
import sys
from typing import Any, Callable

from rich.panel import Panel
from rich.table import Table

from ..cache import MAIL_CACHE, load_cache, save_cache
from ..errors import OutlookAPIError


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
        attach = "\U0001f4ce" if msg.get("HasAttachments") else ""
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
    _console(args).print("[dim]Use 'outlook-cli read <n>' to read an email.[/]")


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
