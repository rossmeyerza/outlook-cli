"""CLI for managing Outlook mail, drafts, calendar, and contacts.

Usage:
  outlook-cli draft create --to someone@example.com --subject "Hello" --body "Hi there"
  outlook-cli draft reply 1 --body "Thanks, I'll reply shortly"
  outlook-cli draft list
  outlook-cli draft show <n>
  outlook-cli draft delete <n>

  outlook-cli mail search colgate
  outlook-cli mail unread
  outlook-cli mail read <n>

  outlook-cli contact search ross

  outlook-cli auth
"""

from __future__ import annotations

import argparse
import sys
from datetime import datetime
from html import escape
from pathlib import Path

from rich.console import Console
from rich.panel import Panel
from rich.table import Table

from . import config
from .cache import CAL_CACHE, MAIL_CACHE, TASK_CACHE, load_cache, save_cache
from .commands import calendar as calendar_commands
from .commands import contacts as contacts_commands
from .commands import mail as mail_commands
from .commands import teams as teams_commands
from .commands import tasks as tasks_commands
from .errors import OutlookAPIError, TokenExpiredError, TokenNotFoundError
from .outlook_client import OutlookClient
from .signatures import load_signature
from .token_manager import TokenManager

console = Console()


def _get_client(
    *,
    base_url: str = config.OUTLOOK_BASE_URL,
    token_domain: str = config.OUTLOOK_TOKEN_DOMAIN,
    token_label: str = "Outlook API",
) -> OutlookClient:
    """Build a client, handling auth errors."""
    tm = TokenManager(token_domain=token_domain, token_label=token_label)
    try:
        _ = tm.token
    except (TokenNotFoundError, TokenExpiredError):
        console.print(f"[yellow]{token_label} token missing or expired, running auth...[/]")
        if not tm.run_reauth():
            console.print("[red]Authentication failed. Run 'outlook-cli auth' to retry.[/]")
            sys.exit(1)
    return OutlookClient(tm, base_url=base_url)


def _get_graph_client() -> OutlookClient:
    """Build a Microsoft Graph client."""
    return _get_client(
        base_url=config.GRAPH_BASE_URL,
        token_domain=config.GRAPH_TOKEN_DOMAIN,
        token_label="Microsoft Graph",
    )


def _load_body(args: argparse.Namespace) -> str:
    """Load a body from --body or --body-file."""
    if args.body_file:
        body_path = Path(args.body_file)
        if not body_path.exists():
            console.print(f"[red]File not found: {body_path}[/]")
            sys.exit(1)
        return body_path.read_text(encoding="utf-8")
    if args.body:
        return args.body
    console.print("[red]Provide --body or --body-file[/]")
    sys.exit(1)


def _load_signature(path: Path) -> str:
    """Load a saved Outlook signature HTML file."""
    try:
        return load_signature(path)
    except FileNotFoundError as e:
        console.print(f"[red]{e}[/]")
        sys.exit(1)


def _text_to_html(text: str) -> str:
    """Convert plain text to simple HTML paragraphs."""
    text = text.replace("\r\n", "\n").strip()
    if not text:
        return "<p></p>"

    parts = []
    for block in text.split("\n\n"):
        block = block.strip()
        if not block:
            continue
        block_html = escape(block).replace("\n", "<br>\n")
        parts.append(f"<p>{block_html}</p>")
    return "".join(parts) or "<p></p>"


def _compose_email_html(
    body: str,
    *,
    is_html: bool,
    signature_html: str,
    quoted_html: str = "",
) -> str:
    """Build the final HTML body with Aptos styling and signature."""
    content_html = body if is_html else _text_to_html(body)
    style = (
        f"font-family: {config.EMAIL_FONT_FAMILY}; "
        f"font-size: {config.EMAIL_FONT_SIZE}; "
        "color: rgb(0, 0, 0);"
    )
    return f'<div style="{style}">{content_html}</div>{signature_html}{quoted_html}'


def cmd_create(args: argparse.Namespace) -> None:
    """Create a new email draft."""
    body = _load_body(args)
    to = [a.strip() for addr in args.to for a in addr.split(",") if a.strip()]
    cc = [a.strip() for addr in args.cc for a in addr.split(",") if a.strip()] if args.cc else None
    bcc = [a.strip() for addr in args.bcc for a in addr.split(",") if a.strip()] if args.bcc else None
    importance = args.importance or "Normal"
    html_body = _compose_email_html(
        body,
        is_html=args.html,
        signature_html=_load_signature(config.SIGNATURE_NEW_FILE),
    )

    client = _get_client()
    try:
        draft = client.create_draft(
            subject=args.subject,
            body=html_body,
            to=to,
            cc=cc,
            bcc=bcc,
            content_type="HTML",
            importance=importance,
        )
    except OutlookAPIError as e:
        console.print(f"[red]Failed to create draft: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    draft_id = draft.get("Id", "")
    recipients = ", ".join(r.get("EmailAddress", {}).get("Address", "") for r in draft.get("ToRecipients", []))

    console.print(Panel(
        f"[bold]Subject:[/] {draft.get('Subject', '')}\n"
        f"[bold]To:[/] {recipients}\n"
        f"[bold]Type:[/] HTML\n"
        f"[bold]ID:[/] [dim]{draft_id}[/]",
        title="[green]Draft created[/]",
        border_style="green",
    ))


def cmd_reply(args: argparse.Namespace) -> None:
    """Create a draft reply tied to an existing message."""
    body = _load_body(args)
    importance = args.importance or "Normal"
    client = _get_client()
    message_id = _resolve_message_id(client, args.message_id)

    try:
        draft = client.create_reply_draft(message_id, reply_all=args.reply_all)
        draft_id = draft.get("Id", "")
        if not draft_id:
            raise OutlookAPIError(0, "Reply draft was created without a draft ID")

        quoted_html = draft.get("Body", {}).get("Content", "")
        if draft.get("Body", {}).get("ContentType") == "Text":
            quoted_html = _text_to_html(quoted_html)

        reply_html = _compose_email_html(
            body,
            is_html=args.html,
            signature_html=_load_signature(config.SIGNATURE_REPLY_FILE),
            quoted_html=quoted_html,
        )
        draft = client.update_draft(
            draft_id,
            body=reply_html,
            content_type="HTML",
            importance=importance,
        )
    except OutlookAPIError as e:
        console.print(f"[red]Failed to create reply draft: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    recipients = ", ".join(r.get("EmailAddress", {}).get("Address", "") for r in draft.get("ToRecipients", []))
    console.print(Panel(
        f"[bold]Subject:[/] {draft.get('Subject', '')}\n"
        f"[bold]To:[/] {recipients}\n"
        f"[bold]Reply all:[/] {'Yes' if args.reply_all else 'No'}\n"
        f"[bold]Type:[/] HTML\n"
        f"[bold]ID:[/] [dim]{draft.get('Id', '')}[/]",
        title="[green]Reply draft created[/]",
        border_style="green",
    ))



def cmd_list(args: argparse.Namespace) -> None:
    """List drafts in the Drafts folder."""
    client = _get_client()
    try:
        drafts = client.list_drafts(top=args.count)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to list drafts: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not drafts:
        console.print("[dim]No drafts found.[/]")
        return

    table = Table(title=f"Drafts ({len(drafts)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Date", style="dim", width=12)
    table.add_column("To", width=30)
    table.add_column("Subject", ratio=1)
    table.add_column("ID", style="dim", width=20, no_wrap=True)

    for i, draft in enumerate(drafts, 1):
        dt = draft.get("CreatedDateTime", "")
        if dt:
            try:
                parsed = datetime.fromisoformat(dt.replace("Z", "+00:00"))
                dt = parsed.strftime("%b %d %H:%M")
            except Exception:
                dt = dt[:16]

        recipients = draft.get("ToRecipients", [])
        to_str = ", ".join(
            r.get("EmailAddress", {}).get("Address", "")
            for r in recipients[:3]
        )
        if len(recipients) > 3:
            to_str += f" +{len(recipients) - 3}"

        subject = draft.get("Subject", "(no subject)")
        draft_id = draft.get("Id", "")
        # Show a short ID suffix for usability
        short_id = draft_id[-16:] if len(draft_id) > 16 else draft_id

        table.add_row(str(i), dt, to_str, subject, short_id)

    console.print(table)
    console.print("[dim]Use the full ID from 'show' to send or delete a draft.[/]")


def cmd_show(args: argparse.Namespace) -> None:
    """Show a single draft."""
    client = _get_client()
    draft_id = _resolve_draft_id(client, args.draft_id)
    try:
        draft = client.get_draft(draft_id)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to get draft: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    recipients = ", ".join(
        r.get("EmailAddress", {}).get("Address", "")
        for r in draft.get("ToRecipients", [])
    )
    cc = ", ".join(
        r.get("EmailAddress", {}).get("Address", "")
        for r in draft.get("CcRecipients", [])
    )
    body_obj = draft.get("Body", {})
    body_preview = draft.get("BodyPreview", body_obj.get("Content", "")[:500])

    lines = [
        f"[bold]Subject:[/] {draft.get('Subject', '')}",
        f"[bold]To:[/] {recipients}",
    ]
    if cc:
        lines.append(f"[bold]CC:[/] {cc}")
    lines.append(f"[bold]Importance:[/] {draft.get('Importance', 'Normal')}")
    lines.append(f"[bold]Created:[/] {draft.get('CreatedDateTime', '')}")
    lines.append(f"[bold]Content type:[/] {body_obj.get('ContentType', '')}")
    lines.append(f"[bold]ID:[/] [dim]{draft.get('Id', '')}[/]")
    lines.append("")
    lines.append(body_preview)

    console.print(Panel("\n".join(lines), title="Draft", border_style="cyan"))


def cmd_delete(args: argparse.Namespace) -> None:
    """Delete a draft."""
    client = _get_client()
    draft_id = _resolve_draft_id(client, args.draft_id)
    try:
        client.delete_draft(draft_id)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to delete draft: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    console.print("[green]Draft deleted.[/]")


# ── Object cache (persisted to disk for cross-invocation read) ────────────


def _html_to_text(html: str) -> str:
    """Simple HTML to plain text conversion."""
    import re
    # Remove style and script blocks
    text = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<script[^>]*>.*?</script>", "", text, flags=re.DOTALL | re.IGNORECASE)
    # Convert <br> and block elements to newlines
    text = re.sub(r"<br\s*/?>\s*", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</(p|div|tr|li|h[1-6])>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<(p|div|tr|li|h[1-6])[^>]*>", "", text, flags=re.IGNORECASE)
    # Strip remaining tags
    text = re.sub(r"<[^>]+>", "", text)
    # Decode common HTML entities
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
    # Collapse whitespace
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _resolve_message_id(client: OutlookClient, ref: str) -> str:
    """Resolve a message reference to a full ID.

    Accepts a numeric index, cached ID suffix, or full ID.
    """
    cached = load_cache(MAIL_CACHE)

    if ref.isdigit():
        idx = int(ref)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print("[red]No search results cached. Run 'outlook-cli mail unread' or 'outlook-cli mail search <query>' first.[/]")
        else:
            console.print(f"[red]Index {idx} out of range. Only {len(cached)} results cached.[/]")
        sys.exit(1)

    if len(ref) < 40 and cached:
        matches = [item for item in cached if item["Id"].endswith(ref)]
        if len(matches) == 1:
            return matches[0]["Id"]
        if len(matches) > 1:
            console.print(f"[red]Ambiguous message ID suffix '{ref}'. Use a longer ID.[/]")
            sys.exit(1)

    return ref


def _format_datetime(value: str) -> str:
    """Format an ISO timestamp for terminal output."""
    if not value:
        return ""
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return parsed.strftime("%b %d %H:%M")
    except Exception:
        return value[:16]


# ── Tasks ─────────────────────────────────────────────────────────────────





def _resolve_task_id(client: OutlookClient, ref: str) -> str:
    if ref.isdigit():
        idx = int(ref)
        cached = load_cache(TASK_CACHE)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print("[red]No tasks cached. Run 'outlook-cli task list' first.[/]")
        else:
            console.print(f"[red]Index {idx} out of range.[/]")
        sys.exit(1)
    return ref






def _resolve_cal_id(client: OutlookClient, ref: str) -> str:
    if ref.isdigit():
        idx = int(ref)
        cached = load_cache(CAL_CACHE)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print("[red]No agenda cached. Run 'outlook-cli cal agenda' first.[/]")
        else:
            console.print(f"[red]Index {idx} out of range.[/]")
        sys.exit(1)
    return ref


def cmd_auth(args: argparse.Namespace) -> None:
    """Force re-authentication."""
    tm = TokenManager()
    if tm.run_reauth(headless=not args.headed):
        console.print(f"[green]Authenticated. Token expires in {tm.expires_in / 60:.0f} minutes.[/]")
    else:
        console.print("[red]Authentication failed.[/]")
        sys.exit(1)


def _resolve_draft_id(client: OutlookClient, ref: str) -> str:
    """Resolve a draft reference to a full ID.

    Accepts:
      - A full Outlook message ID (long string)
      - A numeric index like '1', '2' (matches position in draft list)
      - A partial ID suffix
    """
    # If it looks like a number, treat as list index
    if ref.isdigit():
        idx = int(ref)
        drafts = client.list_drafts(top=50)
        if 1 <= idx <= len(drafts):
            return drafts[idx - 1]["Id"]
        console.print(f"[red]Draft #{idx} not found. Only {len(drafts)} drafts available.[/]")
        sys.exit(1)

    # If it's short, try matching as suffix
    if len(ref) < 40:
        drafts = client.list_drafts(top=50)
        matches = [d for d in drafts if d["Id"].endswith(ref)]
        if len(matches) == 1:
            return matches[0]["Id"]
        if len(matches) > 1:
            console.print(f"[red]Ambiguous ID suffix '{ref}', {len(matches)} matches. Use a longer ID.[/]")
            sys.exit(1)
        console.print(f"[red]No draft found matching suffix '{ref}'.[/]")
        sys.exit(1)

    # Otherwise assume it's a full ID
    return ref


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="outlook-cli",
        description="Manage Outlook mail, drafts, calendar, contacts, and Teams chats",
    )
    sub = parser.add_subparsers(dest="domain", required=True)

    # ── Drafts ────────────────────────────────────────────────────────
    p_draft = sub.add_parser("draft", help="Manage email drafts")
    draft_sub = p_draft.add_subparsers(dest="command", required=True)

    cmd_draft_create = draft_sub.add_parser("create", help="Create a new draft")
    cmd_draft_create.add_argument("--to", required=True, action="append", help="Recipient (repeatable)")
    cmd_draft_create.add_argument("--cc", action="append", help="CC recipient")
    cmd_draft_create.add_argument("--bcc", action="append", help="BCC recipient")
    cmd_draft_create.add_argument("--subject", "-s", required=True, help="Subject line")
    cmd_draft_create.add_argument("--body", "-b", help="Body text (inline)")
    cmd_draft_create.add_argument("--body-file", "-f", help="Read body from a file (.txt or .html)")
    cmd_draft_create.add_argument("--html", action="store_true", help="Treat body input as HTML")
    cmd_draft_create.add_argument("--importance", choices=["Low", "Normal", "High"], default="Normal")
    cmd_draft_create.set_defaults(func=cmd_create)

    cmd_draft_reply = draft_sub.add_parser("reply", help="Create a reply draft from an email")
    cmd_draft_reply.add_argument("message_id", help="Message index, cached ID suffix, or full ID")
    cmd_draft_reply.add_argument("--body", "-b", help="Reply body text (inline)")
    cmd_draft_reply.add_argument("--body-file", "-f", help="Read reply body from a file (.txt or .html)")
    cmd_draft_reply.add_argument("--html", action="store_true", help="Treat body input as HTML")
    cmd_draft_reply.add_argument("--reply-all", action="store_true", help="Reply to all recipients")
    cmd_draft_reply.add_argument("--importance", choices=["Low", "Normal", "High"], default="Normal")
    cmd_draft_reply.set_defaults(func=cmd_reply)

    cmd_draft_list = draft_sub.add_parser("list", help="List drafts")
    cmd_draft_list.add_argument("--count", "-n", type=int, default=20, help="Number of drafts")
    cmd_draft_list.set_defaults(func=cmd_list)

    cmd_draft_show = draft_sub.add_parser("show", help="Show a draft")
    cmd_draft_show.add_argument("draft_id", help="Draft ID, index (#1, #2..), or ID suffix")
    cmd_draft_show.set_defaults(func=cmd_show)

    cmd_draft_del = draft_sub.add_parser("delete", help="Delete a draft")
    cmd_draft_del.add_argument("draft_id", help="Draft ID, index, or ID suffix")
    cmd_draft_del.set_defaults(func=cmd_delete)

    mail_ctx = mail_commands.build_ctx(
        console=console,
        get_client=_get_client,
        html_to_text=_html_to_text,
        format_datetime=_format_datetime,
        resolve_message_id=_resolve_message_id,
    )
    contacts_ctx = contacts_commands.build_ctx(
        console=console,
        get_client=_get_client,
    )

    # ── Mail ──────────────────────────────────────────────────────────
    p_mail = sub.add_parser("mail", help="Read and search emails")
    mail_sub = p_mail.add_subparsers(dest="command", required=True)

    cmd_mail_unread = mail_sub.add_parser("unread", help="List unread emails from the Inbox")
    cmd_mail_unread.add_argument("--count", "-n", type=int, default=50, help="Max results")
    cmd_mail_unread.set_defaults(func=mail_commands.cmd_unread, _mail_ctx=mail_ctx)

    cmd_mail_search = mail_sub.add_parser("search", help="Search emails by keyword")
    cmd_mail_search.add_argument("query", help="Search term")
    cmd_mail_search.add_argument("--count", "-n", type=int, default=20, help="Max results")
    cmd_mail_search.set_defaults(func=mail_commands.cmd_mail, _mail_ctx=mail_ctx)

    cmd_mail_read = mail_sub.add_parser("read", help="Read an email by index or ID")
    cmd_mail_read.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_read.set_defaults(func=mail_commands.cmd_read, _mail_ctx=mail_ctx)

    tasks_ctx = tasks_commands.build_ctx(
        console=console,
        get_client=_get_client,
        resolve_task_id=_resolve_task_id,
    )

    # ── Tasks ─────────────────────────────────────────────────────────
    p_task = sub.add_parser("task", help="Manage tasks")
    task_sub = p_task.add_subparsers(dest="command", required=True)

    cmd_task_l = task_sub.add_parser("list", help="List active tasks")
    cmd_task_l.add_argument("--count", "-n", type=int, default=50, help="Max tasks")
    cmd_task_l.set_defaults(func=tasks_commands.cmd_task_list, _tasks_ctx=tasks_ctx)

    cmd_task_c = task_sub.add_parser("create", help="Create a task")
    cmd_task_c.add_argument("subject", help="Task description")
    cmd_task_c.set_defaults(func=tasks_commands.cmd_task_create, _tasks_ctx=tasks_ctx)

    cmd_task_complete_cmd = task_sub.add_parser("complete", help="Mark task complete")
    cmd_task_complete_cmd.add_argument("task_id", help="Task index or full ID")
    cmd_task_complete_cmd.set_defaults(func=tasks_commands.cmd_task_complete, _tasks_ctx=tasks_ctx)

    cmd_task_delete_cmd = task_sub.add_parser("delete", help="Delete a task")
    cmd_task_delete_cmd.add_argument("task_id", help="Task index or full ID")
    cmd_task_delete_cmd.set_defaults(func=tasks_commands.cmd_task_delete, _tasks_ctx=tasks_ctx)

    calendar_ctx = calendar_commands.build_ctx(
        console=console,
        get_client=_get_client,
        html_to_text=_html_to_text,
        format_datetime=_format_datetime,
        resolve_cal_id=_resolve_cal_id,
    )

    # ── Calendar ──────────────────────────────────────────────────────
    p_cal = sub.add_parser("cal", help="Manage calendar events")
    cal_sub = p_cal.add_subparsers(dest="command", required=True)

    cmd_cal_agenda = cal_sub.add_parser("agenda", help="List upcoming calendar events")
    cmd_cal_agenda.add_argument("--days", "-d", type=int, default=7, help="Days to look ahead")
    cmd_cal_agenda.add_argument("--count", "-n", type=int, default=20, help="Max events")
    cmd_cal_agenda.set_defaults(func=calendar_commands.cmd_cal_agenda, _calendar_ctx=calendar_ctx)

    cmd_cal_show = cal_sub.add_parser("show", help="Show event details")
    cmd_cal_show.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_show.set_defaults(func=calendar_commands.cmd_cal_show, _calendar_ctx=calendar_ctx)

    cmd_cal_create_cmd = cal_sub.add_parser("create", help="Create an event")
    cmd_cal_create_cmd.add_argument("subject", help="Event subject")
    cmd_cal_create_cmd.add_argument("start", help="Start time (e.g. '2026-04-10 14:00')")
    cmd_cal_create_cmd.add_argument("end", help="End time (e.g. '2026-04-10 15:00')")
    cmd_cal_create_cmd.add_argument("--location", "-l", help="Location")
    cmd_cal_create_cmd.add_argument("--body", "-b", help="Event description")
    cmd_cal_create_cmd.add_argument("--attendee", dest="attendees", action="append", help="Attendee email (repeatable)")
    cmd_cal_create_cmd.set_defaults(func=calendar_commands.cmd_cal_create, _calendar_ctx=calendar_ctx)

    cmd_cal_delete_cmd = cal_sub.add_parser("delete", help="Delete an event")
    cmd_cal_delete_cmd.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_delete_cmd.set_defaults(func=calendar_commands.cmd_cal_delete, _calendar_ctx=calendar_ctx)

    cmd_cal_accept_cmd = cal_sub.add_parser("accept", help="Accept an event invitation")
    cmd_cal_accept_cmd.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_accept_cmd.add_argument("--comment", "-m", help="Optional response comment")
    cmd_cal_accept_cmd.add_argument(
        "--no-send-response",
        dest="send_response",
        action="store_false",
        help="Do not send a response email to the organizer",
    )
    cmd_cal_accept_cmd.set_defaults(
        func=calendar_commands.cmd_cal_accept,
        _calendar_ctx=calendar_ctx,
        send_response=True,
    )

    cmd_cal_decline_cmd = cal_sub.add_parser("decline", help="Decline an event invitation")
    cmd_cal_decline_cmd.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_decline_cmd.add_argument("--comment", "-m", help="Optional response comment")
    cmd_cal_decline_cmd.add_argument(
        "--no-send-response",
        dest="send_response",
        action="store_false",
        help="Do not send a response email to the organizer",
    )
    cmd_cal_decline_cmd.set_defaults(
        func=calendar_commands.cmd_cal_decline,
        _calendar_ctx=calendar_ctx,
        send_response=True,
    )

    cmd_cal_tentative_cmd = cal_sub.add_parser("tentative", help="Tentatively accept an event invitation")
    cmd_cal_tentative_cmd.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_tentative_cmd.add_argument("--comment", "-m", help="Optional response comment")
    cmd_cal_tentative_cmd.add_argument(
        "--no-send-response",
        dest="send_response",
        action="store_false",
        help="Do not send a response email to the organizer",
    )
    cmd_cal_tentative_cmd.set_defaults(
        func=calendar_commands.cmd_cal_tentative,
        _calendar_ctx=calendar_ctx,
        send_response=True,
    )

    cmd_cal_cancel_cmd = cal_sub.add_parser("cancel", help="Cancel an event you organize")
    cmd_cal_cancel_cmd.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_cancel_cmd.add_argument("--comment", "-m", help="Optional cancellation comment")
    cmd_cal_cancel_cmd.set_defaults(func=calendar_commands.cmd_cal_cancel, _calendar_ctx=calendar_ctx)

    # ── Contacts ──────────────────────────────────────────────────────
    p_contact = sub.add_parser("contact", help="Manage contacts")
    contact_sub = p_contact.add_subparsers(dest="command", required=True)

    cmd_contact_search = contact_sub.add_parser("search", help="Search contacts by name or email")
    cmd_contact_search.add_argument("query", help="Name or email to search for")
    cmd_contact_search.add_argument("--count", "-n", type=int, default=10, help="Max results")
    cmd_contact_search.set_defaults(func=contacts_commands.cmd_contacts, _contacts_ctx=contacts_ctx)

    # ── Teams ─────────────────────────────────────────────────────────
    p_teams = sub.add_parser("teams", help="Browse Teams chats and messages")
    teams_sub = p_teams.add_subparsers(dest="command", required=True)

    teams_ctx = teams_commands.build_ctx(
        console=console,
        get_graph_client=_get_graph_client,
        format_datetime=_format_datetime,
        html_to_text=_html_to_text,
    )

    cmd_teams_list_cmd = teams_sub.add_parser("list", help="List Teams chats")
    cmd_teams_list_cmd.add_argument("--count", "-n", type=int, default=20, help="Max chats")
    cmd_teams_list_cmd.set_defaults(func=teams_commands.cmd_teams_list, _teams_ctx=teams_ctx)

    cmd_teams_show_cmd = teams_sub.add_parser("show", help="Show Teams chat details")
    cmd_teams_show_cmd.add_argument("chat_id", help="Chat index, cached ID suffix, or full ID")
    cmd_teams_show_cmd.set_defaults(func=teams_commands.cmd_teams_show, _teams_ctx=teams_ctx)

    cmd_teams_messages_cmd = teams_sub.add_parser("messages", help="Read Teams chat messages")
    cmd_teams_messages_cmd.add_argument("chat_id", help="Chat index, cached ID suffix, or full ID")
    cmd_teams_messages_cmd.add_argument("--count", "-n", type=int, default=20, help="Max messages")
    cmd_teams_messages_cmd.set_defaults(func=teams_commands.cmd_teams_messages, _teams_ctx=teams_ctx)

    # ── Auth ──────────────────────────────────────────────────────────
    p_auth = sub.add_parser("auth", help="Force re-authentication")
    p_auth.add_argument(
        "--headed",
        action="store_true",
        help="Open a visible browser instead of running fully headless",
    )

    args = parser.parse_args()

    if args.domain == "auth":
        cmd_auth(args)
    elif hasattr(args, "func"):
        args.func(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
