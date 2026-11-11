from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

from rich.panel import Panel
from rich.table import Table

from .. import config, runtime
from ..errors import CliError, OutlookAPIError
from ..formatting import compose_email_html as _compose_email_html
from ..formatting import parse_recipients as _parse_recipients
from ..mail_compose import attach_to_draft, draft_update_fields, encode_attachments, finish_quoted_draft
from ..outlook_client import OutlookClient
from ..parser import _add_yes_arg, add_output_args
from ..progress import spinner
from ..references import _resolve_message_id
from ..runtime import console
from ..signatures import load_signature

log = logging.getLogger(__name__)


def _load_body(args: argparse.Namespace) -> str:
    """Load a body from --body or --body-file."""
    if args.body_file:
        body_path = Path(args.body_file)
        try:
            return body_path.read_text(encoding="utf-8")
        except (OSError, UnicodeError) as exc:
            raise CliError(f"Cannot read body file {body_path}: {exc}") from exc
    if args.body is not None:
        return args.body
    raise CliError("[red]Provide --body or --body-file[/]")

def _load_signature(path: Path) -> str:
    """Load a saved Outlook signature HTML file."""
    try:
        return load_signature(path)
    except FileNotFoundError as e:
        raise CliError(f"[red]{e}[/]") from e

def cmd_create(args: argparse.Namespace) -> None:
    """Create a new email draft."""
    body = _load_body(args)
    encoded = _encode_files(args)
    to = [a.strip() for addr in args.to for a in addr.split(",") if a.strip()]
    cc = [a.strip() for addr in args.cc for a in addr.split(",") if a.strip()] if args.cc else None
    bcc = [a.strip() for addr in args.bcc for a in addr.split(",") if a.strip()] if args.bcc else None
    importance = args.importance or "Normal"
    html_body = _compose_email_html(
        body,
        is_html=args.html,
        signature_html=_load_signature(config.SIGNATURE_NEW_FILE),
    )

    client = runtime.get_client()
    try:
        with spinner(args, "Creating draft..."):
            draft = client.create_draft(
                subject=args.subject,
                body=html_body,
                to=to,
                cc=cc,
                bcc=bcc,
                content_type="HTML",
                importance=importance,
            )
            attach_to_draft(client, draft, encoded)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to create draft: {e}[/]") from e
    finally:
        client.close()

    if getattr(args, "json", False):
        runtime.print_json(draft)
        return
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
    encoded = _encode_files(args)
    signature = _load_signature(config.SIGNATURE_REPLY_FILE)
    importance = args.importance or "Normal"
    client = runtime.get_client()
    try:
        message_id = _resolve_message_id(client, args.message_id)
        with spinner(args, "Creating reply draft..."):
            draft = client.create_reply_draft(message_id, reply_all=args.reply_all)
            draft = finish_quoted_draft(
                client, draft, body, args.html, signature, importance=importance,
            )
            attach_to_draft(client, draft, encoded)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to create reply draft: {e}[/]") from e
    finally:
        client.close()

    if getattr(args, "json", False):
        runtime.print_json(draft)
        return
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

def _encode_files(args: argparse.Namespace) -> list[dict]:
    try:
        return encode_attachments(getattr(args, "attach", None))
    except ValueError as exc:
        raise CliError(str(exc)) from exc


def _show_saved(args, draft: dict, title: str) -> None:
    if getattr(args, "json", False):
        runtime.print_json(draft)
    else:
        console.print(Panel(f"[bold]ID:[/] {draft.get('Id', '')}", title=title))


def cmd_update(args: argparse.Namespace) -> None:
    """Edit supplied fields without changing an omitted body or signature."""
    encoded = _encode_files(args)
    body = _load_body(args) if args.body is not None or args.body_file is not None else None
    try:
        fields = draft_update_fields(
            args.subject, body, args.to, args.cc, args.bcc,
            args.html, args.importance, args.signature,
        )
    except (ValueError, FileNotFoundError) as exc:
        raise CliError(str(exc)) from exc
    client = runtime.get_client()
    try:
        draft_id = _resolve_draft_id(client, args.draft_id)
        with spinner(args, "Updating draft..."):
            draft = client.update_mail_draft(draft_id, **fields)
            attach_to_draft(client, draft, encoded)
    except (OutlookAPIError, ValueError) as exc:
        raise CliError(f"Failed to update draft: {exc}") from exc
    finally:
        client.close()
    _show_saved(args, draft, "Draft updated")


def cmd_forward(args: argparse.Namespace) -> None:
    """Create a forward draft with a signature and the quoted original."""
    encoded = _encode_files(args)
    body = _load_body(args) if args.body is not None or args.body_file is not None else ""
    signature = _load_signature(config.SIGNATURE_REPLY_FILE)
    recipients = _parse_recipients(args.to)
    if not recipients:
        raise CliError("At least one recipient is required in --to")
    client = runtime.get_client()
    try:
        message_id = _resolve_message_id(client, args.message_id)
        with spinner(args, "Creating forward draft..."):
            draft = client.create_forward_draft(message_id)
            draft = finish_quoted_draft(
                client, draft, body, args.html, signature, to=recipients,
                cc=_parse_recipients(args.cc) or None,
                bcc=_parse_recipients(args.bcc) or None, importance=args.importance,
            )
            attach_to_draft(client, draft, encoded)
    except (OutlookAPIError, ValueError) as exc:
        raise CliError(f"Failed to create forward draft: {exc}") from exc
    finally:
        client.close()
    _show_saved(args, draft, "Forward draft created")


def _register_edits(draft_sub) -> None:
    update = draft_sub.add_parser("update", help="Edit an existing draft")
    update.add_argument("draft_id", help="Draft ID, index, or ID suffix")
    update.add_argument("--subject", "-s", help="Replace the subject; empty clears it")
    update.add_argument("--signature", choices=["none", "new", "reply"], default="none",
                        help="Signature for a replacement body only (default: none)")
    _add_edit_options(update)
    update.set_defaults(func=cmd_update)
    forward = draft_sub.add_parser("forward", help="Create a forward draft")
    forward.add_argument("message_id", help="Message index, cached ID suffix, or full ID")
    _add_edit_options(forward, require_to=True)
    forward.set_defaults(func=cmd_forward)


def _add_edit_options(parser, *, require_to=False) -> None:
    parser.add_argument("--to", action="append", required=require_to, help="Recipient (repeatable)")
    parser.add_argument("--cc", action="append", help="CC recipient (repeatable)")
    parser.add_argument("--bcc", action="append", help="BCC recipient (repeatable)")
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--body", "-b", help="Replacement body text; empty clears it")
    group.add_argument("--body-file", "-f", help="Read the body from a local file")
    parser.add_argument("--html", action="store_true", help="Treat body input as HTML")
    parser.add_argument("--importance", choices=["Low", "Normal", "High"])
    parser.add_argument("--attach", action="append", help="Attach a local file (repeatable, max 3 MiB each)")
    add_output_args(parser)


def cmd_list(args: argparse.Namespace) -> None:
    """List drafts in the Drafts folder."""
    client = runtime.get_client()
    try:
        with spinner(args, "Loading drafts..."):
            drafts = client.list_drafts(top=args.count)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to list drafts: {e}[/]") from e
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
                log.debug("Best-effort operation failed", exc_info=True)
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
    client = runtime.get_client()
    draft_id = _resolve_draft_id(client, args.draft_id)
    try:
        with spinner(args, "Loading draft..."):
            draft = client.get_draft(draft_id)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to get draft: {e}[/]") from e
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
    """Delete a draft, never a received email."""
    client = runtime.get_client()
    try:
        draft_id = _resolve_draft_id(client, args.draft_id)
        if not client.get_draft(draft_id).get("IsDraft", False):
            raise CliError("Refusing to delete: this message is not a draft.")
        with spinner(args, "Deleting draft..."):
            client.delete_draft(draft_id)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to delete draft: {e}[/]") from e
    finally:
        client.close()

    console.print("[green]Draft deleted.[/]")

def _require_send_confirmation(args: argparse.Namespace, description: str) -> None:
    """Exit without sending unless --yes was passed."""
    if getattr(args, "yes", False):
        return
    console.print(f"[yellow]Not sent.[/] Would send: {description}")
    console.print("[dim]Re-run with --yes to send.[/]")
    sys.exit(2)

def cmd_send_draft(args: argparse.Namespace) -> None:
    """Send an existing draft message."""
    _require_send_confirmation(args, f"draft {args.draft_id}")
    client = runtime.get_client()
    draft_id = _resolve_draft_id(client, args.draft_id)
    try:
        client.send_message(draft_id)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to send draft: {e}[/]") from e
    finally:
        client.close()

    console.print("[green]Draft sent.[/]")

def cmd_send_mail(args: argparse.Namespace) -> None:
    """Send a new email immediately."""
    body = _load_body(args)
    encoded = _encode_files(args)
    to = _parse_recipients(args.to)
    _require_send_confirmation(args, f"email '{args.subject}' to {', '.join(to)}")
    cc = _parse_recipients(args.cc)
    bcc = _parse_recipients(args.bcc)
    if args.signature != "none":
        signature_file = config.SIGNATURE_REPLY_FILE if args.signature == "reply" else config.SIGNATURE_NEW_FILE
        body = _compose_email_html(
            body,
            is_html=args.html,
            signature_html=_load_signature(signature_file),
        )
        content_type = "HTML"
    else:
        content_type = "HTML" if args.html else "Text"

    client = runtime.get_client()
    try:
        client.send_mail(
            subject=args.subject,
            body=body,
            to=to,
            cc=cc or None,
            bcc=bcc or None,
            content_type=content_type,
            save_to_sent_items=not args.no_save,
            importance=args.importance,
            **({"attachments": encoded} if getattr(args, "attach", None) else {}),
        )
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to send email: {e}[/]") from e
    finally:
        client.close()

    console.print(f"[green]Email sent:[/] {args.subject}")


# ── Object cache (persisted to disk for cross-invocation read) ────────────

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
        raise CliError(f"[red]Draft #{idx} not found. Only {len(drafts)} drafts available.[/]")

    # If it's short, try matching as suffix
    if len(ref) < 40:
        drafts = client.list_drafts(top=50)
        matches = [d for d in drafts if d["Id"].endswith(ref)]
        if len(matches) == 1:
            return matches[0]["Id"]
        if len(matches) > 1:
            raise CliError(f"[red]Ambiguous ID suffix '{ref}', {len(matches)} matches. Use a longer ID.[/]")
        raise CliError(f"[red]No draft found matching suffix '{ref}'.[/]")

    # Otherwise assume it's a full ID
    return ref


def register(sub) -> None:
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
    cmd_draft_create.add_argument("--attach", action="append", help="Attach a local file (repeatable, max 3 MiB each)")
    add_output_args(cmd_draft_create)
    cmd_draft_create.set_defaults(func=cmd_create)

    cmd_draft_reply = draft_sub.add_parser("reply", help="Create a reply draft from an email")
    cmd_draft_reply.add_argument("message_id", help="Message index, cached ID suffix, or full ID")
    cmd_draft_reply.add_argument("--body", "-b", help="Reply body text (inline)")
    cmd_draft_reply.add_argument("--body-file", "-f", help="Read reply body from a file (.txt or .html)")
    cmd_draft_reply.add_argument("--html", action="store_true", help="Treat body input as HTML")
    cmd_draft_reply.add_argument("--reply-all", action="store_true", help="Reply to all recipients")
    cmd_draft_reply.add_argument("--importance", choices=["Low", "Normal", "High"], default="Normal")
    cmd_draft_reply.add_argument("--attach", action="append", help="Attach a local file (repeatable, max 3 MiB each)")
    add_output_args(cmd_draft_reply)
    cmd_draft_reply.set_defaults(func=cmd_reply)
    _register_edits(draft_sub)

    cmd_draft_list = draft_sub.add_parser("list", help="List drafts")
    cmd_draft_list.add_argument("--count", "-n", type=int, default=20, help="Number of drafts")
    cmd_draft_list.set_defaults(func=cmd_list)

    cmd_draft_show = draft_sub.add_parser("show", help="Show a draft")
    cmd_draft_show.add_argument("draft_id", help="Draft ID, index (#1, #2..), or ID suffix")
    cmd_draft_show.set_defaults(func=cmd_show)

    cmd_draft_del = draft_sub.add_parser("delete", help="Delete a draft")
    cmd_draft_del.add_argument("draft_id", help="Draft ID, index, or ID suffix")
    cmd_draft_del.set_defaults(func=cmd_delete)

    cmd_draft_send = draft_sub.add_parser("send", help="Send an existing draft (requires --yes)")
    cmd_draft_send.add_argument("draft_id", help="Draft ID, index, or ID suffix")
    _add_yes_arg(cmd_draft_send)
    cmd_draft_send.set_defaults(func=cmd_send_draft)
