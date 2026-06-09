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
import base64
import json
import sys
import time
from datetime import datetime
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError
from html import escape
from pathlib import Path

from rich.console import Console
from rich import box
from rich.panel import Panel
from rich.table import Table

from . import config
from .cache import CAL_CACHE, CONTACT_CACHE, MAIL_CACHE, TASK_CACHE, load_cache
from .commands import calendar as calendar_commands
from .commands import contacts as contacts_commands
from .commands import mail as mail_commands
from .commands import teams as teams_commands
from .commands import tasks as tasks_commands
from .commands.signature import cmd_signature_fetch
from .commands.gateway import cmd_gateway_start, cmd_gateway_stop, cmd_gateway_status
from .commands.files import (
    cmd_files_sites,
    cmd_files_list,
    cmd_files_upload,
    cmd_files_download,
    cmd_files_mkdir,
    cmd_files_rename,
    cmd_files_move,
)
from .errors import OutlookAPIError, TokenExpiredError, TokenNotFoundError
from .outlook_client import OutlookClient
from .progress import spinner
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


def _token_available(domain: str) -> bool:
    if not config.TOKENS_FILE.exists():
        return False
    try:
        data = json.loads(config.TOKENS_FILE.read_text())
        token = data.get("tokens", {}).get(domain)
        if not token:
            return False
        claims = _decode_token_claims(token)
        return int(float(claims.get("exp", 0)) - time.time()) > 0
    except Exception:
        return False


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
        with spinner(args, "Creating reply draft..."):
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
        with spinner(args, "Loading drafts..."):
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
        with spinner(args, "Loading draft..."):
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
        with spinner(args, "Deleting draft..."):
            client.delete_draft(draft_id)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to delete draft: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    console.print("[green]Draft deleted.[/]")


def cmd_send_draft(args: argparse.Namespace) -> None:
    """Send an existing draft message."""
    client = _get_client()
    draft_id = _resolve_draft_id(client, args.draft_id)
    try:
        client.send_message(draft_id)
    except OutlookAPIError as e:
        console.print(f"[red]Failed to send draft: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    console.print("[green]Draft sent.[/]")


def _parse_recipients(values: list[str] | None) -> list[str]:
    if not values:
        return []
    return [email.strip() for value in values for email in value.split(",") if email.strip()]


def cmd_send_mail(args: argparse.Namespace) -> None:
    """Send a new email immediately."""
    body = _load_body(args)
    to = _parse_recipients(args.to)
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

    client = _get_client()
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
        )
    except OutlookAPIError as e:
        console.print(f"[red]Failed to send email: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    console.print(f"[green]Email sent:[/] {args.subject}")


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





def _resolve_contact_id(ref: str) -> str:
    if ref.isdigit():
        idx = int(ref)
        cached = load_cache(CONTACT_CACHE)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print("[red]No contacts cached. Run 'outlook-cli contact search <query>' first.[/]")
        else:
            console.print(f"[red]Index {idx} out of range.[/]")
        sys.exit(1)
    return ref


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


def _decode_token_claims(token: str) -> dict[str, object]:
    payload = token.split(".")[1]
    payload += "=" * (-len(payload) % 4)
    return json.loads(base64.urlsafe_b64decode(payload))


def _token_status(domain: str, label: str) -> dict[str, object]:
    scopes: list[str] = []
    audience = ""
    expires_in = -1
    if config.TOKENS_FILE.exists():
        try:
            data = json.loads(config.TOKENS_FILE.read_text())
            token = data.get("tokens", {}).get(domain)
            if token:
                claims = _decode_token_claims(token)
                scopes = str(claims.get("scp", "")).split()
                audience = str(claims.get("aud", ""))
                expires_in = int(float(claims.get("exp", 0)) - time.time())
        except Exception:
            pass
    present = expires_in != -1
    expired = expires_in <= 0
    return {
        "label": label,
        "domain": domain,
        "present": present,
        "expired": expired,
        "expiresInSeconds": expires_in,
        "audience": audience,
        "scopes": scopes,
        "notesReadWrite": "Notes.ReadWrite" in scopes,
    }


def cmd_auth(args: argparse.Namespace) -> None:
    """Manage authentication."""
    command = args.auth_command or "login"
    if command == "status":
        statuses = [
            _token_status(config.OUTLOOK_TOKEN_DOMAIN, "Outlook API"),
            _token_status(config.GRAPH_TOKEN_DOMAIN, "Microsoft Graph"),
            _token_status("substrate.office.com", "Substrate"),
        ]
        notes_scope_present = any(
            bool(item.get("notesReadWrite")) and bool(item.get("present")) and not bool(item.get("expired"))
            for item in statuses
        )
        capabilities = {
            "notesReadWriteAnyToken": notes_scope_present,
            "notesTokenCaptured": notes_scope_present,
        }
        if args.json:
            print(json.dumps({"tokens": statuses, "capabilities": capabilities}, indent=2))
            return
        table = Table(title="Auth status", box=box.SIMPLE)
        table.add_column("Token")
        table.add_column("Domain")
        table.add_column("Status")
        table.add_column("Expires")
        table.add_column("Notes")
        for item in statuses:
            if not item["present"]:
                status = "[red]missing[/]"
                expires = ""
            elif item["expired"]:
                status = "[red]expired[/]"
                expires = "expired"
            else:
                status = "[green]valid[/]"
                expires = f"{int(item['expiresInSeconds']) // 60} min"
            notes = "[green]yes[/]" if item.get("notesReadWrite") else ""
            table.add_row(str(item["label"]), str(item["domain"]), status, expires, notes)
        console.print(table)
        console.print(f"OneNote scope present: {'yes' if notes_scope_present else 'no'}")
        return
    if command == "clear":
        if config.TOKENS_FILE.exists():
            config.TOKENS_FILE.unlink()
            console.print(f"[green]Deleted token cache:[/] {config.TOKENS_FILE}")
        else:
            console.print("[dim]No token cache found.[/]")
        return
    if command == "scopes":
        if not config.TOKENS_FILE.exists():
            console.print("[red]No token cache found. Run 'outlook-cli auth' first.[/]")
            sys.exit(1)
        data = json.loads(config.TOKENS_FILE.read_text())
        output = []
        for domain, token in data.get("tokens", {}).items():
            claims = _decode_token_claims(token)
            scopes = str(claims.get("scp", "")).split()
            output.append({
                "domain": domain,
                "audience": claims.get("aud", ""),
                "tenant": claims.get("tid", ""),
                "appId": claims.get("appid", ""),
                "expiresInSeconds": int(float(claims.get("exp", 0)) - time.time()),
                "scopes": scopes,
                "roles": claims.get("roles", []),
            })
        if args.json:
            print(json.dumps(output, indent=2))
            return
        for item in output:
            console.print(f"[bold]{item['domain']}[/] ({item['audience']})")
            for scope in item["scopes"]:
                console.print(f"  - {scope}")
        return

    tm = TokenManager()
    if tm.run_reauth(headless=not args.headed):
        console.print(f"[green]Authenticated. Token expires in {tm.expires_in / 60:.0f} minutes.[/]")
    else:
        console.print("[red]Authentication failed.[/]")
        sys.exit(1)


def _config_check_items() -> list[dict[str, object]]:
    items: list[dict[str, object]] = []

    def add(name: str, ok: bool, detail: str = "") -> None:
        items.append({"name": name, "ok": ok, "detail": detail})

    add("MS_EMAIL", bool(config.MS_EMAIL), "set" if config.MS_EMAIL else "missing")
    add("MS_PASSWORD", bool(config.MS_PASSWORD), "set" if config.MS_PASSWORD else "missing")
    add("config file", config.CONFIG_FILE.exists(), str(config.CONFIG_FILE))
    try:
        ZoneInfo(config.LOCAL_TIMEZONE)
        add("LOCAL_TIMEZONE", True, config.LOCAL_TIMEZONE)
    except ZoneInfoNotFoundError:
        add("LOCAL_TIMEZONE", False, f"invalid: {config.LOCAL_TIMEZONE}")
    add("OUTLOOK_TIMEZONE", bool(config.OUTLOOK_TIMEZONE), config.OUTLOOK_TIMEZONE or "missing")
    add("SIGNATURE_NEW_FILE", config.SIGNATURE_NEW_FILE.exists(), str(config.SIGNATURE_NEW_FILE))
    add("SIGNATURE_REPLY_FILE", config.SIGNATURE_REPLY_FILE.exists(), str(config.SIGNATURE_REPLY_FILE))
    try:
        config.SESSION_DIR.mkdir(parents=True, exist_ok=True)
        probe = config.SESSION_DIR / ".write-test"
        probe.write_text("ok")
        probe.unlink()
        add("session_state writable", True, str(config.SESSION_DIR))
    except Exception as e:
        add("session_state writable", False, str(e))
    try:
        import playwright.sync_api  # noqa: F401
        add("playwright", True, "installed")
    except Exception as e:
        add("playwright", False, str(e))
    return items


def cmd_config_check(args: argparse.Namespace) -> None:
    """Validate local CLI configuration without printing secrets."""
    items = _config_check_items()
    if args.json:
        print(json.dumps(items, indent=2))
    else:
        table = Table(title="Config check", box=box.SIMPLE)
        table.add_column("Item")
        table.add_column("Status")
        table.add_column("Detail")
        for item in items:
            table.add_row(
                str(item["name"]),
                "[green]ok[/]" if item["ok"] else "[red]fail[/]",
                str(item["detail"]),
            )
        console.print(table)
    if not all(bool(item["ok"]) for item in items):
        sys.exit(1)


def cmd_mailbox_show(args: argparse.Namespace) -> None:
    client = _get_client()
    try:
        with spinner(args, "Loading mailbox settings..."):
            settings = client.get_mailbox_settings()
    except OutlookAPIError as e:
        console.print(f"[red]Failed to get mailbox settings: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
    if args.json:
        print(json.dumps(settings, indent=2))
        return
    table = Table(title="Mailbox settings", box=box.SIMPLE)
    table.add_column("Setting")
    table.add_column("Value")
    for key in ("TimeZone", "Language", "DateFormat", "TimeFormat"):
        if key in settings:
            table.add_row(key, str(settings[key]))
    auto = settings.get("AutomaticRepliesSetting") or {}
    if auto:
        table.add_row("Automatic replies", str(auto.get("Status", "")))
    console.print(table)


def cmd_mailbox_update(args: argparse.Namespace) -> None:
    if not any([args.timezone, args.auto_reply_status, args.internal_reply, args.external_reply]):
        console.print("[red]Provide at least one mailbox setting to update.[/]")
        sys.exit(1)
    automatic_replies = None
    if any([args.auto_reply_status, args.internal_reply, args.external_reply]):
        automatic_replies = {}
        if args.auto_reply_status:
            automatic_replies["Status"] = args.auto_reply_status
        if args.internal_reply is not None:
            automatic_replies["InternalReplyMessage"] = args.internal_reply
        if args.external_reply is not None:
            automatic_replies["ExternalReplyMessage"] = args.external_reply
    client = _get_client()
    try:
        with spinner(args, "Updating mailbox settings..."):
            client.update_mailbox_settings(
                time_zone=args.timezone,
                automatic_replies=automatic_replies,
            )
        console.print("[green]Mailbox settings updated.[/]")
    except OutlookAPIError as e:
        console.print(f"[red]Failed to update mailbox settings: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


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


def add_output_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--json",
        dest="json",
        action="store_true",
        default=argparse.SUPPRESS,
        help="Output machine-readable JSON",
    )
    parser.add_argument(
        "--table",
        dest="json",
        action="store_false",
        default=argparse.SUPPRESS,
        help="Output human-readable tables (default)",
    )


def main() -> None:
    parser = argparse.ArgumentParser(
        prog="outlook-cli",
        description="Manage Outlook mail, drafts, calendar, contacts, and Teams chats",
    )
    parser.set_defaults(json=False)
    add_output_args(parser)
    parser.add_argument(
        "--no-spinner",
        action="store_true",
        help="Disable interactive progress spinners",
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

    # Sending existing drafts is intentionally not exposed to agents.
    # Implementation is preserved in cmd_send_draft / OutlookClient.send_message,
    # but the parser is disabled for safety.

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
        resolve_contact_id=_resolve_contact_id,
    )

    # ── Mail ──────────────────────────────────────────────────────────
    p_mail = sub.add_parser("mail", help="Read and search emails")
    add_output_args(p_mail)
    mail_sub = p_mail.add_subparsers(dest="command", required=True)

    cmd_mail_unread = mail_sub.add_parser("unread", help="List unread emails from the Inbox")
    add_output_args(cmd_mail_unread)
    cmd_mail_unread.add_argument("--count", "-n", type=int, default=50, help="Max results")
    cmd_mail_unread.set_defaults(func=mail_commands.cmd_unread, _mail_ctx=mail_ctx)

    cmd_mail_search = mail_sub.add_parser("search", help="Search emails by keyword")
    add_output_args(cmd_mail_search)
    cmd_mail_search.add_argument("query", help="Search term")
    cmd_mail_search.add_argument("--count", "-n", type=int, default=20, help="Max results")
    cmd_mail_search.set_defaults(func=mail_commands.cmd_mail, _mail_ctx=mail_ctx)

    cmd_mail_read = mail_sub.add_parser("read", help="Read an email by index or ID")
    add_output_args(cmd_mail_read)
    cmd_mail_read.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_read.set_defaults(func=mail_commands.cmd_read, _mail_ctx=mail_ctx)

    cmd_mail_mark_read = mail_sub.add_parser("mark-read", help="Mark an email as read")
    cmd_mail_mark_read.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_mark_read.set_defaults(func=mail_commands.cmd_mark_read, _mail_ctx=mail_ctx)

    cmd_mail_mark_unread = mail_sub.add_parser("mark-unread", help="Mark an email as unread")
    cmd_mail_mark_unread.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_mark_unread.set_defaults(func=mail_commands.cmd_mark_unread, _mail_ctx=mail_ctx)

    cmd_mail_archive = mail_sub.add_parser("archive", help="Archive an email")
    cmd_mail_archive.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_archive.set_defaults(func=mail_commands.cmd_archive, _mail_ctx=mail_ctx)

    cmd_mail_move = mail_sub.add_parser("move", help="Move an email to another folder")
    cmd_mail_move.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_move.add_argument("--folder", "-f", required=True, help="Folder index, well-known name, folder name, or folder ID")
    cmd_mail_move.set_defaults(func=mail_commands.cmd_move, _mail_ctx=mail_ctx)

    cmd_mail_folders = mail_sub.add_parser("folders", help="List mail folders")
    add_output_args(cmd_mail_folders)
    cmd_mail_folders.add_argument("--count", "-n", type=int, default=100, help="Max folders")
    cmd_mail_folders.set_defaults(func=mail_commands.cmd_folders, _mail_ctx=mail_ctx)

    cmd_mail_attachments = mail_sub.add_parser("attachments", help="List message attachments")
    add_output_args(cmd_mail_attachments)
    cmd_mail_attachments.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_attachments.set_defaults(func=mail_commands.cmd_attachments, _mail_ctx=mail_ctx)

    cmd_mail_download_attachments = mail_sub.add_parser("download-attachments", help="Download message attachments")
    cmd_mail_download_attachments.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_download_attachments.add_argument("--dir", "-d", default="attachments", help="Output directory")
    cmd_mail_download_attachments.add_argument("--overwrite", action="store_true", help="Overwrite files with matching names")
    cmd_mail_download_attachments.add_argument(
        "--include-inline",
        action="store_true",
        help="Also download attachments marked as inline (e.g. embedded images)",
    )
    cmd_mail_download_attachments.set_defaults(func=mail_commands.cmd_download_attachments, _mail_ctx=mail_ctx)

    cmd_mail_links = mail_sub.add_parser("links", help="List links and inline image URLs in an email body")
    add_output_args(cmd_mail_links)
    cmd_mail_links.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_links.add_argument(
        "--share-only",
        action="store_true",
        help="Only show SharePoint/OneDrive style links",
    )
    cmd_mail_links.set_defaults(func=mail_commands.cmd_links, _mail_ctx=mail_ctx)

    cmd_mail_download_links = mail_sub.add_parser(
        "download-links",
        help="Download SharePoint/OneDrive links from an email body via Graph",
    )
    cmd_mail_download_links.add_argument("message_id", help="Message index (from last search) or full ID")
    cmd_mail_download_links.add_argument("--dir", "-d", default="attachments", help="Output directory")
    cmd_mail_download_links.add_argument("--overwrite", action="store_true", help="Overwrite files with matching names")
    cmd_mail_download_links.set_defaults(func=mail_commands.cmd_download_links, _mail_ctx=mail_ctx)

    # Direct email sending is intentionally not exposed to agents.
    # Implementation is preserved in cmd_send_mail / OutlookClient.send_mail,
    # but the parser is disabled for safety.

    tasks_ctx = tasks_commands.build_ctx(
        console=console,
        get_client=_get_client,
        resolve_task_id=_resolve_task_id,
    )

    # ── Tasks ─────────────────────────────────────────────────────────
    p_task = sub.add_parser("task", help="Manage tasks")
    add_output_args(p_task)
    task_sub = p_task.add_subparsers(dest="command", required=True)

    cmd_task_l = task_sub.add_parser("list", help="List active tasks")
    add_output_args(cmd_task_l)
    cmd_task_l.add_argument("--count", "-n", type=int, default=50, help="Max tasks")
    cmd_task_l.set_defaults(func=tasks_commands.cmd_task_list, _tasks_ctx=tasks_ctx)

    cmd_task_c = task_sub.add_parser("create", help="Create a task")
    cmd_task_c.add_argument("subject", help="Task description")
    cmd_task_c.set_defaults(func=tasks_commands.cmd_task_create, _tasks_ctx=tasks_ctx)

    cmd_task_update_cmd = task_sub.add_parser("update", help="Update a task")
    cmd_task_update_cmd.add_argument("task_id", help="Task index or full ID")
    cmd_task_update_cmd.add_argument("--subject", help="New task subject")
    cmd_task_update_cmd.add_argument("--due", help="Due date/time")
    cmd_task_update_cmd.add_argument("--importance", choices=["Low", "Normal", "High"], help="Task importance")
    cmd_task_update_cmd.add_argument("--status", choices=["NotStarted", "InProgress", "Completed", "WaitingOnOthers", "Deferred"], help="Task status")
    cmd_task_update_cmd.set_defaults(func=tasks_commands.cmd_task_update, _tasks_ctx=tasks_ctx)

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
    add_output_args(p_cal)
    cal_sub = p_cal.add_subparsers(dest="command", required=True)

    cmd_cal_agenda = cal_sub.add_parser("agenda", help="List upcoming calendar events")
    add_output_args(cmd_cal_agenda)
    cmd_cal_agenda.add_argument("--days", "-d", type=int, default=7, help="Days to look ahead")
    cmd_cal_agenda.add_argument("--count", "-n", type=int, default=20, help="Max events")
    cmd_cal_agenda.add_argument("--plain", action="store_true", help="Output plain text instead of a table")
    cmd_cal_agenda.set_defaults(func=calendar_commands.cmd_cal_agenda, _calendar_ctx=calendar_ctx)

    cmd_cal_show = cal_sub.add_parser("show", help="Show event details")
    add_output_args(cmd_cal_show)
    cmd_cal_show.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_show.set_defaults(func=calendar_commands.cmd_cal_show, _calendar_ctx=calendar_ctx)

    cmd_cal_create_cmd = cal_sub.add_parser("create", help="Create an event")
    cmd_cal_create_cmd.add_argument("subject", help="Event subject")
    cmd_cal_create_cmd.add_argument("start", help="Start time (e.g. '2026-04-10 14:00')")
    cmd_cal_create_cmd.add_argument("end", help="End time (e.g. '2026-04-10 15:00')")
    cmd_cal_create_cmd.add_argument("--location", "-l", help="Location")
    cmd_cal_create_cmd.add_argument("--body", "-b", help="Event description")
    cmd_cal_create_cmd.add_argument("--attendee", dest="attendees", action="append", help="Attendee email (repeatable)")
    cmd_cal_create_cmd.add_argument(
        "--teams",
        action="store_true",
        help="Create as a Teams meeting (adds Teams join link)",
    )
    cmd_cal_create_cmd.set_defaults(func=calendar_commands.cmd_cal_create, _calendar_ctx=calendar_ctx)

    cmd_cal_rooms_cmd = cal_sub.add_parser("rooms", help="Find rooms")
    add_output_args(cmd_cal_rooms_cmd)
    cmd_cal_rooms_cmd.add_argument("--room-list", help="Optional room list email/address")
    cmd_cal_rooms_cmd.set_defaults(func=calendar_commands.cmd_cal_rooms, _calendar_ctx=calendar_ctx)

    cmd_cal_availability_cmd = cal_sub.add_parser("availability", help="Get free/busy availability")
    add_output_args(cmd_cal_availability_cmd)
    cmd_cal_availability_cmd.add_argument("--attendee", required=True, action="append", help="Attendee email (repeatable)")
    cmd_cal_availability_cmd.add_argument("--start", required=True, help="Start time")
    cmd_cal_availability_cmd.add_argument("--end", required=True, help="End time")
    cmd_cal_availability_cmd.add_argument("--interval", type=int, default=30, help="Availability interval minutes")
    cmd_cal_availability_cmd.set_defaults(func=calendar_commands.cmd_cal_availability, _calendar_ctx=calendar_ctx)

    cmd_cal_find_time_cmd = cal_sub.add_parser("find-time", help="Suggest meeting times")
    add_output_args(cmd_cal_find_time_cmd)
    cmd_cal_find_time_cmd.add_argument("--attendee", required=True, action="append", help="Attendee email (repeatable)")
    cmd_cal_find_time_cmd.add_argument("--start", required=True, help="Window start time")
    cmd_cal_find_time_cmd.add_argument("--end", required=True, help="Window end time")
    cmd_cal_find_time_cmd.add_argument("--duration", type=int, default=30, help="Meeting duration minutes")
    cmd_cal_find_time_cmd.add_argument("--count", "-n", type=int, default=10, help="Max suggestions")
    cmd_cal_find_time_cmd.set_defaults(func=calendar_commands.cmd_cal_find_time, _calendar_ctx=calendar_ctx)

    cmd_cal_update_cmd = cal_sub.add_parser("update", help="Update an event")
    cmd_cal_update_cmd.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_update_cmd.add_argument("--subject", help="New event subject")
    cmd_cal_update_cmd.add_argument("--start", help="New start time")
    cmd_cal_update_cmd.add_argument("--end", help="New end time")
    cmd_cal_update_cmd.add_argument("--location", help="New location, use empty string to clear")
    cmd_cal_update_cmd.add_argument("--body", help="New event body, use empty string to clear")
    cmd_cal_update_group = cmd_cal_update_cmd.add_mutually_exclusive_group()
    cmd_cal_update_group.add_argument(
        "--teams",
        dest="teams",
        action="store_true",
        default=None,
        help="Make the event a Teams meeting",
    )
    cmd_cal_update_group.add_argument(
        "--no-teams",
        dest="teams",
        action="store_false",
        help="Remove the Teams meeting flag",
    )
    cmd_cal_update_cmd.set_defaults(func=calendar_commands.cmd_cal_update, _calendar_ctx=calendar_ctx)

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
    add_output_args(p_contact)
    contact_sub = p_contact.add_subparsers(dest="command", required=True)

    cmd_contact_search = contact_sub.add_parser("search", help="Search contacts by name or email")
    add_output_args(cmd_contact_search)
    cmd_contact_search.add_argument("query", help="Name or email to search for")
    cmd_contact_search.add_argument("--count", "-n", type=int, default=10, help="Max results")
    cmd_contact_search.set_defaults(func=contacts_commands.cmd_contacts, _contacts_ctx=contacts_ctx)

    cmd_contact_create = contact_sub.add_parser("create", help="Create a personal contact")
    cmd_contact_create.add_argument("--name", required=True, help="Display name")
    cmd_contact_create.add_argument("--email", required=True, help="Email address")
    cmd_contact_create.add_argument("--given-name", help="Given name")
    cmd_contact_create.add_argument("--surname", help="Surname")
    cmd_contact_create.add_argument("--company", help="Company name")
    cmd_contact_create.add_argument("--mobile", help="Mobile phone")
    cmd_contact_create.set_defaults(func=contacts_commands.cmd_contact_create, _contacts_ctx=contacts_ctx)

    cmd_contact_update = contact_sub.add_parser("update", help="Update a personal contact")
    cmd_contact_update.add_argument("contact_id", help="Contact index from search or full ID")
    cmd_contact_update.add_argument("--name", help="Display name")
    cmd_contact_update.add_argument("--email", help="Email address")
    cmd_contact_update.add_argument("--given-name", help="Given name")
    cmd_contact_update.add_argument("--surname", help="Surname")
    cmd_contact_update.add_argument("--company", help="Company name")
    cmd_contact_update.add_argument("--mobile", help="Mobile phone")
    cmd_contact_update.set_defaults(func=contacts_commands.cmd_contact_update, _contacts_ctx=contacts_ctx)

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
    cmd_teams_list_cmd.add_argument("--sort-received", action="store_true",
        help="Sort by last received message (slower — makes one API call per chat)")
    cmd_teams_list_cmd.set_defaults(func=teams_commands.cmd_teams_list, _teams_ctx=teams_ctx)

    cmd_teams_search_cmd = teams_sub.add_parser("search", help="Search Teams chats by topic or participant")
    cmd_teams_search_cmd.add_argument("query", nargs="+", help="Name, email, or chat topic to search for")
    cmd_teams_search_cmd.add_argument("--count", "-n", type=int, default=20, help="Max matching chats")
    cmd_teams_search_cmd.add_argument("--scan", type=int, default=200, help="Max chats to scan")
    cmd_teams_search_cmd.add_argument("--member-count", type=int, default=100, help="Max members to inspect per chat")
    cmd_teams_search_cmd.set_defaults(func=teams_commands.cmd_teams_search, _teams_ctx=teams_ctx)

    cmd_teams_show_cmd = teams_sub.add_parser("show", help="Show Teams chat details")
    cmd_teams_show_cmd.add_argument("chat_id", help="Chat index, cached ID suffix, or full ID")
    cmd_teams_show_cmd.set_defaults(func=teams_commands.cmd_teams_show, _teams_ctx=teams_ctx)

    cmd_teams_self_cmd = teams_sub.add_parser("self", help="Show your Teams self-chat")
    cmd_teams_self_cmd.set_defaults(func=teams_commands.cmd_teams_self, _teams_ctx=teams_ctx)

    cmd_teams_messages_cmd = teams_sub.add_parser("messages", help="Read Teams chat messages")
    cmd_teams_messages_cmd.add_argument("chat_id", help="Chat index, cached ID suffix, or full ID")
    cmd_teams_messages_cmd.add_argument("--count", "-n", type=int, default=20, help="Max messages")
    cmd_teams_messages_cmd.set_defaults(func=teams_commands.cmd_teams_messages, _teams_ctx=teams_ctx)

    cmd_teams_attachments_cmd = teams_sub.add_parser(
        "attachments",
        help="List Teams attachments and SharePoint/OneDrive links in a chat",
    )
    cmd_teams_attachments_cmd.add_argument("chat_id", help="Chat index, cached ID suffix, or full ID")
    cmd_teams_attachments_cmd.add_argument(
        "--scan",
        type=int,
        default=50,
        help="How many recent messages to scan for attachments and links",
    )
    cmd_teams_attachments_cmd.set_defaults(
        func=teams_commands.cmd_teams_attachments,
        _teams_ctx=teams_ctx,
    )

    cmd_teams_download_attachments_cmd = teams_sub.add_parser(
        "download-attachments",
        help="Download Teams attachments/share links via Graph",
    )
    cmd_teams_download_attachments_cmd.add_argument("chat_id", help="Chat index, cached ID suffix, or full ID")
    cmd_teams_download_attachments_cmd.add_argument(
        "--attachment",
        "-a",
        help="Attachment index from `teams attachments`, full URL, or substring match",
    )
    cmd_teams_download_attachments_cmd.add_argument(
        "--scan",
        type=int,
        default=50,
        help="How many recent messages to scan for attachments and links",
    )
    cmd_teams_download_attachments_cmd.add_argument("--dir", "-d", default="attachments", help="Output directory")
    cmd_teams_download_attachments_cmd.add_argument("--overwrite", action="store_true", help="Overwrite files with matching names")
    cmd_teams_download_attachments_cmd.set_defaults(
        func=teams_commands.cmd_teams_download_attachments,
        _teams_ctx=teams_ctx,
    )

    # Teams sending is intentionally not exposed to agents. Teams reading stays enabled.
    # Implementation is preserved in cmd_teams_send / OutlookClient.send_teams_message,
    # but the parser is disabled for safety.

    # ── Mailbox ───────────────────────────────────────────────────────
    p_mailbox = sub.add_parser("mailbox", help="Manage mailbox settings")
    add_output_args(p_mailbox)
    mailbox_sub = p_mailbox.add_subparsers(dest="command", required=True)
    cmd_mailbox_show_parser = mailbox_sub.add_parser("show", help="Show mailbox settings")
    add_output_args(cmd_mailbox_show_parser)
    cmd_mailbox_show_parser.set_defaults(func=cmd_mailbox_show)
    cmd_mailbox_update_parser = mailbox_sub.add_parser("update", help="Update mailbox settings")
    cmd_mailbox_update_parser.add_argument("--timezone", help="Mailbox timezone")
    cmd_mailbox_update_parser.add_argument("--auto-reply-status", choices=["Disabled", "AlwaysEnabled", "Scheduled"], help="Automatic replies status")
    cmd_mailbox_update_parser.add_argument("--internal-reply", help="Internal auto-reply message")
    cmd_mailbox_update_parser.add_argument("--external-reply", help="External auto-reply message")
    cmd_mailbox_update_parser.set_defaults(func=cmd_mailbox_update)

    # ── Auth ──────────────────────────────────────────────────────────
    p_auth = sub.add_parser("auth", help="Manage authentication")
    add_output_args(p_auth)
    p_auth.add_argument(
        "auth_command",
        nargs="?",
        choices=["login", "status", "clear", "scopes"],
        default="login",
        help="Auth action to run",
    )
    p_auth.add_argument(
        "--headed",
        action="store_true",
        help="Open a visible browser instead of running fully headless",
    )

    # ── Files (OneDrive + SharePoint) ────────────────────────────────
    p_files = sub.add_parser("files", help="Browse and manage OneDrive and SharePoint files")
    files_sub = p_files.add_subparsers(dest="command", required=True)

    cmd_files_sites_p = files_sub.add_parser("sites", help="List SharePoint sites")
    cmd_files_sites_p.set_defaults(func=cmd_files_sites)

    cmd_files_list_p = files_sub.add_parser("list", help="List files and folders")
    cmd_files_list_p.add_argument("path", nargs="?", default="", help="Folder path (default: root)")
    cmd_files_list_p.add_argument("--site", metavar="NAME", help="SharePoint site name (partial match)")
    cmd_files_list_p.set_defaults(func=cmd_files_list)

    cmd_files_upload_p = files_sub.add_parser("upload", help="Upload a file")
    cmd_files_upload_p.add_argument("file", help="Local file path")
    cmd_files_upload_p.add_argument("dest", nargs="?", default="", help="Remote folder path (default: root)")
    cmd_files_upload_p.add_argument("--site", metavar="NAME", help="SharePoint site name")
    cmd_files_upload_p.set_defaults(func=cmd_files_upload)

    cmd_files_download_p = files_sub.add_parser("download", help="Download a file")
    cmd_files_download_p.add_argument("path", help="Remote file path")
    cmd_files_download_p.add_argument("dest", nargs="?", default=".", help="Local output file or directory")
    cmd_files_download_p.add_argument("--site", metavar="NAME", help="SharePoint site name")
    cmd_files_download_p.add_argument("--overwrite", action="store_true", help="Overwrite files with matching names")
    cmd_files_download_p.set_defaults(func=cmd_files_download)

    cmd_files_mkdir_p = files_sub.add_parser("mkdir", help="Create a folder")
    cmd_files_mkdir_p.add_argument("path", help="Folder path to create")
    cmd_files_mkdir_p.add_argument("--site", metavar="NAME", help="SharePoint site name")
    cmd_files_mkdir_p.set_defaults(func=cmd_files_mkdir)

    cmd_files_rename_p = files_sub.add_parser("rename", help="Rename a file or folder")
    cmd_files_rename_p.add_argument("path", help="Current path")
    cmd_files_rename_p.add_argument("name", help="New name")
    cmd_files_rename_p.add_argument("--site", metavar="NAME", help="SharePoint site name")
    cmd_files_rename_p.set_defaults(func=cmd_files_rename)

    cmd_files_move_p = files_sub.add_parser("move", help="Move a file or folder")
    cmd_files_move_p.add_argument("path", help="Source path")
    cmd_files_move_p.add_argument("dest", help="Destination folder path")
    cmd_files_move_p.add_argument("--site", metavar="NAME", help="SharePoint site name")
    cmd_files_move_p.set_defaults(func=cmd_files_move)

    # ── Signature ─────────────────────────────────────────────────────
    p_sig = sub.add_parser("signature", help="Manage email signatures")
    sig_sub = p_sig.add_subparsers(dest="command", required=True)
    cmd_sig_fetch = sig_sub.add_parser(
        "fetch",
        help="Fetch your OWA signature and save to configured files",
    )
    cmd_sig_fetch.add_argument(
        "--headed",
        action="store_true",
        help="Run with a visible browser window (useful if headless fails)",
    )
    cmd_sig_fetch.set_defaults(func=cmd_signature_fetch)

    # ── Gateway ───────────────────────────────────────────────────────
    p_gateway = sub.add_parser("gateway", help="Teams-to-pi gateway daemon")
    gateway_sub = p_gateway.add_subparsers(dest="command", required=True)

    gw_start = gateway_sub.add_parser("start", help="Start the gateway daemon")
    gw_start.add_argument("--chat-id", metavar="ID", help="Teams chat ID to monitor")
    gw_start.add_argument("--self-chat", action="store_true", help="Monitor your Teams self-chat")
    gw_start.add_argument("--trigger", metavar="TEXT", help="Trigger string (default: @Marlow)")
    gw_start.add_argument("--poll", type=int, metavar="SECONDS", help="Poll interval seconds (default: 30)")
    gw_start.add_argument("--provider", metavar="NAME", help="Pi provider name")
    gw_start.add_argument("--model", metavar="MODEL", help="Pi model pattern or ID")
    gw_start.add_argument("--thinking", metavar="LEVEL", help="Pi thinking level")
    gw_start.add_argument("--models", metavar="PATTERNS", help="Comma-separated Pi model cycle patterns")
    gw_start.set_defaults(func=cmd_gateway_start)

    gw_stop = gateway_sub.add_parser("stop", help="Stop the gateway daemon")
    gw_stop.set_defaults(func=cmd_gateway_stop)

    gw_status = gateway_sub.add_parser("status", help="Show gateway status and recent log")
    gw_status.set_defaults(func=cmd_gateway_status)

    # ── Config ────────────────────────────────────────────────────────
    p_config = sub.add_parser("config", help="Inspect local configuration")
    add_output_args(p_config)
    config_sub = p_config.add_subparsers(dest="command", required=True)
    cmd_config_check_parser = config_sub.add_parser("check", help="Validate local configuration")
    add_output_args(cmd_config_check_parser)
    cmd_config_check_parser.set_defaults(func=cmd_config_check)

    args = parser.parse_args()

    if args.domain == "auth":
        cmd_auth(args)
    elif hasattr(args, "func"):
        args.func(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
