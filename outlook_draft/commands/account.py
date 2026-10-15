from __future__ import annotations

import argparse
import base64
import json
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

from rich import box
from rich.panel import Panel
from rich.table import Table

from .. import config, runtime
from ..errors import OutlookAPIError
from ..formatting import compose_email_html as _compose_email_html, text_to_html as _text_to_html, parse_recipients as _parse_recipients
from ..outlook_client import OutlookClient
from ..progress import spinner
from ..references import _resolve_message_id
from ..runtime import console
from ..signatures import load_signature
from ..token_manager import TokenManager
from ..parser import add_output_args, _add_yes_arg

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
    }

def _format_env_value(value: str) -> str:
    escaped = value.replace("\\", "\\\\").replace('"', '\\"')
    return f'"{escaped}"'

def _read_account_email(account: str) -> str:
    path = config.account_env_file(account)
    return _read_env_email(path)

def _read_env_email(path: Path) -> str:
    try:
        for line in path.read_text().splitlines():
            if line.startswith("MS_EMAIL="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    except FileNotFoundError:
        pass
    return ""

def cmd_account_list(args: argparse.Namespace) -> None:
    accounts = config.list_accounts()
    active = config.ACTIVE_ACCOUNT
    rows = []
    if config.CONFIG_FILE.exists():
        rows.append({
            "name": "(default)",
            "active": active is None,
            "email": _read_env_email(config.CONFIG_FILE),
            "envFile": str(config.CONFIG_FILE),
            "dataDir": str(config.DATA_DIR),
            "kind": "legacy",
        })
    rows.extend(
        {
            "name": account,
            "active": account == active,
            "email": _read_account_email(account),
            "envFile": str(config.account_env_file(account)),
            "dataDir": str(config.ACCOUNTS_DATA_DIR / account),
            "kind": "profile",
        }
        for account in accounts
    )
    if args.json:
        print(json.dumps({
            "activeAccount": active,
            "accounts": rows,
        }, indent=2))
        return
    if not rows:
        console.print("[dim]No account profiles or default .env configured.[/]")
        return
    table = Table(title="Outlook accounts", box=box.SIMPLE)
    table.add_column("Active")
    table.add_column("Name")
    table.add_column("Email")
    table.add_column("Env file")
    for row in rows:
        table.add_row(
            "*" if row["active"] else "",
            str(row["name"]),
            str(row["email"]),
            str(row["envFile"]),
        )
    console.print(table)

def cmd_account_current(args: argparse.Namespace) -> None:
    active = config.ACTIVE_ACCOUNT
    payload = {
        "activeAccount": active,
        "email": config.MS_EMAIL,
        "envFile": str(config.account_env_file(active)) if active else str(config.CONFIG_FILE),
        "dataDir": str(config.ACCOUNT_DATA_DIR),
        "sessionDir": str(config.SESSION_DIR),
    }
    if args.json:
        print(json.dumps(payload, indent=2))
        return
    if not active:
        console.print("[yellow]No account profile is active; using legacy/global config.[/]")
        console.print(f"Config:  {config.CONFIG_FILE}")
        console.print(f"Data:    {config.DATA_DIR}")
        return
    console.print(f"Active account: [bold]{active}[/]")
    console.print(f"Email:          {config.MS_EMAIL or '(missing)'}")
    console.print(f"Env file:       {config.account_env_file(active)}")
    console.print(f"Data dir:       {config.ACCOUNT_DATA_DIR}")

def cmd_account_add(args: argparse.Namespace) -> None:
    path = config.account_env_file(args.name)
    if path.exists() and not args.force:
        console.print(f"[red]Account '{args.name}' already exists. Use --force to overwrite.[/]")
        sys.exit(1)
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        f"MS_EMAIL={_format_env_value(args.email)}",
    ]
    if args.password:
        lines.append(f"MS_PASSWORD={_format_env_value(args.password)}")
    else:
        lines.append("MS_PASSWORD=")
    lines.append(f"LOCAL_TIMEZONE={_format_env_value(args.local_timezone)}")
    lines.append(f"OUTLOOK_TIMEZONE={_format_env_value(args.outlook_timezone)}")
    path.write_text("\n".join(lines) + "\n")
    path.chmod(0o600)
    (config.ACCOUNTS_DATA_DIR / args.name).mkdir(parents=True, exist_ok=True)
    if args.switch:
        config.set_active_account(args.name)
        refresh_account_paths()
    console.print(f"[green]Account profile saved:[/] {args.name}")
    console.print(f"Env file: {path}")
    if not args.password:
        console.print("[yellow]MS_PASSWORD is blank; add it to the env file before headless auth.[/]")

def cmd_account_switch(args: argparse.Namespace) -> None:
    if args.name == "default":
        config.clear_active_account()
        refresh_account_paths()
        console.print("[green]Active account:[/] (default)")
        return
    try:
        config.set_active_account(args.name)
    except FileNotFoundError as e:
        console.print(f"[red]{e}[/]")
        sys.exit(1)
    refresh_account_paths()
    console.print(f"[green]Active account:[/] {args.name}")

def cmd_account_remove(args: argparse.Namespace) -> None:
    path = config.account_env_file(args.name)
    if not path.exists():
        console.print(f"[red]No account profile named '{args.name}'.[/]")
        sys.exit(1)
    path.unlink()
    if args.purge_data:
        shutil.rmtree(config.ACCOUNTS_DATA_DIR / args.name, ignore_errors=True)
    if config.ACTIVE_ACCOUNT == args.name:
        config.clear_active_account()
        refresh_account_paths()
    console.print(f"[green]Removed account profile:[/] {args.name}")

def cmd_auth(args: argparse.Namespace) -> None:
    """Manage authentication."""
    command = args.auth_command or "login"
    if command == "status":
        statuses = [
            _token_status(config.OUTLOOK_TOKEN_DOMAIN, "Outlook API"),
            _token_status(config.GRAPH_TOKEN_DOMAIN, "Microsoft Graph"),
        ]
        if args.json:
            print(json.dumps({
                "account": {
                    "active": config.ACTIVE_ACCOUNT,
                    "email": config.MS_EMAIL,
                    "sessionDir": str(config.SESSION_DIR),
                },
                "tokens": statuses,
            }, indent=2))
            return
        title = "Auth status"
        if config.ACTIVE_ACCOUNT:
            title += f" ({config.ACTIVE_ACCOUNT})"
        table = Table(title=title, box=box.SIMPLE)
        table.add_column("Token")
        table.add_column("Domain")
        table.add_column("Status")
        table.add_column("Expires")
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
            table.add_row(str(item["label"]), str(item["domain"]), status, expires)
        console.print(table)
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

    add("account", True, config.ACTIVE_ACCOUNT or "legacy/global")
    add("MS_EMAIL", bool(config.MS_EMAIL), config.MS_EMAIL or "missing")
    add("MS_PASSWORD", bool(config.MS_PASSWORD), "set" if config.MS_PASSWORD else "missing")
    add("config file", config.CONFIG_FILE.exists(), str(config.CONFIG_FILE))
    if config.ACTIVE_ACCOUNT:
        account_file = config.account_env_file(config.ACTIVE_ACCOUNT)
        add("account env file", account_file.exists(), str(account_file))
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
    client = runtime.get_client()
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
    client = runtime.get_client()
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

def refresh_account_paths() -> None:
    """Refresh imported path constants after activating an account."""
    from .. import cache
    from . import mail as mail_commands, calendar as calendar_commands
    from . import contacts as contacts_commands, tasks as tasks_commands, teams as teams_commands

    config.ensure_dirs()
    cache.refresh_paths()

    mail_commands.MAIL_CACHE = cache.MAIL_CACHE
    mail_commands.MAIL_FOLDER_CACHE = cache.MAIL_FOLDER_CACHE
    calendar_commands.CAL_CACHE = cache.CAL_CACHE
    contacts_commands.CONTACT_CACHE = cache.CONTACT_CACHE
    tasks_commands.TASK_CACHE = cache.TASK_CACHE
    teams_commands.TEAMS_CACHE = cache.TEAMS_CACHE


def register(sub) -> None:
    p_account = sub.add_parser("account", help="Manage account profiles")
    add_output_args(p_account)
    account_sub = p_account.add_subparsers(dest="command", required=True)

    cmd_account_list_parser = account_sub.add_parser("list", help="List account profiles")
    add_output_args(cmd_account_list_parser)
    cmd_account_list_parser.set_defaults(func=cmd_account_list)

    cmd_account_current_parser = account_sub.add_parser("current", help="Show the active account")
    add_output_args(cmd_account_current_parser)
    cmd_account_current_parser.set_defaults(func=cmd_account_current)

    cmd_account_add_parser = account_sub.add_parser("add", help="Create an account profile")
    cmd_account_add_parser.add_argument("name", help="Profile name, e.g. wpp or ikea")
    cmd_account_add_parser.add_argument("--email", required=True, help="Account email address")
    cmd_account_add_parser.add_argument("--password", help="Password for headless auth")
    cmd_account_add_parser.add_argument("--local-timezone", default=config.LOCAL_TIMEZONE, help="Local timezone")
    cmd_account_add_parser.add_argument("--outlook-timezone", default=config.OUTLOOK_TIMEZONE, help="Outlook timezone")
    cmd_account_add_parser.add_argument("--switch", action="store_true", help="Make this profile active")
    cmd_account_add_parser.add_argument("--force", action="store_true", help="Overwrite an existing profile")
    cmd_account_add_parser.set_defaults(func=cmd_account_add)

    cmd_account_switch_parser = account_sub.add_parser("switch", help="Set the active account")
    cmd_account_switch_parser.add_argument("name", help="Profile name")
    cmd_account_switch_parser.set_defaults(func=cmd_account_switch)

    cmd_account_remove_parser = account_sub.add_parser("remove", help="Remove an account profile")
    cmd_account_remove_parser.add_argument("name", help="Profile name")
    cmd_account_remove_parser.add_argument("--purge-data", action="store_true", help="Also delete tokens, browser state, caches, signatures, and gateway state")
    cmd_account_remove_parser.set_defaults(func=cmd_account_remove)

def register_mailbox(sub) -> None:
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

def register_auth(sub) -> None:
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

def register_config(sub) -> None:
    p_config = sub.add_parser("config", help="Inspect local configuration")
    add_output_args(p_config)
    config_sub = p_config.add_subparsers(dest="command", required=True)
    cmd_config_check_parser = config_sub.add_parser("check", help="Validate local configuration")
    add_output_args(cmd_config_check_parser)
    cmd_config_check_parser.set_defaults(func=cmd_config_check)
