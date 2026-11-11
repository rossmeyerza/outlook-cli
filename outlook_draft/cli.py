"""Outlook CLI entry point and command dispatch."""
from __future__ import annotations

import argparse
import json
import sys

from . import config
from .commands import (
    account,
    calendar,
    completion,
    contacts,
    copilot,
    drafts,
    files,
    gateway_parser,
    mail,
    mcp,
    signature,
    tasks,
    teams,
)
from .errors import CliError, OutlookDraftError
from .parser import add_output_args
from .runtime import console, error_console
from .safety import SafetyError, check_command, command_allowed, policy


def _extract_account_arg(argv: list[str]) -> tuple[list[str], str | None]:
    """Allow --account to appear before or after the top-level command."""
    cleaned: list[str] = []
    account: str | None = None
    idx = 0
    while idx < len(argv):
        arg = argv[idx]
        if arg == "--account":
            if idx + 1 >= len(argv):
                cleaned.append(arg)
                idx += 1
                continue
            account = argv[idx + 1]
            idx += 2
            continue
        if arg.startswith("--account="):
            account = arg.split("=", 1)[1]
            idx += 1
            continue
        cleaned.append(arg)
        idx += 1
    return cleaned, account

def _extract_safety_args(argv: list[str]) -> tuple[list[str], list[str]]:
    """Move global safety flags before the command without altering its arguments."""
    cleaned, extracted = [], []
    flags = {"--readonly", "--no-send", "--dry-run"}
    valued = {"--enable-commands", "--disable-commands"}
    index = 0
    while index < len(argv):
        arg = argv[index]
        if arg == "--":
            cleaned.extend(argv[index:])
            break
        name = arg.split("=", 1)[0]
        if name in flags | valued:
            extracted.append(arg)
            if name in valued and "=" not in arg and index + 1 < len(argv):
                index += 1
                extracted.append(argv[index])
        else:
            cleaned.append(arg)
        index += 1
    return cleaned, extracted


def _selected_command(parser: argparse.ArgumentParser, args: argparse.Namespace) -> str:
    parts = []
    while True:
        action = next((item for item in parser._actions
                       if isinstance(item, argparse._SubParsersAction)), None)
        if action is None:
            return ".".join(parts)
        selected = getattr(args, action.dest, None)
        if selected not in action.choices:
            return ".".join(parts)
        parts.append(selected)
        parser = action.choices[selected]


def _print_plan(args: argparse.Namespace, command: str) -> None:
    target_fields = {"message_id", "draft_id", "event_id", "task_id", "chat_id", "chat_ref",
                     "name", "folder", "parent", "site", "library", "item_id", "file_id",
                     "message_ids", "action", "id"}
    targets = {key: value for key, value in vars(args).items()
               if key in target_fields and value is not None}
    plan = {"kind": "plan", "command": command, "targets": targets}
    if args.json:
        console.print(json.dumps(plan), markup=False, highlight=False, soft_wrap=True)
    else:
        console.print(f"Plan only (no client, authentication or writes): {command}", markup=False)
        for key, value in targets.items():
            console.print(f"  {key}: {value}", markup=False)


def _run_with_safety(args: argparse.Namespace, parser: argparse.ArgumentParser) -> None:
    command = _selected_command(parser, args)
    if not command_allowed(command, args.enable_commands, args.disable_commands):
        raise SafetyError(f"Command policy blocks '{command}'.")
    if args.dry_run:
        _print_plan(args, command)
        return
    with policy(readonly=args.readonly, no_send=args.no_send):
        check_command(command)
        _dispatch(args, parser)


def _add_safety_args(parser: argparse.ArgumentParser) -> None:
    flags = {
        "--readonly": "Block write requests, including confirmed sends",
        "--no-send": "Block message sends and event changes that can send invitations",
        "--dry-run": "Print a plan only, with no client, authentication or writes",
    }
    for flag, description in flags.items():
        parser.add_argument(flag, action="store_true", help=description)
    for flag, description in (("--enable-commands", "Allow only these command prefixes"),
                              ("--disable-commands", "Block these command prefixes")):
        parser.add_argument(flag, action="append", metavar="PREFIXES",
                            help=f"{description} (comma-separated; dots or spaces separate command names)")


def build_parser(account_override: str | None = None) -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="outlook-cli",
        description="Manage Outlook mail, drafts, calendar, contacts, and Teams chats",
    )
    parser.set_defaults(json=False)
    parser.set_defaults(account=account_override)
    add_output_args(parser)
    parser.add_argument(
        "--no-spinner",
        action="store_true",
        help="Disable interactive progress spinners",
    )
    parser.add_argument(
        "--account",
        help="Use an account profile for this invocation",
    )
    _add_safety_args(parser)
    sub = parser.add_subparsers(dest="domain", required=True)

    sub._parser = parser
    account.register(sub)
    drafts.register(sub)
    mail.register(sub)
    tasks.register(sub)
    calendar.register(sub)
    contacts.register(sub)
    teams.register(sub)
    account.register_mailbox(sub)
    account.register_auth(sub)
    files.register(sub)
    copilot.register(sub)
    signature.register(sub)
    gateway_parser.register(sub)
    account.register_config(sub)
    mcp.register(sub)
    completion.register(sub)

    return parser


def main() -> None:
    argv, account_override = _extract_account_arg(sys.argv[1:])
    argv, safety_args = _extract_safety_args(argv)
    parser = build_parser(account_override)

    # Let argcomplete service tab-completion requests for bash/zsh/fish.
    try:
        import argcomplete
    except ImportError:
        pass
    else:
        argcomplete.autocomplete(parser)

    args = parser.parse_args(safety_args + argv)

    try:
        _run_with_safety(args, parser)
    except (CliError, OutlookDraftError) as error:
        if str(error):
            # In JSON mode, stdout carries only JSON.
            out = error_console if args.json else console
            out.print(str(error) if isinstance(error, CliError) else f"[red]{error}[/]")
        raise SystemExit(getattr(error, "exit_code", 1)) from error


def _dispatch(args: argparse.Namespace, parser: argparse.ArgumentParser) -> None:
    if args.account:
        if not config.account_exists(args.account):
            raise CliError(f"[red]No account profile named '{args.account}'.[/]")
        config.activate_account(args.account)
    account.refresh_account_paths()

    if args.domain == "auth":
        account.cmd_auth(args)
    elif hasattr(args, "func"):
        args.func(args)
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
