"""Outlook CLI entry point and command dispatch."""
from __future__ import annotations

import argparse
import sys

from . import config
from .commands import account, drafts, mail, tasks, calendar, contacts, teams, files
from .commands import signature, gateway_parser, mcp, completion
from .parser import add_output_args
from .runtime import console

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

def main() -> None:
    argv, account_override = _extract_account_arg(sys.argv[1:])
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
    signature.register(sub)
    gateway_parser.register(sub)
    account.register_config(sub)
    mcp.register(sub)
    completion.register(sub)

    # Let argcomplete service tab-completion requests for bash/zsh/fish.
    try:
        import argcomplete
    except ImportError:
        pass
    else:
        argcomplete.autocomplete(parser)

    args = parser.parse_args(argv)

    if args.account:
        if not config.account_exists(args.account):
            console.print(f"[red]No account profile named '{args.account}'.[/]")
            sys.exit(1)
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
