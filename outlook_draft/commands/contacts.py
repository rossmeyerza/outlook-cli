from __future__ import annotations

import argparse
import sys
from typing import Any, Callable

from rich.table import Table

from ..errors import OutlookAPIError


def build_ctx(*, console: Any, get_client: Callable[[], Any]) -> dict[str, Any]:
    return {
        "console": console,
        "get_client": get_client,
    }


def _ctx(args: argparse.Namespace) -> dict[str, Any]:
    return args._contacts_ctx


def _console(args: argparse.Namespace):
    return _ctx(args)["console"]


def _get_client(args: argparse.Namespace):
    return _ctx(args)["get_client"]()


def _type_label(person: dict[str, Any]) -> str:
    ptype = person.get("PersonType", {}).get("Subclass", "")
    return {
        "OrganizationUser": "Org directory",
        "ImplicitContact": "Recent contact",
        "PersonalContact": "Personal",
    }.get(ptype, ptype)


def cmd_contacts(args: argparse.Namespace) -> None:
    """Search contacts and directory by name or email."""
    client = _get_client(args)
    try:
        people = client.search_people(args.query, top=args.count)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to search contacts: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not people:
        _console(args).print(f"[dim]No contacts found for '{args.query}'.[/]")
        return

    table = Table(title=f"Contacts matching '{args.query}'")
    table.add_column("#", style="dim", width=3)
    table.add_column("Name", width=30)
    table.add_column("Email", ratio=1)
    table.add_column("Type", style="dim", width=16)

    for i, person in enumerate(people, 1):
        emails = person.get("ScoredEmailAddresses", [])
        email_str = ", ".join(e.get("Address", "") for e in emails)
        table.add_row(
            str(i),
            person.get("DisplayName", ""),
            email_str,
            _type_label(person),
        )

    _console(args).print(table)
