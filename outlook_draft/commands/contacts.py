from __future__ import annotations

import argparse
import sys
from typing import Any, Callable

from rich.table import Table

import json

from ..cache import CONTACT_CACHE
from ..errors import OutlookAPIError


def build_ctx(*, console: Any, get_client: Callable[[], Any], resolve_contact_id: Callable[[str], str] | None = None) -> dict[str, Any]:
    return {
        "console": console,
        "get_client": get_client,
        "resolve_contact_id": resolve_contact_id,
    }


def _ctx(args: argparse.Namespace) -> dict[str, Any]:
    return args._contacts_ctx


def _console(args: argparse.Namespace):
    return _ctx(args)["console"]


def _get_client(args: argparse.Namespace):
    return _ctx(args)["get_client"]()


def _resolve_contact_id(args: argparse.Namespace, ref: str) -> str:
    resolver = _ctx(args).get("resolve_contact_id")
    return resolver(ref) if resolver else ref


def _type_label(person: dict[str, Any]) -> str:
    ptype = person.get("PersonType", {}).get("Subclass", "")
    return {
        "OrganizationUser": "Org directory",
        "ImplicitContact": "Recent contact",
        "PersonalContact": "Personal",
    }.get(ptype, ptype)


def _contact_id(person: dict[str, Any]) -> str:
    return person.get("Id") or person.get("id") or ""


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

    CONTACT_CACHE.write_text(json.dumps([{"Id": _contact_id(p)} for p in people if _contact_id(p)]))

    table = Table(title=f"Contacts matching '{args.query}'")
    table.add_column("#", style="dim", width=3)
    table.add_column("Name", ratio=1, min_width=18)
    table.add_column("Email", ratio=2, min_width=28, overflow="fold")
    table.add_column("Type", style="dim", width=16)
    table.add_column("ID", style="dim", width=16, no_wrap=True)

    for i, person in enumerate(people, 1):
        emails = person.get("ScoredEmailAddresses", [])
        email_str = ", ".join(e.get("Address", "") for e in emails)
        cid = _contact_id(person)
        table.add_row(
            str(i),
            person.get("DisplayName", ""),
            email_str,
            _type_label(person),
            cid[-16:],
        )

    _console(args).print(table)


def cmd_contact_create(args: argparse.Namespace) -> None:
    client = _get_client(args)
    try:
        contact = client.create_contact(
            display_name=args.name,
            email=args.email,
            given_name=args.given_name,
            surname=args.surname,
            company=args.company,
            mobile_phone=args.mobile,
        )
        _console(args).print(f"[green]Created contact:[/] {contact.get('DisplayName', args.name)}")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to create contact: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_contact_update(args: argparse.Namespace) -> None:
    if not any([args.name, args.email, args.given_name, args.surname, args.company, args.mobile]):
        _console(args).print("[red]Provide at least one field to update.[/]")
        sys.exit(1)
    client = _get_client(args)
    contact_id = _resolve_contact_id(args, args.contact_id)
    try:
        client.update_contact(
            contact_id,
            display_name=args.name,
            email=args.email,
            given_name=args.given_name,
            surname=args.surname,
            company=args.company,
            mobile_phone=args.mobile,
        )
        _console(args).print("[green]Contact updated.[/]")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to update contact: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
