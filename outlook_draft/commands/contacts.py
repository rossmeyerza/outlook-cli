from __future__ import annotations

import argparse
import json
from typing import Any

from rich.table import Table
from rich.text import Text

from .. import runtime
from ..cache import CONTACT_CACHE
from ..directory import organisation_view, validate_org_options
from ..errors import CliError, OutlookAPIError, OutlookDraftError
from ..parser import add_output_args
from ..progress import spinner
from ..references import _resolve_contact_id


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
    client = runtime.get_client()
    try:
        with spinner(args, f"Searching contacts for '{args.query}'..."):
            people = client.search_people(args.query, top=args.count)
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to search contacts: {e}[/]") from e
    finally:
        client.close()

    if not people:
        if args.json:
            print(json.dumps([]))
            return
        runtime.console.print(f"[dim]No contacts found for '{args.query}'.[/]")
        return

    CONTACT_CACHE.write_text(json.dumps([{"Id": _contact_id(p)} for p in people if _contact_id(p)]))

    if args.json:
        print(json.dumps([
            {
                "index": i,
                "id": _contact_id(person),
                "displayName": person.get("DisplayName", ""),
                "emails": [
                    email.get("Address", "")
                    for email in person.get("ScoredEmailAddresses", [])
                    if email.get("Address")
                ],
                "type": _type_label(person),
            }
            for i, person in enumerate(people, 1)
        ], indent=2, ensure_ascii=False))
        return

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

    runtime.console.print(table)


def _org_text(value) -> Text:
    return Text(str(value) if value is not None and value != "" else "Not recorded or visible")


def _print_organisation(result: dict[str, Any]) -> None:
    person = result["person"]
    runtime.console.print(Text("Company: ") + _org_text(person["company"]))
    runtime.console.print(Text("Department: ") + _org_text(person["department"]))
    table = Table(title="Organisation details")
    for column in ("Relationship", "Name", "Job title", "Email"):
        table.add_column(column, overflow="fold")
    rows = [("Person", person), ("Manager", result["manager"])]
    rows.extend(("Direct report", report) for report in result["directReports"])
    rows.extend((f"Manager level {level}", manager) for level, manager in enumerate((result["managementChain"] or [])[1:], 2))
    for relationship, entry in rows:
        if entry:
            table.add_row(Text(relationship), _org_text(entry["displayName"] or entry["id"]), _org_text(entry["jobTitle"]), _org_text(entry["email"]))
        else:
            label = "Unavailable" if result["managerStatus"] == "unavailable" else "Not recorded or visible"
            table.add_row(Text(relationship), Text(label), Text(""), Text(""))
    runtime.console.print(table)
    suffix = "" if result["directReportsComplete"] else " returned; the list is incomplete"
    runtime.console.print(Text(f"Direct reports: {result['directReportsCount']}{suffix}"))
    if result["managementChainComplete"] is False:
        runtime.console.print(Text("The management chain is incomplete.", style="yellow"))
    for warning in result["warnings"]:
        runtime.console.print(Text(f"Warning: {warning}", style="yellow"))


def cmd_contact_org(args: argparse.Namespace) -> None:
    """Show the available organisation details without a directory change."""
    try:
        user = validate_org_options(args.user, args.chain, args.max_depth)
    except ValueError as exc:
        raise CliError(str(exc)) from exc
    client = runtime.get_graph_client()
    try:
        with spinner(args, "Read organisation details..."):
            result = organisation_view(client, user, args.chain, args.max_depth)
    except (OutlookDraftError, ValueError) as exc:
        raise CliError(f"Cannot read organisation details: {exc}") from exc
    finally:
        client.close()
    if args.json:
        runtime.print_json(result)
    else:
        _print_organisation(result)


def cmd_contact_create(args: argparse.Namespace) -> None:
    client = runtime.get_client()
    try:
        with spinner(args, "Creating contact..."):
            contact = client.create_contact(
                display_name=args.name,
                email=args.email,
                given_name=args.given_name,
                surname=args.surname,
                company=args.company,
                mobile_phone=args.mobile,
            )
        runtime.console.print(f"[green]Created contact:[/] {contact.get('DisplayName', args.name)}")
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to create contact: {e}[/]") from e
    finally:
        client.close()


def cmd_contact_update(args: argparse.Namespace) -> None:
    if not any([args.name, args.email, args.given_name, args.surname, args.company, args.mobile]):
        raise CliError("[red]Provide at least one field to update.[/]")
    client = runtime.get_client()
    contact_id = _resolve_contact_id(args.contact_id)
    try:
        with spinner(args, "Updating contact..."):
            client.update_contact(
                contact_id,
                display_name=args.name,
                email=args.email,
                given_name=args.given_name,
                surname=args.surname,
                company=args.company,
                mobile_phone=args.mobile,
            )
        runtime.console.print("[green]Contact updated.[/]")
    except OutlookAPIError as e:
        raise CliError(f"[red]Failed to update contact: {e}[/]") from e
    finally:
        client.close()


def register(sub) -> None:
    p_contact = sub.add_parser("contact", help="Manage contacts")
    add_output_args(p_contact)
    contact_sub = p_contact.add_subparsers(dest="command", required=True)

    cmd_contact_search = contact_sub.add_parser("search", help="Search contacts by name or email")
    add_output_args(cmd_contact_search)
    cmd_contact_search.add_argument("query", help="Name or email to search for")
    cmd_contact_search.add_argument("--count", "-n", type=int, default=10, help="Max results")
    cmd_contact_search.set_defaults(func=cmd_contacts)

    cmd_org = contact_sub.add_parser("org", help="Read company, manager and direct reports from the directory")
    add_output_args(cmd_org)
    cmd_org.add_argument("user", nargs="?", default="me", help="Email, UPN, directory user ID, unique name or me")
    cmd_org.add_argument("--chain", action="store_true", help="Include the recorded management chain")
    cmd_org.add_argument("--max-depth", type=int, help="Limit the management chain; requires --chain")
    cmd_org.set_defaults(func=cmd_contact_org)

    cmd_contact_create_parser = contact_sub.add_parser("create", help="Create a personal contact")
    cmd_contact_create_parser.add_argument("--name", required=True, help="Display name")
    cmd_contact_create_parser.add_argument("--email", required=True, help="Email address")
    cmd_contact_create_parser.add_argument("--given-name", help="Given name")
    cmd_contact_create_parser.add_argument("--surname", help="Surname")
    cmd_contact_create_parser.add_argument("--company", help="Company name")
    cmd_contact_create_parser.add_argument("--mobile", help="Mobile phone")
    cmd_contact_create_parser.set_defaults(func=cmd_contact_create)

    cmd_contact_update_parser = contact_sub.add_parser("update", help="Update a personal contact")
    cmd_contact_update_parser.add_argument("contact_id", help="Contact index from search or full ID")
    cmd_contact_update_parser.add_argument("--name", help="Display name")
    cmd_contact_update_parser.add_argument("--email", help="Email address")
    cmd_contact_update_parser.add_argument("--given-name", help="Given name")
    cmd_contact_update_parser.add_argument("--surname", help="Surname")
    cmd_contact_update_parser.add_argument("--company", help="Company name")
    cmd_contact_update_parser.add_argument("--mobile", help="Mobile phone")
    cmd_contact_update_parser.set_defaults(func=cmd_contact_update)
