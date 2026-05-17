from __future__ import annotations

import argparse
import sys
from datetime import datetime, timezone
from typing import Any, Callable

import dateutil.parser
from rich.panel import Panel
from rich.table import Table

from ..cache import CAL_CACHE, save_cache
from ..errors import OutlookAPIError


JsonDict = dict[str, Any]


def build_ctx(
    *,
    console: Any,
    get_client: Callable[[], Any],
    html_to_text: Callable[[str], str],
    format_datetime: Callable[[str], str],
    resolve_cal_id: Callable[[Any, str], str],
) -> dict[str, Any]:
    return {
        "console": console,
        "get_client": get_client,
        "html_to_text": html_to_text,
        "format_datetime": format_datetime,
        "resolve_cal_id": resolve_cal_id,
    }


def _ctx(args: argparse.Namespace) -> dict[str, Any]:
    return args._calendar_ctx


def _console(args: argparse.Namespace):
    return _ctx(args)["console"]


def _get_client(args: argparse.Namespace):
    return _ctx(args)["get_client"]()


def _html_to_text(args: argparse.Namespace, html: str) -> str:
    return _ctx(args)["html_to_text"](html)


def _format_datetime(args: argparse.Namespace, value: str) -> str:
    return _ctx(args)["format_datetime"](value)


def _resolve_cal_id(args: argparse.Namespace, client: Any, ref: str) -> str:
    return _ctx(args)["resolve_cal_id"](client, ref)


def _agenda_row(event: JsonDict) -> tuple[str, str]:
    start = event.get("Start", {}).get("DateTime", "")
    end = event.get("End", {}).get("DateTime", "")
    if not start:
        return "", ""

    try:
        parsed_start = datetime.fromisoformat(start[:23])
        parsed_end = datetime.fromisoformat(end[:23])
        date_str = parsed_start.strftime("%a %b %d")
        if event.get("IsAllDay", False):
            return date_str, "All Day"
        return date_str, f"{parsed_start.strftime('%H:%M')}-{parsed_end.strftime('%H:%M')}"
    except Exception:
        return start[:10], start[11:16]


def _event_when(event: JsonDict) -> str:
    start = event.get("Start", {}).get("DateTime", "")
    end = event.get("End", {}).get("DateTime", "")
    if not start:
        return ""

    try:
        parsed_start = datetime.fromisoformat(start[:23])
        parsed_end = datetime.fromisoformat(end[:23])
        if event.get("IsAllDay"):
            return f"{parsed_start.strftime('%A, %b %d, %Y')} (All Day)"
        return f"{parsed_start.strftime('%A, %b %d, %Y %H:%M')} - {parsed_end.strftime('%H:%M')}"
    except Exception:
        return _format_datetime_placeholder(start)


def _format_datetime_placeholder(value: str) -> str:
    return value


def cmd_cal_agenda(args: argparse.Namespace) -> None:
    """List upcoming calendar events."""
    client = _get_client(args)
    try:
        events = client.get_agenda(days=args.days, top=args.count)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to fetch calendar: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if not events:
        _console(args).print("[dim]No upcoming events.[/]")
        return

    save_cache(CAL_CACHE, events)

    table = Table(title=f"Agenda (next {args.days} days)")
    table.add_column("#", style="dim", width=3)
    table.add_column("Date", width=12)
    table.add_column("Time", width=12)
    table.add_column("Subject", ratio=1)
    table.add_column("Location", width=25)
    table.add_column("Status", width=10)

    for i, event in enumerate(events, 1):
        date_str, time_str = _agenda_row(event)
        status = "[red]Cancelled[/]" if event.get("IsCancelled", False) else ""
        style = "dim" if event.get("IsCancelled", False) else ""
        table.add_row(
            str(i),
            date_str,
            time_str,
            event.get("Subject", "(no subject)"),
            event.get("Location", {}).get("DisplayName", ""),
            status,
            style=style,
        )

    _console(args).print(table)
    _console(args).print("[dim]Use 'outlook-cli cal show <n>' to see event details.[/]")


def cmd_cal_show(args: argparse.Namespace) -> None:
    """Show event details."""
    client = _get_client(args)
    event_id = _resolve_cal_id(args, client, args.event_id)
    try:
        event = client.get_event(event_id)
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to fetch event: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    organizer = event.get("Organizer", {}).get("EmailAddress", {})
    attendees = event.get("Attendees", [])
    lines = [
        f"[bold]Subject:[/] {event.get('Subject', '')}",
        f"[bold]When:[/] {_event_when(event)}",
        f"[bold]Where:[/] {event.get('Location', {}).get('DisplayName', '')}",
        f"[bold]Organizer:[/] {organizer.get('Name', '')} <{organizer.get('Address', '')}>",
    ]
    if event.get("IsCancelled"):
        lines.insert(0, "[red bold]*** CANCELLED ***[/]")

    if attendees:
        attendee_lines = []
        for attendee in attendees:
            email = attendee.get("EmailAddress", {})
            name = email.get("Name", email.get("Address", ""))
            response = attendee.get("Status", {}).get("Response", "None")
            if response == "Accepted":
                marker = "[green]✓[/]"
            elif response == "Declined":
                marker = "[red]✗[/]"
            elif response == "TentativelyAccepted":
                marker = "[yellow]?[/]"
            else:
                marker = "[dim]-[/]"
            attendee_lines.append(f"{marker} {name}")
        lines.append(f"[bold]Attendees:[/] {', '.join(attendee_lines)}")

    body = event.get("Body", {})
    content = body.get("Content", "")
    if body.get("ContentType") == "HTML":
        content = _html_to_text(args, content)

    lines.append("")
    lines.append(content[:3000])
    _console(args).print(Panel("\n".join(lines), title="Event", border_style="magenta"))


def cmd_cal_create(args: argparse.Namespace) -> None:
    """Create a calendar event."""
    try:
        start_dt = dateutil.parser.parse(args.start)
        end_dt = dateutil.parser.parse(args.end)
    except Exception as e:
        _console(args).print(f"[red]Failed to parse dates: {e}[/]")
        _console(args).print("Try formats like: '2026-04-10 14:00' or 'tomorrow 2pm'")
        sys.exit(1)

    if start_dt.tzinfo is None:
        start_dt = start_dt.astimezone()
    if end_dt.tzinfo is None:
        end_dt = end_dt.astimezone()

    start_utc = start_dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")
    end_utc = end_dt.astimezone(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S")

    client = _get_client(args)
    try:
        client.create_event(
            subject=args.subject,
            start_dt=start_utc,
            end_dt=end_utc,
            location=args.location,
            body=args.body,
            attendees=args.attendees,
        )
        _console(args).print(f"[green]Created event:[/] {args.subject}")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to create event: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_cal_delete(args: argparse.Namespace) -> None:
    """Delete a calendar event."""
    client = _get_client(args)
    event_id = _resolve_cal_id(args, client, args.event_id)
    try:
        event = client.get_event(event_id)
        client.delete_event(event_id)
        _console(args).print(f"[green]Deleted event:[/] {event.get('Subject', event_id)}")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to delete event: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def _respond_to_event(args: argparse.Namespace, response: str, label: str) -> None:
    client = _get_client(args)
    event_id = _resolve_cal_id(args, client, args.event_id)
    try:
        event = client.get_event(event_id)
        client.respond_to_event(
            event_id,
            response,
            comment=args.comment or "",
            send_response=args.send_response,
        )
        _console(args).print(f"[green]{label} event:[/] {event.get('Subject', event_id)}")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to {label.lower()} event: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_cal_accept(args: argparse.Namespace) -> None:
    """Accept a calendar event."""
    _respond_to_event(args, "accept", "Accepted")


def cmd_cal_decline(args: argparse.Namespace) -> None:
    """Decline a calendar event."""
    _respond_to_event(args, "decline", "Declined")
