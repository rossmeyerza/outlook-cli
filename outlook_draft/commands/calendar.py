from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime
from typing import Any, Callable

from rich.panel import Panel
from rich.table import Table

from ..calendar_time import outlook_datetime, parse_local_datetime
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


def _resolve_cal_id(args: argparse.Namespace, client: Any, ref: str) -> str:
    return _ctx(args)["resolve_cal_id"](client, ref)


def _parse_event_dt(value: str) -> datetime | None:
    if not value:
        return None
    try:
        return datetime.fromisoformat(value[:23])
    except Exception:
        return None


def _agenda_row(event: JsonDict) -> tuple[str, str]:
    start = event.get("Start", {}).get("DateTime", "")
    end = event.get("End", {}).get("DateTime", "")
    parsed_start = _parse_event_dt(start)
    parsed_end = _parse_event_dt(end)
    if not parsed_start:
        return start[:10], start[11:16]

    date_str = parsed_start.strftime("%a %b %d")
    if event.get("IsAllDay", False):
        return date_str, "All Day"
    if parsed_end:
        return date_str, f"{parsed_start.strftime('%H:%M')}-{parsed_end.strftime('%H:%M')}"
    return date_str, parsed_start.strftime("%H:%M")


def _event_when(event: JsonDict) -> str:
    start = event.get("Start", {}).get("DateTime", "")
    end = event.get("End", {}).get("DateTime", "")
    parsed_start = _parse_event_dt(start)
    parsed_end = _parse_event_dt(end)
    if not parsed_start:
        return start
    if event.get("IsAllDay"):
        return f"{parsed_start.strftime('%A, %b %d, %Y')} (All Day)"
    if parsed_end:
        return f"{parsed_start.strftime('%A, %b %d, %Y %H:%M')} - {parsed_end.strftime('%H:%M')}"
    return parsed_start.strftime("%A, %b %d, %Y %H:%M")


def _event_recurrence(event: JsonDict) -> dict[str, Any]:
    recurrence = event.get("Recurrence") or {}
    pattern = recurrence.get("Pattern") or {}
    range_info = recurrence.get("Range") or {}
    return {
        "type": event.get("Type", ""),
        "seriesMasterId": event.get("SeriesMasterId", ""),
        "pattern": pattern,
        "range": range_info,
    }


def _event_summary(event: JsonDict, index: int | None = None) -> dict[str, Any]:
    date_str, time_str = _agenda_row(event)
    organizer = event.get("Organizer", {}).get("EmailAddress", {})
    return {
        "index": index,
        "id": event.get("Id", ""),
        "subject": event.get("Subject", ""),
        "date": date_str,
        "time": time_str,
        "start": event.get("Start", {}),
        "end": event.get("End", {}),
        "location": event.get("Location", {}).get("DisplayName", ""),
        "organizer": {
            "name": organizer.get("Name", ""),
            "address": organizer.get("Address", ""),
        },
        "isAllDay": event.get("IsAllDay", False),
        "isCancelled": event.get("IsCancelled", False),
        "recurrence": _event_recurrence(event),
    }


def _print_json(value: Any) -> None:
    print(json.dumps(value, indent=2, ensure_ascii=False))


def _print_plain_agenda(args: argparse.Namespace, events: list[JsonDict]) -> None:
    for i, event in enumerate(events, 1):
        summary = _event_summary(event, i)
        cancelled = " [cancelled]" if summary["isCancelled"] else ""
        location = f" | {summary['location']}" if summary["location"] else ""
        _console(args).print(
            f"{i}. {summary['date']} {summary['time']} | {summary['subject'] or '(no subject)'}"
            f"{location}{cancelled}"
        )


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

    save_cache(CAL_CACHE, events)

    if args.json:
        _print_json([_event_summary(event, i) for i, event in enumerate(events, 1)])
        return

    if not events:
        _console(args).print("[dim]No upcoming events.[/]")
        return

    if args.plain:
        _print_plain_agenda(args, events)
        return

    table = Table(title=f"Agenda (next {args.days} days)")
    table.add_column("#", style="dim", width=3)
    table.add_column("Date", width=12)
    table.add_column("Time", width=12)
    table.add_column("Subject", ratio=2, min_width=22)
    table.add_column("Location", ratio=1, min_width=18)
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

    if args.json:
        data = _event_summary(event)
        data["attendees"] = event.get("Attendees", [])
        data["body"] = event.get("Body", {})
        _print_json(data)
        return

    organizer = event.get("Organizer", {}).get("EmailAddress", {})
    attendees = event.get("Attendees", [])
    recurrence = _event_recurrence(event)
    lines = [
        f"[bold]Subject:[/] {event.get('Subject', '')}",
        f"[bold]When:[/] {_event_when(event)}",
        f"[bold]Where:[/] {event.get('Location', {}).get('DisplayName', '')}",
        f"[bold]Organizer:[/] {organizer.get('Name', '')} <{organizer.get('Address', '')}>",
    ]
    if recurrence["type"]:
        lines.append(f"[bold]Type:[/] {recurrence['type']}")
    if recurrence["seriesMasterId"]:
        lines.append(f"[bold]Series master ID:[/] [dim]{recurrence['seriesMasterId']}[/]")
    if recurrence["pattern"]:
        lines.append(f"[bold]Recurrence:[/] {recurrence['pattern']}")
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


def _parse_event_datetimes(args: argparse.Namespace) -> tuple[str, str]:
    try:
        start_dt = parse_local_datetime(args.start)
        end_dt = parse_local_datetime(args.end)
    except Exception as e:
        _console(args).print(f"[red]Failed to parse dates: {e}[/]")
        _console(args).print("Try formats like: '2026-04-10 14:00' or 'tomorrow 2pm'")
        sys.exit(1)
    return outlook_datetime(start_dt), outlook_datetime(end_dt)


def cmd_cal_create(args: argparse.Namespace) -> None:
    """Create a calendar event."""
    start_local, end_local = _parse_event_datetimes(args)

    client = _get_client(args)
    try:
        client.create_event(
            subject=args.subject,
            start_dt=start_local,
            end_dt=end_local,
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


def _availability_summary(item: JsonDict) -> str:
    schedule = item.get("ScheduleId", "")
    error = item.get("Error", {})
    if error:
        return f"{schedule}: ERROR {error.get('Message', error)}"
    return f"{schedule}: {item.get('AvailabilityView', '')}"


def cmd_cal_rooms(args: argparse.Namespace) -> None:
    client = _get_client(args)
    try:
        rooms = client.find_rooms(room_list=args.room_list)
    except OutlookAPIError as e:
        message = str(e)
        if "findRooms" in message and "Resource not found" in message:
            _console(args).print(
                "[yellow]Room discovery is not available through this Outlook REST token/tenant.[/] "
                "Use `cal availability` or `cal find-time` with known room email addresses instead."
            )
        else:
            _console(args).print(f"[red]Failed to find rooms: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
    if args.json:
        _print_json(rooms)
        return
    table = Table(title=f"Rooms ({len(rooms)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Name", ratio=1)
    table.add_column("Address", ratio=1)
    for i, room in enumerate(rooms, 1):
        table.add_row(str(i), room.get("Name", ""), room.get("Address", ""))
    _console(args).print(table)


def cmd_cal_availability(args: argparse.Namespace) -> None:
    start_local, end_local = _parse_event_datetimes(args)
    client = _get_client(args)
    try:
        availability = client.get_schedule(
            schedules=args.attendee,
            start_dt=start_local,
            end_dt=end_local,
            interval_minutes=args.interval,
        )
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to get availability: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
    if args.json:
        _print_json(availability)
        return
    for item in availability:
        _console(args).print(_availability_summary(item))


def cmd_cal_find_time(args: argparse.Namespace) -> None:
    start_local, end_local = _parse_event_datetimes(args)
    client = _get_client(args)
    try:
        suggestions = client.find_meeting_times(
            attendees=args.attendee,
            start_dt=start_local,
            end_dt=end_local,
            duration_minutes=args.duration,
            max_candidates=args.count,
        )
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to find meeting times: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
    if args.json:
        _print_json(suggestions)
        return
    table = Table(title=f"Meeting suggestions ({len(suggestions)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Start", ratio=1)
    table.add_column("End", ratio=1)
    table.add_column("Confidence", width=12)
    for i, suggestion in enumerate(suggestions, 1):
        slot = suggestion.get("MeetingTimeSlot", {})
        table.add_row(
            str(i),
            slot.get("Start", {}).get("DateTime", ""),
            slot.get("End", {}).get("DateTime", ""),
            str(suggestion.get("Confidence", "")),
        )
    _console(args).print(table)


def cmd_cal_update(args: argparse.Namespace) -> None:
    """Update a calendar event."""
    if not any([args.subject, args.start, args.end, args.location is not None, args.body is not None]):
        _console(args).print("[red]Provide at least one field to update.[/]")
        sys.exit(1)
    if bool(args.start) != bool(args.end):
        _console(args).print("[red]Provide both --start and --end when changing time.[/]")
        sys.exit(1)

    start_local = end_local = None
    if args.start and args.end:
        start_local, end_local = _parse_event_datetimes(args)

    client = _get_client(args)
    event_id = _resolve_cal_id(args, client, args.event_id)
    try:
        event = client.get_event(event_id)
        client.update_event(
            event_id,
            subject=args.subject,
            start_dt=start_local,
            end_dt=end_local,
            location=args.location,
            body=args.body,
        )
        _console(args).print(f"[green]Updated event:[/] {event.get('Subject', event_id)}")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to update event: {e}[/]")
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


def cmd_cal_tentative(args: argparse.Namespace) -> None:
    """Tentatively accept a calendar event."""
    _respond_to_event(args, "tentativelyAccept", "Tentatively accepted")


def cmd_cal_cancel(args: argparse.Namespace) -> None:
    """Cancel a calendar event as organizer."""
    client = _get_client(args)
    event_id = _resolve_cal_id(args, client, args.event_id)
    try:
        event = client.get_event(event_id)
        client.cancel_event(event_id, comment=args.comment or "")
        _console(args).print(f"[green]Cancelled event:[/] {event.get('Subject', event_id)}")
    except OutlookAPIError as e:
        _console(args).print(f"[red]Failed to cancel event: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
