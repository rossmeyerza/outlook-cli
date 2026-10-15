from __future__ import annotations

from ..parser import add_output_args, _add_yes_arg


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
from ..progress import spinner
from .. import runtime
from ..formatting import html_to_text, format_datetime
from ..references import _resolve_message_id, _resolve_cal_id, _resolve_task_id, _resolve_contact_id


JsonDict = dict[str, Any]


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


def _print_plain_agenda(args: argparse.Namespace, events: list[JsonDict]) -> None:
    for i, event in enumerate(events, 1):
        summary = _event_summary(event, i)
        cancelled = " [cancelled]" if summary["isCancelled"] else ""
        location = f" | {summary['location']}" if summary["location"] else ""
        runtime.console.print(
            f"{i}. {summary['date']} {summary['time']} | {summary['subject'] or '(no subject)'}"
            f"{location}{cancelled}"
        )


def cmd_cal_agenda(args: argparse.Namespace) -> None:
    """List upcoming calendar events."""
    client = runtime.get_client()
    try:
        with spinner(args, "Loading calendar agenda..."):
            events = client.get_agenda(days=args.days, top=args.count)
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to fetch calendar: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    save_cache(CAL_CACHE, events)

    if args.json:
        runtime.print_json([_event_summary(event, i) for i, event in enumerate(events, 1)])
        return

    if not events:
        runtime.console.print("[dim]No upcoming events.[/]")
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

    runtime.console.print(table)
    runtime.console.print("[dim]Use 'outlook-cli cal show <n>' to see event details.[/]")


def cmd_cal_show(args: argparse.Namespace) -> None:
    """Show event details."""
    client = runtime.get_client()
    event_id = _resolve_cal_id(client, args.event_id)
    try:
        with spinner(args, "Loading calendar event..."):
            event = client.get_event(event_id)
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to fetch event: {e}[/]")
        sys.exit(1)
    finally:
        client.close()

    if args.json:
        data = _event_summary(event)
        data["attendees"] = event.get("Attendees", [])
        data["body"] = event.get("Body", {})
        runtime.print_json(data)
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
        content = html_to_text(content)

    lines.append("")
    lines.append(content[:3000])
    runtime.console.print(Panel("\n".join(lines), title="Event", border_style="magenta"))


def _parse_event_datetimes(args: argparse.Namespace) -> tuple[str, str]:
    try:
        start_dt = parse_local_datetime(args.start)
        end_dt = parse_local_datetime(args.end)
    except Exception as e:
        runtime.console.print(f"[red]Failed to parse dates: {e}[/]")
        runtime.console.print("Try formats like: '2026-04-10 14:00' or 'tomorrow 2pm'")
        sys.exit(1)
    return outlook_datetime(start_dt), outlook_datetime(end_dt)


def cmd_cal_create(args: argparse.Namespace) -> None:
    """Create a calendar event."""
    start_local, end_local = _parse_event_datetimes(args)

    client = runtime.get_client()
    try:
        with spinner(args, "Creating calendar event..."):
            client.create_event(
                subject=args.subject,
                start_dt=start_local,
                end_dt=end_local,
                location=args.location,
                body=args.body,
                attendees=args.attendees,
                is_online_meeting=args.teams,
            )
        runtime.console.print(f"[green]Created event:[/] {args.subject}")
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to create event: {e}[/]")
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
    client = runtime.get_client()
    try:
        with spinner(args, "Finding rooms..."):
            rooms = client.find_rooms(room_list=args.room_list)
    except OutlookAPIError as e:
        message = str(e)
        if "findRooms" in message and "Resource not found" in message:
            runtime.console.print(
                "[yellow]Room discovery is not available through this Outlook REST token/tenant.[/] "
                "Use `cal availability` or `cal find-time` with known room email addresses instead."
            )
        else:
            runtime.console.print(f"[red]Failed to find rooms: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
    if args.json:
        runtime.print_json(rooms)
        return
    table = Table(title=f"Rooms ({len(rooms)})")
    table.add_column("#", style="dim", width=3)
    table.add_column("Name", ratio=1)
    table.add_column("Address", ratio=1)
    for i, room in enumerate(rooms, 1):
        table.add_row(str(i), room.get("Name", ""), room.get("Address", ""))
    runtime.console.print(table)


def cmd_cal_availability(args: argparse.Namespace) -> None:
    start_local, end_local = _parse_event_datetimes(args)
    client = runtime.get_client()
    try:
        with spinner(args, "Checking availability..."):
            availability = client.get_schedule(
                schedules=args.attendee,
                start_dt=start_local,
                end_dt=end_local,
                interval_minutes=args.interval,
            )
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to get availability: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
    if args.json:
        runtime.print_json(availability)
        return
    for item in availability:
        runtime.console.print(_availability_summary(item))


def cmd_cal_find_time(args: argparse.Namespace) -> None:
    start_local, end_local = _parse_event_datetimes(args)
    client = runtime.get_client()
    try:
        with spinner(args, "Finding meeting times..."):
            suggestions = client.find_meeting_times(
                attendees=args.attendee,
                start_dt=start_local,
                end_dt=end_local,
                duration_minutes=args.duration,
                max_candidates=args.count,
            )
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to find meeting times: {e}[/]")
        sys.exit(1)
    finally:
        client.close()
    if args.json:
        runtime.print_json(suggestions)
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
    runtime.console.print(table)


def cmd_cal_update(args: argparse.Namespace) -> None:
    """Update a calendar event."""
    if args.teams is None:
        teams_change = False
    else:
        teams_change = True
    if not any([
        args.subject,
        args.start,
        args.end,
        args.location is not None,
        args.body is not None,
        teams_change,
    ]):
        runtime.console.print("[red]Provide at least one field to update.[/]")
        sys.exit(1)
    if bool(args.start) != bool(args.end):
        runtime.console.print("[red]Provide both --start and --end when changing time.[/]")
        sys.exit(1)

    start_local = end_local = None
    if args.start and args.end:
        start_local, end_local = _parse_event_datetimes(args)

    client = runtime.get_client()
    event_id = _resolve_cal_id(client, args.event_id)
    try:
        with spinner(args, "Updating calendar event..."):
            event = client.get_event(event_id)
            client.update_event(
                event_id,
                subject=args.subject,
                start_dt=start_local,
                end_dt=end_local,
                location=args.location,
                body=args.body,
                is_online_meeting=args.teams,
            )
        runtime.console.print(f"[green]Updated event:[/] {event.get('Subject', event_id)}")
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to update event: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def cmd_cal_delete(args: argparse.Namespace) -> None:
    """Delete a calendar event."""
    client = runtime.get_client()
    event_id = _resolve_cal_id(client, args.event_id)
    try:
        with spinner(args, "Deleting calendar event..."):
            event = client.get_event(event_id)
            client.delete_event(event_id)
        runtime.console.print(f"[green]Deleted event:[/] {event.get('Subject', event_id)}")
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to delete event: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def _respond_to_event(args: argparse.Namespace, response: str, label: str) -> None:
    client = runtime.get_client()
    event_id = _resolve_cal_id(client, args.event_id)
    try:
        with spinner(args, f"{label} calendar event..."):
            event = client.get_event(event_id)
            client.respond_to_event(
                event_id,
                response,
                comment=args.comment or "",
                send_response=args.send_response,
            )
        runtime.console.print(f"[green]{label} event:[/] {event.get('Subject', event_id)}")
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to {label.lower()} event: {e}[/]")
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
    client = runtime.get_client()
    event_id = _resolve_cal_id(client, args.event_id)
    try:
        with spinner(args, "Cancelling calendar event..."):
            event = client.get_event(event_id)
            client.cancel_event(event_id, comment=args.comment or "")
        runtime.console.print(f"[green]Cancelled event:[/] {event.get('Subject', event_id)}")
    except OutlookAPIError as e:
        runtime.console.print(f"[red]Failed to cancel event: {e}[/]")
        sys.exit(1)
    finally:
        client.close()


def register(sub) -> None:
    p_cal = sub.add_parser("cal", help="Manage calendar events")
    add_output_args(p_cal)
    cal_sub = p_cal.add_subparsers(dest="command", required=True)

    cmd_cal_agenda_parser = cal_sub.add_parser("agenda", help="List upcoming calendar events")
    add_output_args(cmd_cal_agenda_parser)
    cmd_cal_agenda_parser.add_argument("--days", "-d", type=int, default=7, help="Days to look ahead")
    cmd_cal_agenda_parser.add_argument("--count", "-n", type=int, default=20, help="Max events")
    cmd_cal_agenda_parser.add_argument("--plain", action="store_true", help="Output plain text instead of a table")
    cmd_cal_agenda_parser.set_defaults(func=cmd_cal_agenda)

    cmd_cal_show_parser = cal_sub.add_parser("show", help="Show event details")
    add_output_args(cmd_cal_show_parser)
    cmd_cal_show_parser.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_show_parser.set_defaults(func=cmd_cal_show)

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
    cmd_cal_create_cmd.set_defaults(func=cmd_cal_create)

    cmd_cal_rooms_cmd = cal_sub.add_parser("rooms", help="Find rooms")
    add_output_args(cmd_cal_rooms_cmd)
    cmd_cal_rooms_cmd.add_argument("--room-list", help="Optional room list email/address")
    cmd_cal_rooms_cmd.set_defaults(func=cmd_cal_rooms)

    cmd_cal_availability_cmd = cal_sub.add_parser("availability", help="Get free/busy availability")
    add_output_args(cmd_cal_availability_cmd)
    cmd_cal_availability_cmd.add_argument("--attendee", required=True, action="append", help="Attendee email (repeatable)")
    cmd_cal_availability_cmd.add_argument("--start", required=True, help="Start time")
    cmd_cal_availability_cmd.add_argument("--end", required=True, help="End time")
    cmd_cal_availability_cmd.add_argument("--interval", type=int, default=30, help="Availability interval minutes")
    cmd_cal_availability_cmd.set_defaults(func=cmd_cal_availability)

    cmd_cal_find_time_cmd = cal_sub.add_parser("find-time", help="Suggest meeting times")
    add_output_args(cmd_cal_find_time_cmd)
    cmd_cal_find_time_cmd.add_argument("--attendee", required=True, action="append", help="Attendee email (repeatable)")
    cmd_cal_find_time_cmd.add_argument("--start", required=True, help="Window start time")
    cmd_cal_find_time_cmd.add_argument("--end", required=True, help="Window end time")
    cmd_cal_find_time_cmd.add_argument("--duration", type=int, default=30, help="Meeting duration minutes")
    cmd_cal_find_time_cmd.add_argument("--count", "-n", type=int, default=10, help="Max suggestions")
    cmd_cal_find_time_cmd.set_defaults(func=cmd_cal_find_time)

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
    cmd_cal_update_cmd.set_defaults(func=cmd_cal_update)

    cmd_cal_delete_cmd = cal_sub.add_parser("delete", help="Delete an event")
    cmd_cal_delete_cmd.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_delete_cmd.set_defaults(func=cmd_cal_delete)

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
        func=cmd_cal_accept,
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
        func=cmd_cal_decline,
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
        func=cmd_cal_tentative,
        send_response=True,
    )

    cmd_cal_cancel_cmd = cal_sub.add_parser("cancel", help="Cancel an event you organize")
    cmd_cal_cancel_cmd.add_argument("event_id", help="Event index (from agenda) or full ID")
    cmd_cal_cancel_cmd.add_argument("--comment", "-m", help="Optional cancellation comment")
    cmd_cal_cancel_cmd.set_defaults(func=cmd_cal_cancel)
