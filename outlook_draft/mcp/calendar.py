"""Calendar tools."""

from __future__ import annotations

import logging
from datetime import datetime
from typing import Literal

from ..availability import (
    availability_report,
    check_interval,
    find_slots_report,
    interval_for_duration,
    make_options,
)
from ..calendar_time import outlook_datetime, parse_local_datetime
from ..commands.calendar import _event_recurrence, _event_summary, _event_when
from .clients import clamp, outlook_client
from .shaping import JsonDict, address, body_text, truncate

RESPONSE_ACTIONS = {"accept": "accept", "tentative": "tentativelyAccept", "decline": "decline"}


log = logging.getLogger(__name__)


def _summary(event: JsonDict) -> JsonDict:
    summary = _event_summary(event)
    summary.pop("index", None)
    summary["when"] = _event_when(event)
    return summary


def _parse_local_range(start: str, end: str) -> tuple[datetime, datetime]:
    try:
        start_dt, end_dt = parse_local_datetime(start), parse_local_datetime(end)
    except Exception as exc:
        log.debug("Best-effort operation failed", exc_info=True)
        raise ValueError(
            f"Could not parse dates ({exc}). Use formats like '2026-04-10 14:00' or 'tomorrow 2pm'."
        ) from exc
    if end_dt <= start_dt:
        raise ValueError(f"end ({end}) must be later than start ({start})")
    return start_dt, end_dt


def _parse_range(start: str, end: str) -> tuple[str, str]:
    start_dt, end_dt = _parse_local_range(start, end)
    return outlook_datetime(start_dt), outlook_datetime(end_dt)


def _free_busy(attendees: list[str], start_dt: datetime, end_dt: datetime, interval: int) -> list[JsonDict]:
    with outlook_client() as client:
        return client.get_schedule(
            schedules=attendees,
            start_dt=outlook_datetime(start_dt),
            end_dt=outlook_datetime(end_dt),
            interval_minutes=interval,
        )


def cal_agenda(days: int = 7, count: int = 50) -> list[JsonDict]:
    """List upcoming calendar events for the next N days. Read-only. Returns event summaries with full IDs."""
    with outlook_client() as client:
        events = client.get_agenda(days=clamp(days, high=90), top=clamp(count, high=200))
    return [_summary(e) for e in events]


def cal_get(event_id: str) -> JsonDict:
    """Get one calendar event by full ID, including attendees, body text and recurrence. Read-only."""
    with outlook_client() as client:
        event = client.get_event(event_id)
    text, truncated = truncate(body_text(event.get("Body")), 10000)
    return {
        **_summary(event),
        "recurrence": _event_recurrence(event),
        "attendees": [
            {
                **address(a),
                "type": a.get("Type", ""),
                "response": (a.get("Status") or {}).get("Response", ""),
            }
            for a in event.get("Attendees") or []
        ],
        "isOnlineMeeting": event.get("IsOnlineMeeting", False),
        "joinUrl": (event.get("OnlineMeeting") or {}).get("JoinUrl", ""),
        "body": text,
        "bodyTruncated": truncated,
    }


def cal_availability(
    attendees: list[str],
    start: str,
    end: str,
    interval_minutes: int = 30,
    day_start: str = "08:00",
    day_end: str = "18:00",
    weekdays_only: bool = True,
    tentative_is_free: bool = True,
    include_raw: bool = False,
) -> JsonDict:
    """Get free/busy for people or rooms between two local times, decoded. Read-only.

    Returns each attendee's non-free ranges (busy, tentative, outOfOffice, workingElsewhere)
    and their Outlook working hours, plus commonFree windows between day_start and day_end
    where no attendee is busy or out of office. Each window lists attendees who are
    tentative or outside their own working hours; clear=true means neither.
    Times are in the configured local timezone. include_raw adds Outlook's digit string.
    """
    start_dt, end_dt = _parse_local_range(start, end)
    interval = check_interval(interval_minutes)
    options = make_options(day_start, day_end, weekdays_only, tentative_is_free)
    items = _free_busy(attendees, start_dt, end_dt, interval)
    return availability_report(items, start_dt, interval, options, include_raw=include_raw)


def cal_find_time(
    attendees: list[str],
    start: str,
    end: str,
    duration_minutes: int = 30,
    count: int = 10,
    day_start: str = "08:00",
    day_end: str = "18:00",
    weekdays_only: bool = True,
    tentative_is_free: bool = True,
) -> JsonDict:
    """Find meeting slots where no attendee is busy or out of office. Read-only. Creates nothing.

    Computed from Outlook free/busy, between day_start and day_end local time.
    Best slots first: clear slots, then slots with the fewest tentative attendees or
    attendees outside their own working hours, then by start time.
    """
    start_dt, end_dt = _parse_local_range(start, end)
    interval = interval_for_duration(duration_minutes)
    options = make_options(day_start, day_end, weekdays_only, tentative_is_free)
    items = _free_busy(attendees, start_dt, end_dt, interval)
    return find_slots_report(items, start_dt, interval, duration_minutes, clamp(count, high=50), options)


def cal_create(
    subject: str,
    start: str,
    end: str,
    attendees: list[str] | None = None,
    location: str | None = None,
    body: str | None = None,
    online_meeting: bool = False,
) -> JsonDict:
    """Create a calendar event. WARNING: if attendees are given, Outlook emails invitations to them immediately. Omit attendees for a private event."""
    start_dt, end_dt = _parse_range(start, end)
    with outlook_client() as client:
        event = client.create_event(
            subject=subject,
            start_dt=start_dt,
            end_dt=end_dt,
            location=location,
            body=body,
            attendees=attendees or None,
            is_online_meeting=online_meeting,
        )
    return {**_summary(event), "invitationsSentTo": list(attendees or [])}


def cal_respond(
    event_id: str,
    response: Literal["accept", "tentative", "decline"],
    comment: str = "",
    send_response: bool = True,
) -> JsonDict:
    """Accept, tentatively accept or decline an invitation. WARNING: when send_response is true, the organizer is emailed your response."""
    if response not in RESPONSE_ACTIONS:
        raise ValueError(f"response must be one of {sorted(RESPONSE_ACTIONS)}, got {response!r}")
    with outlook_client() as client:
        event = client.get_event(event_id)
        client.respond_to_event(
            event_id,
            RESPONSE_ACTIONS[response],
            comment=comment,
            send_response=send_response,
        )
    return {
        "id": event_id,
        "subject": event.get("Subject", ""),
        "response": response,
        "responseSent": send_response,
    }


TOOLS = [cal_agenda, cal_get, cal_availability, cal_find_time, cal_create, cal_respond]
