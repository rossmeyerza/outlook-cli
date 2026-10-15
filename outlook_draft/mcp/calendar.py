"""Calendar tools."""

from __future__ import annotations

import logging
from typing import Literal

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


def _parse_range(start: str, end: str) -> tuple[str, str]:
    try:
        return (
            outlook_datetime(parse_local_datetime(start)),
            outlook_datetime(parse_local_datetime(end)),
        )
    except Exception as exc:
        log.debug("Best-effort operation failed", exc_info=True)
        raise ValueError(
            f"Could not parse dates ({exc}). Use formats like '2026-04-10 14:00' or 'tomorrow 2pm'."
        ) from exc


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
    attendees: list[str], start: str, end: str, interval_minutes: int = 30
) -> list[JsonDict]:
    """Get free/busy for people or rooms between two local times. Read-only. Times are in the configured local timezone."""
    start_dt, end_dt = _parse_range(start, end)
    with outlook_client() as client:
        items = client.get_schedule(
            schedules=attendees,
            start_dt=start_dt,
            end_dt=end_dt,
            interval_minutes=interval_minutes,
        )
    return [
        {
            "schedule": item.get("ScheduleId", ""),
            "availabilityView": item.get("AvailabilityView", ""),
            "error": (item.get("Error") or {}).get("Message", ""),
        }
        for item in items
    ]


def cal_find_time(
    attendees: list[str],
    start: str,
    end: str,
    duration_minutes: int = 30,
    count: int = 10,
) -> list[JsonDict]:
    """Ask Outlook to suggest meeting times for attendees within a window. Read-only. Does not create anything."""
    start_dt, end_dt = _parse_range(start, end)
    with outlook_client() as client:
        suggestions = client.find_meeting_times(
            attendees=attendees,
            start_dt=start_dt,
            end_dt=end_dt,
            duration_minutes=duration_minutes,
            max_candidates=clamp(count, high=50),
        )
    return [
        {
            "start": (s.get("MeetingTimeSlot") or {}).get("Start", {}).get("DateTime", ""),
            "end": (s.get("MeetingTimeSlot") or {}).get("End", {}).get("DateTime", ""),
            "confidence": s.get("Confidence", ""),
        }
        for s in suggestions
    ]


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
