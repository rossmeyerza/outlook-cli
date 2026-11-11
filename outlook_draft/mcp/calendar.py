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
from .validation import require_id

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
        events = client.get_agenda(days=clamp(days), top=clamp(count))
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
        "isOrganizer": event.get("IsOrganizer"),
        "changeKey": event.get("ChangeKey", ""),
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
    outlook: bool = False,
) -> JsonDict:
    """Find meeting slots where no attendee is busy or out of office. Read-only. Creates nothing.

    Computed from Outlook free/busy, between day_start and day_end local time.
    Best slots first: clear slots, then slots with the fewest tentative attendees or
    attendees outside their own working hours, then by start time.
    Set outlook=true for native Outlook suggestions; local slot filters then do not apply.
    """
    start_dt, end_dt = _parse_local_range(start, end)
    interval = interval_for_duration(duration_minutes)
    if outlook:
        with outlook_client() as client:
            suggestions = client.find_meeting_times(attendees=attendees, start_dt=outlook_datetime(start_dt),
                          end_dt=outlook_datetime(end_dt), duration_minutes=duration_minutes, max_candidates=clamp(count))
        return {"mode": "outlook", "suggestions": suggestions}
    options = make_options(day_start, day_end, weekdays_only, tentative_is_free)
    items = _free_busy(attendees, start_dt, end_dt, interval)
    return find_slots_report(items, start_dt, interval, duration_minutes, clamp(count), options)


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


def _change_scope(event: JsonDict) -> str:
    kind = event.get("Type")
    if kind == "SeriesMaster":
        return "series"
    if kind in {"Occurrence", "Exception"}:
        return "occurrence"
    if kind == "SingleInstance" and not event.get("SeriesMasterId") and not event.get("Recurrence"):
        return "single"
    return "unknown"


def _change_preview(event: JsonDict, event_id: str, operation: str, allow_series: bool, comment: str) -> JsonDict:
    scope = _change_scope(event)
    organiser = event.get("IsOrganizer")
    attendees_known = isinstance(event.get("Attendees"), list)
    attendees = event["Attendees"] if attendees_known else []
    reasons = []
    if not attendees_known:
        reasons.append("Attendee data is unavailable. Review the event before any change.")
    if scope == "unknown":
        reasons.append("Cannot determine the event scope safely.")
    if scope == "series" and allow_series is not True:
        reasons.append("This event is a whole series. Set allow_series=true only after explicit approval for the whole series.")
    if operation == "delete" and attendees and organiser is not False:
        reasons.append("Use cal_cancel for an organiser meeting with attendees; deletion can send cancellation notices.")
    if operation == "update" and attendees and type(organiser) is not bool:
        reasons.append("Cannot determine whether this update notifies attendees.")
    if operation == "cancel":
        if organiser is not True:
            reasons.append("Only the organiser can cancel this meeting. Use cal_respond to decline an invitation.")
        if not attendees:
            reasons.append("This event has no attendees. Use cal_delete for a private appointment.")
        if event.get("IsCancelled"):
            reasons.append("This meeting is already cancelled. Check the calendar before any further action.")
    return {
        "id": event_id, "operation": operation, "event": _summary(event), "scope": scope,
        "isOrganizer": organiser, "attendeeCount": len(attendees) if attendees_known else None, "changeKey": event.get("ChangeKey", ""),
        "attendees": [{**address(attendee), "type": attendee.get("Type", "")} for attendee in attendees],
        "willNotifyAttendees": operation == "cancel" and bool(attendees) and organiser is True,
        "comment": comment if operation == "cancel" else "", "canApply": not reasons,
        "reasons": reasons, "confirmed": False, "deleted": False, "cancelled": False,
    }


def _calendar_change(event_id: str, operation: str, confirm: bool, allow_series: bool,
                     comment: str = "", expected_change_key: str | None = None) -> JsonDict:
    if not isinstance(event_id, str) or not event_id.strip() or event_id.strip().isdecimal() or event_id in {".", ".."}:
        raise ValueError("Use a non-empty full event ID, not a cached CLI index.")
    if any(ord(char) < 32 or ord(char) == 127 for char in event_id):
        raise ValueError("The event ID must not contain control characters.")
    if type(confirm) is not bool or type(allow_series) is not bool:
        raise ValueError("confirm and allow_series must be booleans")
    if not isinstance(comment, str):
        raise ValueError("comment must be a string")
    if expected_change_key is not None and (not isinstance(expected_change_key, str) or not expected_change_key.strip()):
        raise ValueError("expected_change_key must be a non-empty string")
    with outlook_client() as client:
        event = client.get_event(event_id)
        result = _change_preview(event, event_id, operation, allow_series, comment)
        if not confirm:
            return result
        if not result["canApply"]:
            raise ValueError(" ".join(result["reasons"]))
        if expected_change_key is not None and event.get("ChangeKey") != expected_change_key:
            raise ValueError("The event changed after the preview. Review a fresh preview before confirmation.")
        if operation == "delete":
            client.delete_event(event_id)
            result["deleted"] = True
        else:
            client.cancel_event(event_id, comment=comment)
            result["cancelled"] = True
        result["confirmed"] = True
    return result


def cal_delete(event_id: str, confirm: bool = False, allow_series: bool = False,
               expected_change_key: str | None = None) -> JsonDict:
    """Preview removal of a calendar entry, not the invitation email. Set confirm=true only after approval of the exact event and scope. Organiser meetings with attendees require cal_cancel instead. A series master requires explicit whole-series approval and allow_series=true; occurrence IDs affect only that occurrence. Optional expected_change_key checks the preview version. Never retry uncertain results automatically."""
    return _calendar_change(event_id, "delete", confirm, allow_series, expected_change_key=expected_change_key)


def cal_cancel(event_id: str, comment: str = "", confirm: bool = False, allow_series: bool = False,
               expected_change_key: str | None = None) -> JsonDict:
    """Preview cancellation of a meeting you organise. Confirmation notifies attendees. Set confirm=true only after approval of the exact event, scope, comment and notifications. A series master also requires explicit whole-series approval and allow_series=true. Use cal_respond to decline someone else's invitation. Optional expected_change_key checks the preview version. Never retry uncertain results automatically."""
    return _calendar_change(event_id, "cancel", confirm, allow_series, comment, expected_change_key)


def cal_rooms(room_list: str | None = None) -> list[JsonDict]:
    """Discover meeting rooms, optionally in a room list. Read-only. Tenant support and permissions can limit this endpoint."""
    with outlook_client() as client:
        return client.find_rooms(room_list=room_list)


def cal_update(event_id: str, subject: str | None = None, start: str | None = None,
               end: str | None = None, location: str | None = None, body: str | None = None,
               online_meeting: bool | None = None, confirm: bool = False, allow_series: bool = False,
               expected_change_key: str | None = None) -> JsonDict:
    """Preview calendar edits. Omitted fields stay unchanged; empty location/body clears them. Time edits require both start and end. Updates can notify attendees. Set confirm=true only after approval of the event, changes and notifications. A series master also requires allow_series=true. Never retry uncertain writes automatically."""
    require_id(event_id, "event_id")
    if (start is None) != (end is None):
        raise ValueError("Time edits require both start and end.")
    fields = {key: value for key, value in dict(subject=subject, location=location, body=body,
              is_online_meeting=online_meeting).items() if value is not None}
    if start is not None:
        fields["start_dt"], fields["end_dt"] = _parse_range(start, end)
    if not fields:
        raise ValueError("Supply at least one event field to change.")
    with outlook_client() as client:
        event = client.get_event(event_id)
        result = _change_preview(event, event_id, "update", allow_series, "")
        result.update(changes=fields, updated=False, willNotifyAttendees=bool(event.get("Attendees")) and event.get("IsOrganizer") is True)
        if confirm is True:
            if not result["canApply"]:
                raise ValueError(" ".join(result["reasons"]))
            if expected_change_key is not None and event.get("ChangeKey") != expected_change_key:
                raise ValueError("The event changed after the preview. Review a fresh preview.")
            result.update(confirmed=True, updated=True, event=_summary(client.update_event(event_id, **fields)))
    return result


TOOLS = [cal_agenda, cal_get, cal_availability, cal_find_time, cal_create, cal_respond, cal_delete, cal_cancel, cal_rooms, cal_update]
