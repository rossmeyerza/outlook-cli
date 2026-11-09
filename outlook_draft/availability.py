"""Decode Outlook free/busy data into readable ranges and common free slots.

Outlook's getSchedule returns an AvailabilityView for each person. It is a
string with one digit for each interval from the requested start time:
0 free, 1 tentative, 2 busy, 3 out of office, 4 working elsewhere.
This module turns the digits into time ranges and finds the slots where
no attendee is busy. It does no HTTP, so the CLI and the MCP server share it.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from datetime import UTC, datetime, time, timedelta
from typing import Any
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

JsonDict = dict[str, Any]

STATUS_NAMES = {
    "0": "free",
    "1": "tentative",
    "2": "busy",
    "3": "outOfOffice",
    "4": "workingElsewhere",
}
# A person who works elsewhere is still available for an online meeting.
AVAILABLE = frozenset({"0", "4"})
TENTATIVE = "1"
MIN_INTERVAL, MAX_INTERVAL = 5, 1440  # getSchedule limits
DAY_NAMES = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")

# Outlook gives WorkingHours time zones as Windows names. Python needs IANA names.
WINDOWS_TIMEZONES = {
    "UTC": "UTC",
    "GMT Standard Time": "Europe/London",
    "Greenwich Standard Time": "Atlantic/Reykjavik",
    "W. Europe Standard Time": "Europe/Berlin",
    "Romance Standard Time": "Europe/Paris",
    "Central Europe Standard Time": "Europe/Budapest",
    "Central European Standard Time": "Europe/Warsaw",
    "E. Europe Standard Time": "Europe/Chisinau",
    "FLE Standard Time": "Europe/Kyiv",
    "GTB Standard Time": "Europe/Bucharest",
    "Turkey Standard Time": "Europe/Istanbul",
    "Russian Standard Time": "Europe/Moscow",
    "South Africa Standard Time": "Africa/Johannesburg",
    "Arabian Standard Time": "Asia/Dubai",
    "India Standard Time": "Asia/Kolkata",
    "Singapore Standard Time": "Asia/Singapore",
    "China Standard Time": "Asia/Shanghai",
    "Tokyo Standard Time": "Asia/Tokyo",
    "AUS Eastern Standard Time": "Australia/Sydney",
    "New Zealand Standard Time": "Pacific/Auckland",
    "Eastern Standard Time": "America/New_York",
    "Central Standard Time": "America/Chicago",
    "Mountain Standard Time": "America/Denver",
    "Pacific Standard Time": "America/Los_Angeles",
    "E. South America Standard Time": "America/Sao_Paulo",
}


@dataclass(frozen=True)
class SlotOptions:
    """Limits for candidate slots, in the local time zone."""

    day_start: time = time(8, 0)
    day_end: time = time(18, 0)
    weekdays_only: bool = True
    tentative_is_free: bool = True


@dataclass(frozen=True)
class WorkingHours:
    days: frozenset[int]
    start_minute: int
    end_minute: int
    tz: ZoneInfo
    label: str


@dataclass(frozen=True)
class Schedule:
    schedule_id: str
    view: str
    working_hours: WorkingHours | None


@dataclass(frozen=True)
class Slot:
    start: datetime
    end: datetime
    tentative: tuple[str, ...]
    outside_hours: tuple[str, ...]

    @property
    def issues(self) -> int:
        return len(self.tentative) + len(self.outside_hours)

    def to_dict(self) -> JsonDict:
        return {
            "start": format_local(self.start),
            "end": format_local(self.end),
            "label": format_span(self.start, self.end),
            "minutes": int((self.end - self.start).total_seconds() // 60),
            "clear": self.issues == 0,
            "tentative": list(self.tentative),
            "outsideWorkingHours": list(self.outside_hours),
        }


def parse_clock(value: str) -> time:
    """Parse 'HH:MM' (seconds and fractions are ignored)."""
    try:
        hours, minutes = str(value).strip().split(":")[:2]
        return time(int(hours), int(minutes))
    except ValueError as exc:
        raise ValueError(f"Use HH:MM for a time of day, got {value!r}") from exc


def make_options(
    day_start: str = "08:00",
    day_end: str = "18:00",
    weekdays_only: bool = True,
    tentative_is_free: bool = True,
) -> SlotOptions:
    start, end = parse_clock(day_start), parse_clock(day_end)
    if end <= start:
        raise ValueError(f"day_end ({day_end}) must be later than day_start ({day_start})")
    return SlotOptions(start, end, weekdays_only, tentative_is_free)


def check_interval(minutes: int) -> int:
    if not MIN_INTERVAL <= minutes <= MAX_INTERVAL:
        raise ValueError(f"interval must be {MIN_INTERVAL} to {MAX_INTERVAL} minutes, got {minutes}")
    return minutes


def interval_for_duration(duration: int) -> int:
    """Pick a free/busy interval that divides the meeting length."""
    if duration <= 0 or duration > MAX_INTERVAL:
        raise ValueError(f"duration must be 1 to {MAX_INTERVAL} minutes, got {duration}")
    return max(MIN_INTERVAL, math.gcd(duration, 30))


def format_local(value: datetime) -> str:
    return value.strftime("%Y-%m-%dT%H:%M")


def format_span(start: datetime, end: datetime) -> str:
    if start.date() == end.date() or end == datetime.combine(start.date() + timedelta(days=1), time(0), end.tzinfo):
        return f"{start:%a %d %b} {start:%H:%M}-{end:%H:%M}"
    return f"{start:%a %d %b %H:%M} to {end:%a %d %b %H:%M}"


def slot_time(start: datetime, interval: int, index: int) -> datetime:
    """Start of interval `index`. UTC arithmetic keeps steps correct across DST changes."""
    moment = start.astimezone(UTC) + timedelta(minutes=interval * index)
    return moment.astimezone(start.tzinfo)


def _zone(name: str) -> ZoneInfo | None:
    try:
        return ZoneInfo(WINDOWS_TIMEZONES.get(name, name))
    except (ZoneInfoNotFoundError, ValueError):
        return None


def _day_label(days: frozenset[int]) -> str:
    ordered = sorted(days)
    short = [DAY_NAMES[d][:3] for d in ordered]
    if len(ordered) > 2 and ordered == list(range(ordered[0], ordered[-1] + 1)):
        return f"{short[0]}-{short[-1]}"
    return ",".join(short)


def parse_working_hours(raw: JsonDict | None) -> WorkingHours | None:
    """Parse Outlook WorkingHours. Return None when the data or time zone is unknown."""
    if not raw:
        return None
    tz_name = (raw.get("TimeZone") or {}).get("Name", "")
    tz = _zone(tz_name)
    if tz is None:
        return None
    days = frozenset(DAY_NAMES.index(d) for d in raw.get("DaysOfWeek") or [] if d in DAY_NAMES)
    start = parse_clock(raw.get("StartTime") or "00:00")
    end = parse_clock(raw.get("EndTime") or "00:00")
    start_minute = start.hour * 60 + start.minute
    end_minute = end.hour * 60 + end.minute
    if end_minute <= start_minute:
        end_minute = 24 * 60  # an end of 00:00 means midnight
    label = f"{_day_label(days)} {start:%H:%M}-{end:%H:%M} {tz_name}"
    return WorkingHours(days, start_minute, end_minute, tz, label)


def _minutes_in_day(start: datetime, end: datetime) -> tuple[int, int] | None:
    """Return start and end as minutes after midnight, or None if the span crosses days."""
    first = start.hour * 60 + start.minute
    last = first + int((end - start).total_seconds() // 60)
    return (first, last) if last <= 24 * 60 else None


def in_working_hours(hours: WorkingHours, start: datetime, end: datetime) -> bool:
    local_start = start.astimezone(hours.tz)
    if local_start.weekday() not in hours.days:
        return False
    span = _minutes_in_day(local_start, end.astimezone(hours.tz))
    return span is not None and hours.start_minute <= span[0] and span[1] <= hours.end_minute


def _in_day_window(start: datetime, end: datetime, options: SlotOptions) -> bool:
    if options.weekdays_only and start.weekday() >= 5:
        return False
    span = _minutes_in_day(start, end)
    if span is None:
        return False
    window_start = options.day_start.hour * 60 + options.day_start.minute
    window_end = options.day_end.hour * 60 + options.day_end.minute
    return window_start <= span[0] and span[1] <= window_end


def decode_ranges(view: str, start: datetime, interval: int) -> list[JsonDict]:
    """Merge runs of the same non-free status into time ranges."""
    ranges: list[JsonDict] = []
    index = 0
    while index < len(view):
        code = view[index]
        run_end = index
        while run_end < len(view) and view[run_end] == code:
            run_end += 1
        if code != "0":
            range_start = slot_time(start, interval, index)
            range_end = slot_time(start, interval, run_end)
            ranges.append({
                "status": STATUS_NAMES.get(code, f"unknown({code})"),
                "start": format_local(range_start),
                "end": format_local(range_end),
                "label": format_span(range_start, range_end),
            })
        index = run_end
    return ranges


def schedules_from_response(items: list[JsonDict]) -> tuple[list[Schedule], list[JsonDict]]:
    """Split getSchedule items into usable schedules and attendees with no data."""
    schedules: list[Schedule] = []
    unknown: list[JsonDict] = []
    for item in items:
        schedule_id = item.get("ScheduleId", "")
        error = (item.get("Error") or {}).get("Message", "")
        view = item.get("AvailabilityView") or ""
        if error or not view:
            unknown.append({"schedule": schedule_id, "error": error or "Outlook returned no free/busy data"})
            continue
        schedules.append(Schedule(schedule_id, view, parse_working_hours(item.get("WorkingHours"))))
    return schedules, unknown


def _evaluate(schedules: list[Schedule], first: int, steps: int, slot: Slot, options: SlotOptions) -> Slot | None:
    """Return the slot with its flags, or None when an attendee cannot attend."""
    tentative: list[str] = []
    outside: list[str] = []
    for schedule in schedules:
        codes = schedule.view[first:first + steps]
        # A short view or an unknown digit counts as busy, to stay safe.
        if len(codes) < steps or any(c not in AVAILABLE and c != TENTATIVE for c in codes):
            return None
        if TENTATIVE in codes:
            if not options.tentative_is_free:
                return None
            tentative.append(schedule.schedule_id)
        hours = schedule.working_hours
        if hours and not in_working_hours(hours, slot.start, slot.end):
            outside.append(schedule.schedule_id)
    return Slot(slot.start, slot.end, tuple(tentative), tuple(outside))


def candidate_slots(
    schedules: list[Schedule],
    start: datetime,
    interval: int,
    duration: int,
    options: SlotOptions,
) -> list[Slot]:
    """Every slot of `duration` minutes, on interval steps, where no attendee is busy."""
    if not schedules:
        return []
    steps = max(1, math.ceil(duration / interval))
    length = min(len(s.view) for s in schedules)
    slots: list[Slot] = []
    for first in range(0, length - steps + 1):
        slot_start = slot_time(start, interval, first)
        slot_end = slot_start.astimezone(UTC) + timedelta(minutes=duration)
        empty = Slot(slot_start, slot_end.astimezone(start.tzinfo), (), ())
        if not _in_day_window(empty.start, empty.end, options):
            continue
        slot = _evaluate(schedules, first, steps, empty, options)
        if slot:
            slots.append(slot)
    return slots


def merge_windows(slots: list[Slot]) -> list[Slot]:
    """Join back-to-back slots that have the same tentative and working-hours flags."""
    windows: list[Slot] = []
    for slot in slots:
        last = windows[-1] if windows else None
        if (
            last
            and last.end == slot.start
            and last.tentative == slot.tentative
            and last.outside_hours == slot.outside_hours
        ):
            windows[-1] = Slot(last.start, slot.end, last.tentative, last.outside_hours)
        else:
            windows.append(slot)
    return windows


def _settings(start: datetime, options: SlotOptions, unknown: list[JsonDict]) -> JsonDict:
    result: JsonDict = {
        "timezone": str(start.tzinfo),
        "dayStart": f"{options.day_start:%H:%M}",
        "dayEnd": f"{options.day_end:%H:%M}",
        "weekdaysOnly": options.weekdays_only,
        "tentativeIsFree": options.tentative_is_free,
        "unknown": unknown,
    }
    if unknown:
        names = ", ".join(u["schedule"] for u in unknown)
        result["warning"] = f"Free slots ignore these attendees, because Outlook returned no data for them: {names}"
    return result


def availability_report(
    items: list[JsonDict],
    start: datetime,
    interval: int,
    options: SlotOptions,
    include_raw: bool = False,
) -> JsonDict:
    """Decoded ranges for each attendee, plus the windows where nobody is busy."""
    schedules, unknown = schedules_from_response(items)
    attendees = []
    for schedule in schedules:
        entry: JsonDict = {
            "schedule": schedule.schedule_id,
            "workingHours": schedule.working_hours.label if schedule.working_hours else "unknown",
            "ranges": decode_ranges(schedule.view, start, interval),
        }
        if include_raw:
            entry["availabilityView"] = schedule.view
        attendees.append(entry)
    windows = merge_windows(candidate_slots(schedules, start, interval, interval, options))
    return {
        **_settings(start, options, unknown),
        "intervalMinutes": interval,
        "attendees": attendees,
        "commonFree": [window.to_dict() for window in windows],
    }


def find_slots_report(
    items: list[JsonDict],
    start: datetime,
    interval: int,
    duration: int,
    count: int,
    options: SlotOptions,
) -> JsonDict:
    """Meeting slots where nobody is busy. Clear slots come first, then by start time."""
    schedules, unknown = schedules_from_response(items)
    slots = candidate_slots(schedules, start, interval, duration, options)
    ranked = sorted(slots, key=lambda slot: (slot.issues, slot.start))[:count]
    return {
        **_settings(start, options, unknown),
        "durationMinutes": duration,
        "slots": [slot.to_dict() for slot in ranked],
    }
