from __future__ import annotations

from datetime import date, datetime, time, timedelta
from zoneinfo import ZoneInfo, ZoneInfoNotFoundError

import dateutil.parser

from . import config


def local_timezone() -> ZoneInfo:
    """Return the configured local timezone."""
    try:
        return ZoneInfo(config.LOCAL_TIMEZONE)
    except ZoneInfoNotFoundError as exc:
        raise ValueError(f"Unknown LOCAL_TIMEZONE: {config.LOCAL_TIMEZONE}") from exc


def parse_local_datetime(value: str) -> datetime:
    """Parse a user-provided datetime into the configured local timezone.

    Naive datetimes are treated as local. Aware datetimes are converted to local.
    """
    parsed = dateutil.parser.parse(value)
    tz = local_timezone()
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=tz)
    return parsed.astimezone(tz)


def parse_range_bound(value: str, *, is_end: bool) -> datetime:
    """Parse an ISO date or datetime for one bound of a calendar view.

    A naive datetime is local. A date-only start is local midnight. A date-only end
    is inclusive, so it resolves to the local midnight after that date.
    """
    text = value.strip()
    try:
        day = date.fromisoformat(text)
    except ValueError:
        day = None
    if day is not None:
        if is_end:
            day += timedelta(days=1)
        return datetime.combine(day, time(), tzinfo=local_timezone())
    parsed = datetime.fromisoformat(text)
    return parsed.replace(tzinfo=local_timezone()) if parsed.tzinfo is None else parsed


def calendar_window(
    days: int | None, start: str | None, end: str | None, *, prefix: str = ""
) -> tuple[datetime, datetime] | None:
    """Resolve an explicit calendar range, or return None for the default look-ahead.

    Raises ValueError when days is combined with a range, when only one bound is given,
    when a bound is not an ISO date or datetime, or when start is not before end.
    """
    if start is None and end is None:
        return None
    if days is not None:
        raise ValueError(f"Use {prefix}days or {prefix}start and {prefix}end, not both.")
    if start is None or end is None:
        raise ValueError(f"Give both {prefix}start and {prefix}end.")
    bounds = []
    for name, value in (("start", start), ("end", end)):
        try:
            bounds.append(parse_range_bound(value, is_end=name == "end"))
        except ValueError as exc:
            raise ValueError(
                f"{prefix}{name} must be an ISO date or datetime, for example 2026-11-16 "
                f"or 2026-11-16T09:00 (got {value!r})."
            ) from exc
    start_dt, end_dt = bounds
    if start_dt >= end_dt:
        raise ValueError(f"{prefix}start ({start}) must be before {prefix}end ({end}).")
    return start_dt, end_dt


def outlook_datetime(value: datetime) -> str:
    """Format a datetime for Outlook REST API dateTime fields."""
    return value.strftime("%Y-%m-%dT%H:%M:%S")


def outlook_timezone() -> str:
    """Return the configured Microsoft Outlook timezone name."""
    return config.OUTLOOK_TIMEZONE


def outlook_timezone_prefer_header() -> str:
    """Return the Outlook REST Prefer header value for timezone-aware responses."""
    return f'outlook.timezone="{outlook_timezone()}"'
