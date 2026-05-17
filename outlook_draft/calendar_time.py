from __future__ import annotations

from datetime import datetime
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


def outlook_datetime(value: datetime) -> str:
    """Format a datetime for Outlook REST API dateTime fields."""
    return value.strftime("%Y-%m-%dT%H:%M:%S")


def outlook_timezone() -> str:
    """Return the configured Microsoft Outlook timezone name."""
    return config.OUTLOOK_TIMEZONE


def outlook_timezone_prefer_header() -> str:
    """Return the Outlook REST Prefer header value for timezone-aware responses."""
    return f'outlook.timezone="{outlook_timezone()}"'
