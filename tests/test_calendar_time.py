from __future__ import annotations

from datetime import datetime
from zoneinfo import ZoneInfo

import pytest

from outlook_draft import calendar_time, config


def test_parse_naive_datetime_uses_configured_local_timezone(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(config, "LOCAL_TIMEZONE", "Europe/London")

    parsed = calendar_time.parse_local_datetime("2026-05-18 16:30")

    assert parsed == datetime(2026, 5, 18, 16, 30, tzinfo=ZoneInfo("Europe/London"))
    assert calendar_time.outlook_datetime(parsed) == "2026-05-18T16:30:00"


def test_parse_aware_datetime_converts_to_configured_local_timezone(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(config, "LOCAL_TIMEZONE", "Europe/London")

    parsed = calendar_time.parse_local_datetime("2026-05-18T15:30:00Z")

    assert parsed.hour == 16
    assert parsed.minute == 30
    assert parsed.tzinfo == ZoneInfo("Europe/London")


def test_invalid_local_timezone_raises_helpful_error(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(config, "LOCAL_TIMEZONE", "Not/AZone")

    with pytest.raises(ValueError, match="Unknown LOCAL_TIMEZONE"):
        calendar_time.local_timezone()


def test_outlook_timezone_prefer_header(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(config, "OUTLOOK_TIMEZONE", "GMT Standard Time")

    assert calendar_time.outlook_timezone_prefer_header() == 'outlook.timezone="GMT Standard Time"'
