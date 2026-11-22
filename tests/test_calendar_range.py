"""Calendar view for an explicit date range (cal agenda --start/--end, cal_agenda start/end)."""
from __future__ import annotations

from datetime import UTC, datetime, timedelta
from unittest.mock import Mock

import httpx
import pytest
from mcp_fakes import FakeClient, use_client
from test_cli_smoke import isolated_cli, run_command  # noqa: F401  (pytest fixture import)
from test_outlook_client_calendar import RecordingClient

from outlook_draft import config, runtime
from outlook_draft.calendar_time import calendar_window, parse_range_bound
from outlook_draft.mcp import calendar as mcp_calendar

ZONE = "Africa/Johannesburg"  # UTC+02:00 all year, so offsets are visible


@pytest.fixture(autouse=True)
def local_zone(monkeypatch):
    monkeypatch.setattr(config, "LOCAL_TIMEZONE", ZONE)


def _event(**overrides):
    base = {
        "Id": "ev-1",
        "Subject": "Client workshop",
        "Start": {"DateTime": "2026-11-20T14:00:00.0000000", "TimeZone": "GMT Standard Time"},
        "End": {"DateTime": "2026-11-20T15:30:00.0000000", "TimeZone": "GMT Standard Time"},
        "IsAllDay": False,
        "IsCancelled": False,
        "ShowAs": "Busy",
        "ResponseStatus": {"Response": "Accepted", "Time": "2026-11-18T09:00:00Z"},
    }
    base.update(overrides)
    return base


# ── Range parsing ──────────────────────────────────────────────────────

def test_date_only_start_is_local_midnight_and_end_is_inclusive():
    start = parse_range_bound("2026-11-16", is_end=False)
    end = parse_range_bound("2026-11-20", is_end=True)
    assert start.isoformat() == "2026-11-16T00:00:00+02:00"
    assert end.isoformat() == "2026-11-21T00:00:00+02:00"


def test_basic_format_date_is_still_an_inclusive_date():
    assert parse_range_bound("20261120", is_end=True).isoformat() == "2026-11-21T00:00:00+02:00"


def test_naive_datetime_is_local_and_aware_datetime_keeps_its_offset():
    assert parse_range_bound("2026-11-16T09:30", is_end=False).isoformat() == "2026-11-16T09:30:00+02:00"
    assert parse_range_bound("2026-11-16 09:30", is_end=True).isoformat() == "2026-11-16T09:30:00+02:00"
    assert parse_range_bound("2026-11-16T09:30:00Z", is_end=False).utcoffset() == timedelta(0)


def test_window_is_none_without_a_range():
    assert calendar_window(None, None, None) is None
    assert calendar_window(14, None, None) is None


def test_single_day_range_is_valid():
    start, end = calendar_window(None, "2026-11-16", "2026-11-16")
    assert end - start == timedelta(days=1)


@pytest.mark.parametrize("days, start, end, message", [
    (7, "2026-11-16", "2026-11-20", "not both"),
    (7, "2026-11-16", None, "not both"),
    (None, "2026-11-16", None, "both start and end"),
    (None, None, "2026-11-20", "both start and end"),
    (None, "2026-11-20", "2026-11-16", "must be before"),
    (None, "2026-11-16T10:00", "2026-11-16T10:00", "must be before"),
    (None, "next monday", "2026-11-20", "ISO date or datetime"),
    (None, "2026-11-16", "16/11/2026", "ISO date or datetime"),
])
def test_invalid_ranges_are_rejected(days, start, end, message):
    with pytest.raises(ValueError, match=message):
        calendar_window(days, start, end)


# ── Client ─────────────────────────────────────────────────────────────

def test_calendar_view_sends_utc_instants_and_timesheet_fields(monkeypatch):
    monkeypatch.setattr(config, "OUTLOOK_TIMEZONE", "GMT Standard Time")
    client = RecordingClient()
    start, end = calendar_window(None, "2026-11-16", "2026-11-20")

    client.get_calendar_view(start, end)

    call = client.calls[0]
    assert call["path"] == "/me/calendarview"
    assert call["params"]["startDateTime"] == "2026-11-15T22:00:00+00:00"
    assert call["params"]["endDateTime"] == "2026-11-20T22:00:00+00:00"
    assert call["params"]["$top"] == "100"
    for field in ("subject", "start", "end", "isAllDay", "isCancelled", "showAs", "responseStatus"):
        assert field in call["params"]["$select"].split(",")
    assert call["extra_headers"] == {"Prefer": 'outlook.timezone="GMT Standard Time"'}


class PagingClient(RecordingClient):
    def __init__(self, pages):
        super().__init__()
        self.pages = list(pages)

    def _request(self, method, path, **kwargs):
        super()._request(method, path, **kwargs)
        return httpx.Response(200, json=self.pages.pop(0))


def _pages():
    return [
        {"value": [{"Id": "1"}, {"Id": "2"}], "@odata.nextLink": "https://outlook.office.com/api/v2.0/next?page=2"},
        {"value": [{"Id": "3"}]},
    ]


def test_calendar_view_follows_next_links_until_the_range_ends():
    client = PagingClient(_pages())
    start = datetime(2026, 11, 16, tzinfo=UTC)

    events = client.get_calendar_view(start, start + timedelta(days=5))

    assert [e["Id"] for e in events] == ["1", "2", "3"]
    assert client.calls[1]["path"] == "https://outlook.office.com/api/v2.0/next?page=2"
    assert client.calls[1]["params"] is None


def _day(day, *, all_day=True, start="00:00", end_day=None, end="00:00"):
    return {
        "Id": f"{day}-{start}",
        "IsAllDay": all_day,
        "Start": {"DateTime": f"2026-11-{day}T{start}:00.0000000", "TimeZone": "South Africa Standard Time"},
        "End": {"DateTime": f"2026-11-{end_day or day + 1}T{end}:00.0000000", "TimeZone": "South Africa Standard Time"},
    }


def test_calendar_view_drops_all_day_events_outside_the_dates():
    # Outlook can return the neighbouring day's all-day event when it was created in another zone.
    page = {"value": [
        _day(15), _day(16), _day(20), _day(21), _day(19, end_day=29),
        _day(20, all_day=False, start="23:00", end_day=21, end="01:00"),
        {"Id": "no-dates", "IsAllDay": True},
    ]}
    client = PagingClient([page])
    start, end = calendar_window(None, "2026-11-16", "2026-11-20")

    events = client.get_calendar_view(start, end)

    assert [e["Id"] for e in events] == ["16-00:00", "20-00:00", "19-00:00", "20-23:00", "no-dates"]


def test_calendar_view_count_ignores_dropped_events():
    page = {"value": [_day(15), _day(16), _day(17)], "@odata.nextLink": "https://outlook.office.com/next"}
    client = PagingClient([page, {"value": [_day(18)]}])
    start, end = calendar_window(None, "2026-11-16", "2026-11-20")

    events = client.get_calendar_view(start, end, top=3)

    assert [e["Id"] for e in events] == ["16-00:00", "17-00:00", "18-00:00"]


def test_calendar_view_stops_at_the_requested_count():
    client = PagingClient(_pages())
    start = datetime(2026, 11, 16, tzinfo=UTC)

    events = client.get_calendar_view(start, start + timedelta(days=5), top=2)

    assert [e["Id"] for e in events] == ["1", "2"]
    assert len(client.calls) == 1
    assert client.calls[0]["params"]["$top"] == "2"


def test_get_event_selects_show_as_and_response_status():
    client = RecordingClient()
    client.get_event("event-1")
    select = client.calls[0]["params"]["$select"].split(",")
    assert "showAs" in select and "responseStatus" in select


# ── CLI ────────────────────────────────────────────────────────────────

def test_cli_range_reads_the_calendar_view_without_a_count(monkeypatch, capsys, isolated_cli):  # noqa: F811
    import json

    isolated_cli.get_calendar_view.return_value = [_event()]
    result = run_command(monkeypatch, capsys, ["cal", "agenda", "--start", "2026-11-16", "--end", "2026-11-20", "--json"])

    assert result["exit"] == 0
    kwargs = isolated_cli.get_calendar_view.call_args.kwargs
    assert kwargs["start"].isoformat() == "2026-11-16T00:00:00+02:00"
    assert kwargs["end"].isoformat() == "2026-11-21T00:00:00+02:00"
    assert kwargs["top"] is None
    isolated_cli.get_agenda.assert_not_called()
    row = json.loads(result["stdout"])[0]
    assert row["showAs"] == "Busy"
    assert row["responseStatus"] == "Accepted"
    assert row["isAllDay"] is False and row["isCancelled"] is False


def test_cli_range_table_shows_the_range_and_count_limit(monkeypatch, capsys, isolated_cli):  # noqa: F811
    isolated_cli.get_calendar_view.return_value = [_event(), _event(Id="ev-2")]
    result = run_command(monkeypatch, capsys,
                         ["cal", "agenda", "--start", "2026-11-16", "--end", "2026-11-20", "-n", "2"])

    assert result["exit"] == 0
    assert isolated_cli.get_calendar_view.call_args.kwargs["top"] == 2
    assert "Agenda (2026-11-16 to 2026-11-20)" in result["stdout"]
    assert "Showing the first 2 events" in result["stdout"]


def test_cli_default_still_looks_ahead_seven_days(monkeypatch, capsys, isolated_cli):  # noqa: F811
    result = run_command(monkeypatch, capsys, ["cal", "agenda"])

    assert result["exit"] == 0
    assert isolated_cli.get_agenda.call_args.kwargs == {"days": 7, "top": 20}
    assert "Agenda (next 7 days)" in result["stdout"]
    isolated_cli.get_calendar_view.assert_not_called()


@pytest.mark.parametrize("argv, message", [
    (["--days", "3", "--start", "2026-11-16", "--end", "2026-11-20"], "Use --days or --start and --end, not both."),
    (["-d", "3", "--end", "2026-11-20"], "not both"),
    (["--start", "2026-11-16"], "Give both --start and --end."),
    (["--start", "2026-11-20", "--end", "2026-11-16"], "--start (2026-11-20) must be before --end (2026-11-16)."),
    (["--start", "yesterday", "--end", "2026-11-16"], "--start must be an ISO date or datetime"),
])
def test_cli_rejects_bad_ranges_before_it_builds_a_client(monkeypatch, capsys, isolated_cli, argv, message):  # noqa: F811
    fail = Mock(side_effect=AssertionError("Invalid input must not authenticate"))
    monkeypatch.setattr(runtime, "get_client", fail)

    result = run_command(monkeypatch, capsys, ["cal", "agenda", *argv])

    assert result["exit"] == 2
    assert message in result["stdout"]
    fail.assert_not_called()


# ── MCP ────────────────────────────────────────────────────────────────

def test_mcp_range_returns_every_event_with_timesheet_fields(monkeypatch):
    fake = use_client(monkeypatch, mcp_calendar, "outlook_client", FakeClient(get_calendar_view=[_event()]))

    result = mcp_calendar.cal_agenda(start="2026-11-16", end="2026-11-20")

    kwargs = fake.call("get_calendar_view")[1]
    assert kwargs["start"].isoformat() == "2026-11-16T00:00:00+02:00"
    assert kwargs["end"].isoformat() == "2026-11-21T00:00:00+02:00"
    assert kwargs["top"] is None
    assert "get_agenda" not in fake.names()
    assert result[0]["showAs"] == "Busy"
    assert result[0]["responseStatus"] == "Accepted"
    assert "index" not in result[0]


def test_mcp_range_passes_an_explicit_count(monkeypatch):
    fake = use_client(monkeypatch, mcp_calendar, "outlook_client", FakeClient(get_calendar_view=[]))
    mcp_calendar.cal_agenda(start="2026-11-16T08:00", end="2026-11-16T18:00", count=5)
    assert fake.call("get_calendar_view")[1]["top"] == 5


def test_mcp_default_is_unchanged(monkeypatch):
    fake = use_client(monkeypatch, mcp_calendar, "outlook_client", FakeClient(get_agenda=[]))
    mcp_calendar.cal_agenda()
    assert fake.call("get_agenda")[1] == {"days": 7, "top": 50}


@pytest.mark.parametrize("kwargs, message", [
    ({"days": 3, "start": "2026-11-16", "end": "2026-11-20"}, "Use days or start and end, not both."),
    ({"end": "2026-11-20"}, "Give both start and end."),
    ({"start": "2026-11-20", "end": "2026-11-16"}, "must be before"),
])
def test_mcp_rejects_bad_ranges_without_a_client(monkeypatch, kwargs, message):
    fake = use_client(monkeypatch, mcp_calendar, "outlook_client", FakeClient())
    with pytest.raises(ValueError, match=message):
        mcp_calendar.cal_agenda(**kwargs)
    assert not fake.names()
