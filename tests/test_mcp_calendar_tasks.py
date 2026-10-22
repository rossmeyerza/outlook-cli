from __future__ import annotations

import json
import time
from pathlib import Path

import jwt
import pytest
from mcp_fakes import FakeClient, use_client

from outlook_draft import config
from outlook_draft.mcp import auth, calendar, contacts, tasks


def _event(**overrides):
    base = {
        "Id": "ev-1",
        "Subject": "Standup",
        "Start": {"DateTime": "2026-10-15T09:00:00.0000000", "TimeZone": "GMT Standard Time"},
        "End": {"DateTime": "2026-10-15T09:30:00.0000000", "TimeZone": "GMT Standard Time"},
        "Location": {"DisplayName": "Room 1"},
        "Organizer": {"EmailAddress": {"Name": "Ann", "Address": "ann@example.com"}},
        "Attendees": [
            {"EmailAddress": {"Name": "Bob", "Address": "bob@example.com"}, "Type": "Required", "Status": {"Response": "Accepted"}}
        ],
        "Body": {"ContentType": "HTML", "Content": "<p>Agenda</p>"},
        "OnlineMeeting": {"JoinUrl": "https://teams/join"},
        "IsOnlineMeeting": True,
    }
    base.update(overrides)
    return base


def test_cal_agenda_summaries_have_ids_and_no_index(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_agenda=[_event()]))

    result = calendar.cal_agenda(days=1)

    assert fake.call("get_agenda")[1] == {"days": 1, "top": 50}
    assert result[0]["id"] == "ev-1"
    assert result[0]["location"] == "Room 1"
    assert "index" not in result[0]
    assert "09:00" in result[0]["when"]


def test_cal_get_includes_attendees_and_body_text(monkeypatch):
    use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=_event()))

    result = calendar.cal_get("ev-1")

    assert result["body"] == "Agenda"
    assert result["attendees"] == [
        {"name": "Bob", "address": "bob@example.com", "type": "Required", "response": "Accepted"}
    ]
    assert result["joinUrl"] == "https://teams/join"


def test_cal_availability_parses_local_times(monkeypatch):
    fake = FakeClient(get_schedule=[{"ScheduleId": "bob@example.com", "AvailabilityView": "0022"}])
    use_client(monkeypatch, calendar, "outlook_client", fake)

    result = calendar.cal_availability(["bob@example.com"], "2026-10-16 09:00", "2026-10-16 11:00")

    kwargs = fake.call("get_schedule")[1]
    assert kwargs["start_dt"] == "2026-10-16T09:00:00"
    assert kwargs["end_dt"] == "2026-10-16T11:00:00"
    assert result[0]["availabilityView"] == "0022"


def test_cal_find_time_shapes_suggestions(monkeypatch):
    suggestion = {
        "MeetingTimeSlot": {"Start": {"DateTime": "a"}, "End": {"DateTime": "b"}},
        "Confidence": "High",
    }
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(find_meeting_times=[suggestion]))

    result = calendar.cal_find_time(["bob@example.com"], "2026-10-16 09:00", "2026-10-16 17:00", duration_minutes=45)

    assert fake.call("find_meeting_times")[1]["duration_minutes"] == 45
    assert result == [{"start": "a", "end": "b", "confidence": "High"}]


def test_bad_dates_raise_value_error(monkeypatch):
    use_client(monkeypatch, calendar, "outlook_client", FakeClient())

    with pytest.raises(ValueError, match="Could not parse dates"):
        calendar.cal_availability(["a@example.com"], "not a date", "also not")


def test_cal_create_maps_arguments_and_reports_invitations(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(create_event=_event()))

    result = calendar.cal_create("Standup", "2026-10-15 09:00", "2026-10-15 09:30", location="Room 1", online_meeting=True)

    kwargs = fake.call("create_event")[1]
    assert kwargs["attendees"] is None
    assert kwargs["is_online_meeting"] is True
    assert result["invitationsSentTo"] == []
    assert "invitations" in calendar.cal_create.__doc__


def test_cal_respond_maps_tentative(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=_event()))

    result = calendar.cal_respond("ev-1", "tentative", comment="maybe", send_response=False)

    assert fake.call("respond_to_event") == (
        ("ev-1", "tentativelyAccept"),
        {"comment": "maybe", "send_response": False},
    )
    assert result["responseSent"] is False
    with pytest.raises(ValueError):
        calendar.cal_respond("ev-1", "maybe")  # type: ignore[arg-type]


def test_contact_search(monkeypatch):
    person = {
        "Id": "p1",
        "DisplayName": "Ann",
        "ScoredEmailAddresses": [{"Address": "ann@example.com"}],
        "PersonType": {"Subclass": "OrganizationUser"},
    }
    fake = use_client(monkeypatch, contacts, "outlook_client", FakeClient(search_people=[person]))

    assert contacts.contact_search("ann", count=3) == [
        {"id": "p1", "displayName": "Ann", "emails": ["ann@example.com"], "type": "Org directory"}
    ]
    assert fake.call("search_people") == (("ann",), {"top": 3})


def test_task_tools(monkeypatch):
    task = {"Id": "t1", "Subject": "Do it", "Status": "NotStarted", "DueDateTime": {"DateTime": "2026-10-20"}}
    fake = FakeClient(list_tasks=[task], create_task=task)
    use_client(monkeypatch, tasks, "outlook_client", fake)

    assert tasks.task_list()[0]["due"] == "2026-10-20"
    assert tasks.task_create("  Do it ")["id"] == "t1"
    assert fake.call("create_task")[0] == ("Do it",)
    assert tasks.task_complete("t1") == {"id": "t1", "status": "Completed"}
    with pytest.raises(ValueError):
        tasks.task_create("   ")


def _token(expires_in: int) -> str:
    return jwt.encode({"exp": time.time() + expires_in, "upn": "ross@example.com"}, "k" * 32, algorithm="HS256")


def test_auth_status_reports_validity_without_reauth(monkeypatch, tmp_path: Path):
    tokens = tmp_path / "tokens.json"
    tokens.write_text(json.dumps({"tokens": {config.OUTLOOK_TOKEN_DOMAIN: _token(3600), config.GRAPH_TOKEN_DOMAIN: _token(-60)}}))
    monkeypatch.setattr(config, "TOKENS_FILE", tokens)

    outlook, graph, copilot = auth.auth_status()["tokens"]

    assert outlook["valid"] is True and 59 <= outlook["minutesLeft"] <= 60
    assert outlook["user"] == "ross@example.com"
    assert graph["valid"] is False and graph["expired"] is True
    assert copilot["domain"] == config.COPILOT_TOKEN_DOMAIN and copilot["present"] is False


def test_auth_status_with_no_tokens_file(monkeypatch, tmp_path: Path):
    monkeypatch.setattr(config, "TOKENS_FILE", tmp_path / "none.json")

    assert [t["present"] for t in auth.auth_status()["tokens"]] == [False, False, False]
