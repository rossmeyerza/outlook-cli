"""Calendar removal tools preview first and never retry a write."""
from __future__ import annotations

import asyncio
from unittest.mock import Mock

import pytest
from mcp_fakes import FakeClient, use_client
from test_outlook_client_calendar import RecordingClient

from outlook_draft.errors import OutlookAPIError
from outlook_draft.mcp import calendar
from outlook_draft.mcp.server import build_server
from outlook_draft.outlook_client import OutlookClient
from outlook_draft.safety import SafetyError, policy


def event(**fields):
    return {"Id": "event-id", "Subject": "Check", "Type": "SingleInstance", "IsOrganizer": True,
            "ChangeKey": "version-1", "Start": {"DateTime": "2026-11-11T10:00:00"},
            "End": {"DateTime": "2026-11-11T10:30:00"}, "Attendees": [], **fields}


def meeting(**fields):
    return event(Attendees=[{"EmailAddress": {"Address": "person@example.test", "Name": "Person"}, "Type": "Required"}], **fields)


@pytest.mark.parametrize("operation", ["delete", "cancel"])
def test_preview_never_writes(monkeypatch, operation):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=meeting()))
    result = getattr(calendar, "cal_" + operation)("event-id")
    assert result["confirmed"] is False
    assert result["event"]["subject"] == "Check"
    assert result["scope"] == "single"
    assert result["attendees"][0]["address"] == "person@example.test"
    assert not {"delete_event", "cancel_event"} & set(fake.names())
    assert result["willNotifyAttendees"] is (operation == "cancel")


def test_delete_private_entry(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=event()))
    result = calendar.cal_delete("event-id", confirm=True, expected_change_key="version-1")
    assert result["confirmed"] and result["deleted"] and not result["cancelled"]
    assert fake.call("delete_event") == (("event-id",), {})
    assert "cancel_event" not in fake.names()


def test_delete_received_calendar_entry_does_not_send_decline(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=meeting(IsOrganizer=False)))
    result = calendar.cal_delete("event-id", confirm=True)
    assert result["deleted"] and not result["willNotifyAttendees"]
    assert "respond_to_event" not in fake.names()


@pytest.mark.parametrize("organiser", [True, None, "unknown"])
def test_delete_never_silently_cancels_organised_meeting(monkeypatch, organiser):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=meeting(IsOrganizer=organiser)))
    preview = calendar.cal_delete("event-id")
    assert not preview["canApply"] and "cal_cancel" in " ".join(preview["reasons"])
    with pytest.raises(ValueError, match="cal_cancel"):
        calendar.cal_delete("event-id", confirm=True)
    assert "delete_event" not in fake.names()


def test_cancel_organiser_meeting_reports_notifications(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=meeting()))
    result = calendar.cal_cancel("event-id", comment="No longer required", confirm=True)
    assert result["cancelled"] and not result["deleted"] and result["willNotifyAttendees"]
    assert result["comment"] == "No longer required"
    assert fake.call("cancel_event") == (("event-id",), {"comment": "No longer required"})


@pytest.mark.parametrize("fields, match", [({"IsOrganizer": False}, "organiser"),
                                             ({"IsOrganizer": None}, "organiser"),
                                             ({"IsCancelled": True}, "already cancelled"),
                                             ({"Attendees": []}, "cal_delete")])
def test_invalid_cancel_has_no_write(monkeypatch, fields, match):
    current = meeting()
    current.update(fields)
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=current))
    with pytest.raises(ValueError, match=match):
        calendar.cal_cancel("event-id", confirm=True)
    assert "cancel_event" not in fake.names()


@pytest.mark.parametrize("operation", ["delete", "cancel"])
def test_series_requires_separate_explicit_approval(monkeypatch, operation):
    current = meeting(Type="SeriesMaster") if operation == "cancel" else event(Type="SeriesMaster")
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=current))
    method = getattr(calendar, "cal_" + operation)
    preview = method("event-id")
    assert preview["scope"] == "series" and not preview["canApply"]
    with pytest.raises(ValueError, match="whole series"):
        method("event-id", confirm=True)
    assert operation + "_event" not in fake.names()
    assert method("event-id", allow_series=True, confirm=True)["confirmed"]


@pytest.mark.parametrize("kind", ["Occurrence", "Exception"])
def test_occurrence_never_follows_series_master(monkeypatch, kind):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=event(Type=kind, SeriesMasterId="master-id")))
    result = calendar.cal_delete("occurrence-id", confirm=True)
    assert result["scope"] == "occurrence"
    assert fake.call("delete_event")[0] == ("occurrence-id",)
    assert fake.names().count("get_event") == 1


def test_unknown_scope_and_stale_preview_fail_before_write(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=event(Type="FutureType")))
    with pytest.raises(ValueError, match="scope"):
        calendar.cal_delete("event-id", confirm=True)
    fake.returns["get_event"] = event(ChangeKey="version-2")
    with pytest.raises(ValueError, match="changed after"):
        calendar.cal_delete("event-id", confirm=True, expected_change_key="version-1")
    assert "delete_event" not in fake.names()


@pytest.mark.parametrize("fields", [{"Type": None}, {"Attendees": None}, {"Type": "SingleInstance", "Recurrence": {"Pattern": {}}}])
def test_incomplete_or_conflicting_metadata_never_permits_removal(monkeypatch, fields):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=event(**fields)))
    assert not calendar.cal_delete("event-id")["canApply"]
    with pytest.raises(ValueError):
        calendar.cal_delete("event-id", confirm=True)
    assert "delete_event" not in fake.names()


@pytest.mark.parametrize("value", ["", " ", "1", " 1 ", "x\n", None])
def test_bad_id_never_authenticates(monkeypatch, value):
    monkeypatch.setattr(calendar, "outlook_client", Mock(side_effect=AssertionError("No client")))
    with pytest.raises(ValueError):
        calendar.cal_delete(value, confirm=True)


@pytest.mark.parametrize("options", [{"confirm": "true"}, {"allow_series": 1}, {"expected_change_key": ""}])
def test_bad_options_never_authenticate(monkeypatch, options):
    monkeypatch.setattr(calendar, "outlook_client", Mock(side_effect=AssertionError("No client")))
    with pytest.raises(ValueError):
        calendar.cal_delete("event-id", **options)


@pytest.mark.parametrize("operation", ["delete", "cancel"])
def test_write_failure_has_no_retry(monkeypatch, operation):
    def fail(*args, **kwargs):
        raise OutlookAPIError(503, "Uncertain result")
    current = meeting() if operation == "cancel" else event()
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=current, **{operation + "_event": fail}))
    with pytest.raises(OutlookAPIError):
        getattr(calendar, "cal_" + operation)("event-id", confirm=True)
    assert fake.names().count(operation + "_event") == 1


def test_client_paths_are_encoded_and_writes_have_zero_retries():
    client = RecordingClient()
    try:
        client.get_event("event/a+b")
        client.delete_event("event/a+b")
        client.cancel_event("event/a+b", comment="Reason")
        assert client.calls[0]["path"] == "/me/events/event%2Fa%2Bb"
        assert "isOrganizer" in client.calls[0]["params"]["$select"]
        assert "changeKey" in client.calls[0]["params"]["$select"]
        assert client.calls[1]["path"] == "/me/events/event%2Fa%2Bb"
        assert client.calls[2]["path"] == "/me/events/event%2Fa%2Bb/cancel"
        assert [call["max_retries"] for call in client.calls[1:]] == [0, 0]
    finally:
        client.close()


@pytest.mark.parametrize("mode", ["readonly", "no_send"])
@pytest.mark.parametrize("operation", ["delete_event", "cancel_event"])
def test_policy_blocks_calendar_writes_before_network(mode, operation):
    client = RecordingClient()
    client._request = OutlookClient._request.__get__(client)
    client._ensure_client = Mock(side_effect=AssertionError("No network"))
    try:
        with policy(**{mode: True}), pytest.raises(SafetyError):
            getattr(client, operation)("event-id")
        client._ensure_client.assert_not_called()
    finally:
        client.close()


def test_new_tools_have_confirmation_and_destructive_annotations():
    tools = {tool.name: tool for tool in asyncio.run(build_server().list_tools())}
    for name in ("cal_delete", "cal_cancel"):
        tool = tools[name]
        assert tool.inputSchema["properties"]["confirm"]["default"] is False
        assert tool.inputSchema["properties"]["allow_series"]["default"] is False
        assert tool.annotations.destructiveHint is True
        assert tool.annotations.idempotentHint is False
    with policy(readonly=True):
        names = {tool.name for tool in asyncio.run(build_server().list_tools())}
    assert not {"cal_delete", "cal_cancel"} & names
