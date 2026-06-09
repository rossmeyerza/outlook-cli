from __future__ import annotations

from typing import Any

import httpx
import pytest

from outlook_draft import config
from outlook_draft.errors import SendingDisabledError
from outlook_draft.outlook_client import OutlookClient


class DummyTokenManager:
    token = "token"

    def force_reload(self) -> None:
        pass


class RecordingClient(OutlookClient):
    def __init__(self) -> None:
        super().__init__(DummyTokenManager())  # type: ignore[arg-type]
        self.calls: list[dict[str, Any]] = []

    def _request(
        self,
        method: str,
        path: str,
        *,
        params: dict[str, str] | None = None,
        json_body: dict | None = None,
        extra_headers: dict[str, str] | None = None,
        max_retries: int = 2,
    ) -> httpx.Response:
        self.calls.append(
            {
                "method": method,
                "path": path,
                "params": params,
                "json_body": json_body,
                "extra_headers": extra_headers,
                "max_retries": max_retries,
            }
        )
        return httpx.Response(200, json={"value": [], "id": "item-1"})


def test_calendar_find_rooms() -> None:
    client = RecordingClient()

    client.find_rooms(room_list="rooms@example.com")

    assert client.calls[0]["method"] == "GET"
    assert client.calls[0]["path"] == "/me/findRooms"
    assert client.calls[0]["params"] == {"roomList": "rooms@example.com"}


def test_calendar_get_schedule_uses_configured_timezone(monkeypatch) -> None:
    monkeypatch.setattr(config, "OUTLOOK_TIMEZONE", "GMT Standard Time")
    client = RecordingClient()

    client.get_schedule(
        schedules=["a@example.com"],
        start_dt="2026-05-20T09:00:00",
        end_dt="2026-05-20T17:00:00",
        interval_minutes=15,
    )

    call = client.calls[0]
    assert call["method"] == "POST"
    assert call["path"] == "/me/calendar/getSchedule"
    assert call["json_body"]["Schedules"] == ["a@example.com"]
    assert call["json_body"]["AvailabilityViewInterval"] == 15
    assert call["json_body"]["StartTime"] == {
        "DateTime": "2026-05-20T09:00:00",
        "TimeZone": "GMT Standard Time",
    }


def test_calendar_find_meeting_times_payload(monkeypatch) -> None:
    monkeypatch.setattr(config, "OUTLOOK_TIMEZONE", "GMT Standard Time")
    client = RecordingClient()

    client.find_meeting_times(
        attendees=["a@example.com", "b@example.com"],
        start_dt="2026-05-20T09:00:00",
        end_dt="2026-05-20T17:00:00",
        duration_minutes=45,
        max_candidates=5,
    )

    payload = client.calls[0]["json_body"]
    assert client.calls[0]["path"] == "/me/findMeetingTimes"
    assert payload["MeetingDuration"] == "PT45M"
    assert payload["MaxCandidates"] == 5
    assert payload["Attendees"] == [
        {"EmailAddress": {"Address": "a@example.com"}, "Type": "Required"},
        {"EmailAddress": {"Address": "b@example.com"}, "Type": "Required"},
    ]


def test_contact_create_and_update_payloads() -> None:
    client = RecordingClient()

    client.create_contact(
        display_name="Ada Lovelace",
        email="ada@example.com",
        given_name="Ada",
        surname="Lovelace",
        company="Maths",
        mobile_phone="123",
    )
    client.update_contact("contact-1", display_name="Ada L", email="ada.l@example.com")

    assert client.calls[0]["method"] == "POST"
    assert client.calls[0]["path"] == "/me/contacts"
    assert client.calls[0]["json_body"]["DisplayName"] == "Ada Lovelace"
    assert client.calls[0]["json_body"]["EmailAddresses"] == [
        {"Address": "ada@example.com", "Name": "Ada Lovelace"}
    ]
    assert client.calls[1]["method"] == "PATCH"
    assert client.calls[1]["path"] == "/me/contacts/contact-1"
    assert client.calls[1]["json_body"]["EmailAddresses"] == [
        {"Address": "ada.l@example.com", "Name": "Ada L"}
    ]


def test_mailbox_settings_methods() -> None:
    client = RecordingClient()

    client.get_mailbox_settings()
    client.update_mailbox_settings(
        time_zone="GMT Standard Time",
        automatic_replies={"Status": "Disabled"},
    )

    assert client.calls[0]["method"] == "GET"
    assert client.calls[0]["path"] == "/me/mailboxSettings"
    assert client.calls[1]["method"] == "PATCH"
    assert client.calls[1]["path"] == "/me/mailboxSettings"
    assert client.calls[1]["json_body"] == {
        "TimeZone": "GMT Standard Time",
        "AutomaticRepliesSetting": {"Status": "Disabled"},
    }


def test_update_task_payload(monkeypatch) -> None:
    monkeypatch.setattr(config, "OUTLOOK_TIMEZONE", "GMT Standard Time")
    client = RecordingClient()

    client.update_task(
        "task-1",
        subject="Updated",
        due_dt="2026-05-20T17:00:00",
        importance="High",
        status="InProgress",
    )

    assert client.calls[0]["method"] == "PATCH"
    assert client.calls[0]["path"] == "/me/tasks/task-1"
    assert client.calls[0]["json_body"] == {
        "Subject": "Updated",
        "DueDateTime": {"DateTime": "2026-05-20T17:00:00", "TimeZone": "GMT Standard Time"},
        "Importance": "High",
        "Status": "InProgress",
    }


def test_send_teams_message_is_disabled_for_agent_safety() -> None:
    client = RecordingClient()

    with pytest.raises(SendingDisabledError):
        client.send_teams_message("chat-1", "Hello", content_type="text")

    assert client.calls == []


def test_list_teams_chats_caps_page_size_to_graph_limit() -> None:
    client = RecordingClient()

    client.list_teams_chats(top=55)

    assert client.calls[0]["path"] == "/me/chats"
    assert client.calls[0]["params"]["$top"] == "50"
    assert client.calls[0]["params"]["$orderby"] == "lastMessagePreview/createdDateTime desc"



def test_list_teams_chat_members_follows_paging_without_top_query() -> None:
    class PagingClient(RecordingClient):
        def _request(
            self,
            method: str,
            path: str,
            *,
            params: dict[str, str] | None = None,
            json_body: dict | None = None,
            extra_headers: dict[str, str] | None = None,
            max_retries: int = 2,
        ) -> httpx.Response:
            self.calls.append(
                {
                    "method": method,
                    "path": path,
                    "params": params,
                    "json_body": json_body,
                    "extra_headers": extra_headers,
                    "max_retries": max_retries,
                }
            )
            if path == "/chats/chat-1/members":
                return httpx.Response(
                    200,
                    json={
                        "value": [{"id": f"member-{i}"} for i in range(50)],
                        "@odata.nextLink": "https://graph.microsoft.com/v1.0/chats/chat-1/members?page=2",
                    },
                )
            return httpx.Response(200, json={"value": [{"id": "member-50"}]})

    client = PagingClient()

    members = client.list_teams_chat_members("chat-1", top=51)

    assert len(members) == 51
    assert client.calls[0]["path"] == "/chats/chat-1/members"
    assert client.calls[0]["params"] is None
    assert client.calls[1]["path"] == "https://graph.microsoft.com/v1.0/chats/chat-1/members?page=2"
    assert client.calls[1]["params"] is None


def test_teams_sender_handles_null_from() -> None:
    from outlook_draft.commands.teams import _teams_sender

    assert _teams_sender({"from": None}) == "Unknown"
