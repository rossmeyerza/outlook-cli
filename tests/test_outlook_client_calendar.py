from __future__ import annotations

from typing import Any

import httpx

from outlook_draft import config
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
        return httpx.Response(200, json={"value": [], "id": "event-1"})


def test_create_event_uses_configured_outlook_timezone(monkeypatch) -> None:
    monkeypatch.setattr(config, "OUTLOOK_TIMEZONE", "GMT Standard Time")
    client = RecordingClient()

    client.create_event(
        "Subject",
        "2026-05-18T16:30:00",
        "2026-05-18T16:45:00",
        location="Room",
        body="Body",
        attendees=["a@example.com"],
    )

    call = client.calls[0]
    assert call["method"] == "POST"
    assert call["path"] == "/me/events"
    assert call["json_body"]["Start"] == {
        "DateTime": "2026-05-18T16:30:00",
        "TimeZone": "GMT Standard Time",
    }
    assert call["json_body"]["End"] == {
        "DateTime": "2026-05-18T16:45:00",
        "TimeZone": "GMT Standard Time",
    }


def test_agenda_requests_configured_timezone_header(monkeypatch) -> None:
    monkeypatch.setattr(config, "OUTLOOK_TIMEZONE", "GMT Standard Time")
    client = RecordingClient()

    client.get_agenda()

    assert client.calls[0]["extra_headers"] == {
        "Prefer": 'outlook.timezone="GMT Standard Time"'
    }


def test_respond_to_event_posts_expected_payload() -> None:
    client = RecordingClient()

    client.respond_to_event(
        "event-1",
        "tentativelyAccept",
        comment="Maybe",
        send_response=False,
    )

    assert client.calls[0]["method"] == "POST"
    assert client.calls[0]["path"] == "/me/events/event-1/tentativelyAccept"
    assert client.calls[0]["json_body"] == {
        "Comment": "Maybe",
        "SendResponse": False,
    }


def test_cancel_event_posts_comment() -> None:
    client = RecordingClient()

    client.cancel_event("event-1", comment="Cancelled")

    assert client.calls[0]["method"] == "POST"
    assert client.calls[0]["path"] == "/me/events/event-1/cancel"
    assert client.calls[0]["json_body"] == {"Comment": "Cancelled"}


def test_delete_event_uses_delete_method() -> None:
    client = RecordingClient()

    client.delete_event("event-1")

    assert client.calls[0]["method"] == "DELETE"
    assert client.calls[0]["path"] == "/me/events/event-1"
