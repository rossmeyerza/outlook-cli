from __future__ import annotations

from typing import Any

import httpx
import pytest

from outlook_draft.errors import OutlookAPIError, SendingDisabledError, TokenExpiredError
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
        return httpx.Response(200, json={"value": [], "Id": "message-1"})


def test_update_message_read_state_marks_read() -> None:
    client = RecordingClient()

    client.update_message_read_state("message-1", is_read=True)

    assert client.calls[0]["method"] == "PATCH"
    assert client.calls[0]["path"] == "/me/messages/message-1"
    assert client.calls[0]["json_body"] == {"IsRead": True}


def test_send_existing_draft_is_disabled_for_agent_safety() -> None:
    client = RecordingClient()

    with pytest.raises(SendingDisabledError):
        client.send_message("message-1")

    assert client.calls == []


def test_send_mail_is_disabled_for_agent_safety() -> None:
    client = RecordingClient()

    with pytest.raises(SendingDisabledError):
        client.send_mail(
            subject="Hello",
            body="Body",
            to=["a@example.com"],
            cc=["b@example.com"],
            content_type="HTML",
            save_to_sent_items=False,
            importance="High",
        )

    assert client.calls == []


def test_list_mail_folders_selects_folder_fields() -> None:
    client = RecordingClient()

    client.list_mail_folders(top=25)

    assert client.calls[0]["method"] == "GET"
    assert client.calls[0]["path"] == "/me/mailfolders"
    assert client.calls[0]["params"]["$top"] == "25"
    assert "displayName" in client.calls[0]["params"]["$select"]


def test_move_and_archive_message() -> None:
    client = RecordingClient()

    client.move_message("message-1", "folder-1")
    client.archive_message("message-2")

    assert client.calls[0]["path"] == "/me/messages/message-1/move"
    assert client.calls[0]["json_body"] == {"DestinationId": "folder-1"}
    assert client.calls[1]["path"] == "/me/messages/message-2/move"
    assert client.calls[1]["json_body"] == {"DestinationId": "archive"}


def test_attachment_methods() -> None:
    client = RecordingClient()

    client.list_message_attachments("message-1")
    client.get_message_attachment("message-1", "attachment-1")

    assert client.calls[0]["path"] == "/me/messages/message-1/attachments"
    assert client.calls[1]["path"] == "/me/messages/message-1/attachments/attachment-1"


def test_outlook_client_converts_token_failure_to_api_error() -> None:
    class ExpiredTokenManager:
        def get_token(self, *, auto_reauth: bool = False) -> str:
            raise TokenExpiredError("Outlook token has expired")

        def run_reauth(self, *, headless: bool = True) -> bool:
            return False

    client = OutlookClient(ExpiredTokenManager())  # type: ignore[arg-type]

    with pytest.raises(OutlookAPIError) as excinfo:
        client._request("GET", "/me")

    assert excinfo.value.status == 401
    assert "outlook-cli auth" in str(excinfo.value)
