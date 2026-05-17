from __future__ import annotations

from typing import Any

import httpx

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


def test_send_existing_draft_posts_send_action() -> None:
    client = RecordingClient()

    client.send_message("message-1")

    assert client.calls[0]["method"] == "POST"
    assert client.calls[0]["path"] == "/me/messages/message-1/send"


def test_send_mail_posts_payload() -> None:
    client = RecordingClient()

    client.send_mail(
        subject="Hello",
        body="Body",
        to=["a@example.com"],
        cc=["b@example.com"],
        content_type="HTML",
        save_to_sent_items=False,
        importance="High",
    )

    call = client.calls[0]
    assert call["method"] == "POST"
    assert call["path"] == "/me/sendmail"
    assert call["json_body"]["SaveToSentItems"] is False
    assert call["json_body"]["Message"]["Subject"] == "Hello"
    assert call["json_body"]["Message"]["Body"] == {"ContentType": "HTML", "Content": "Body"}
    assert call["json_body"]["Message"]["ToRecipients"] == [
        {"EmailAddress": {"Address": "a@example.com"}}
    ]
    assert call["json_body"]["Message"]["CcRecipients"] == [
        {"EmailAddress": {"Address": "b@example.com"}}
    ]
    assert call["json_body"]["Message"]["Importance"] == "High"


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
