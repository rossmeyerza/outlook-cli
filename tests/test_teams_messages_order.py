from __future__ import annotations

from typing import Any

import httpx

from outlook_draft.outlook_client import OutlookClient

NEXT = "https://graph.microsoft.com/v1.0/chats/chat-1/messages?$skiptoken=abc"


class DummyTokenManager:
    token = "token"

    def force_reload(self) -> None:
        pass


class PagingClient(OutlookClient):
    """Returns a first page of 50 messages with a next link, then a second page."""

    def __init__(self, first: int = 50, second: int = 20) -> None:
        super().__init__(DummyTokenManager())  # type: ignore[arg-type]
        self.calls: list[dict[str, Any]] = []
        self.first = first
        self.second = second

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
        self.calls.append({"method": method, "path": path, "params": params, "extra_headers": extra_headers})
        if len(self.calls) == 1:
            value = [{"id": f"m{i}"} for i in range(self.first)]
            return httpx.Response(200, json={"value": value, "@odata.nextLink": NEXT})
        value = [{"id": f"n{i}"} for i in range(self.second)]
        return httpx.Response(200, json={"value": value})


def test_messages_requested_newest_first() -> None:
    client = PagingClient()

    client.list_teams_messages("chat-1", top=10)

    first = client.calls[0]
    assert first["path"] == "/chats/chat-1/messages"
    assert first["params"] == {"$top": "10", "$orderby": "createdDateTime desc"}
    assert first["extra_headers"] == {"Prefer": "include-unknown-enum-members"}
    assert len(client.calls) == 1


def test_messages_follow_next_link_above_fifty() -> None:
    client = PagingClient()

    messages = client.list_teams_messages("chat-1", top=60)

    assert len(messages) == 60
    assert client.calls[0]["params"]["$top"] == "50"
    assert client.calls[1]["path"] == NEXT
    assert client.calls[1]["params"] is None


def test_messages_stop_when_no_next_link() -> None:
    client = PagingClient(first=50, second=5)

    messages = client.list_teams_messages("chat-1", top=200)

    assert len(messages) == 55
    assert len(client.calls) == 2


def test_message_metadata_uses_same_order() -> None:
    client = PagingClient()

    client.list_teams_message_metadata("chat-1", top=1)

    assert client.calls[0]["params"] == {"$top": "1", "$orderby": "createdDateTime desc"}
