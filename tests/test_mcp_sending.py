from __future__ import annotations

import asyncio

import pytest
from mcp.shared.memory import create_connected_server_and_client_session
from mcp_fakes import FakeClient, use_client

from outlook_draft.errors import OutlookAPIError
from outlook_draft.mcp import sending
from outlook_draft.mcp.server import build_server


@pytest.fixture
def fake(monkeypatch):
    client = FakeClient(
        get_draft={"Id": "draft-1", "IsDraft": True, "ToRecipients": [{"EmailAddress": {"Address": "a@example.com"}}]},
        send_teams_message={"id": "msg-1"},
    )
    use_client(monkeypatch, sending, "outlook_client", client)
    use_client(monkeypatch, sending, "graph_client", client)
    monkeypatch.setattr(sending, "load_signature", lambda path: "<p>Signature</p>")
    return client


@pytest.mark.parametrize("confirm", [False, True])
def test_mail_send(fake, confirm):
    result = sending.mail_send(["a@example.com"], "Subject", "Body", cc=["b@example.com"], confirm=confirm)
    assert result["sent"] is confirm
    assert "Signature" in result["bodyHtml"]
    assert ("send_mail" in fake.names()) is confirm
    assert "create_draft" not in fake.names()
    if confirm:
        _, args = fake.call("send_mail")
        assert args["content_type"] == "HTML"
        assert args["save_to_sent_items"] is True
        assert args["cc"] == ["b@example.com"]


@pytest.mark.parametrize("confirm", [False, True])
def test_draft_send(fake, confirm):
    assert sending.draft_send("draft-1", confirm=confirm)["sent"] is confirm
    assert ("send_message" in fake.names()) is confirm


@pytest.mark.parametrize("confirm", [False, True])
def test_teams_send(fake, confirm):
    result = sending.teams_send("chat-1", "Hello", html=True, confirm=confirm)
    assert result["sent"] is confirm
    assert ("send_teams_message" in fake.names()) is confirm
    if confirm:
        assert fake.call("send_teams_message") == (("chat-1", "Hello"), {"content_type": "html"})


def test_draft_refuses_non_drafts_and_missing_recipients(fake):
    fake.returns["get_draft"] = {"IsDraft": False}
    with pytest.raises(ValueError, match="not a draft"):
        sending.draft_send("id", confirm=True)
    fake.returns["get_draft"] = {"IsDraft": True}
    with pytest.raises(ValueError, match="recipient"):
        sending.draft_send("id", confirm=True)
    assert "send_message" not in fake.names()


@pytest.mark.parametrize("call", [
    lambda: sending.mail_send([], "s", "b", confirm=True),
    lambda: sending.mail_send(["a@example.com"], "s", " ", confirm=True),
    lambda: sending.mail_send(["a@example.com"], "s", "b", importance="Urgent", confirm=True),
    lambda: sending.teams_send(" ", "hello", confirm=True),
    lambda: sending.teams_send("chat", " ", confirm=True),
    lambda: sending.draft_send(" ", confirm=True),
])
def test_invalid_input_does_not_send(fake, call):
    with pytest.raises(ValueError):
        call()
    assert not any(name.startswith("send_") for name in fake.names())


def test_send_failure_propagates_without_retry(fake):
    def fail(**kwargs):
        raise OutlookAPIError(503, "uncertain result")
    fake.returns["send_mail"] = fail
    with pytest.raises(OutlookAPIError):
        sending.mail_send(["a@example.com"], "s", "b", confirm=True)
    assert fake.names().count("send_mail") == 1


def test_mcp_preview_and_confirm(fake):
    async def scenario():
        async with create_connected_server_and_client_session(build_server()) as session:
            preview = await session.call_tool("teams_send", {"chat_id": "chat-1", "content": "Hello"})
            assert "send_teams_message" not in fake.names()
            sent = await session.call_tool("teams_send", {"chat_id": "chat-1", "content": "Hello", "confirm": True})
            return preview, sent
    preview, sent = asyncio.run(scenario())
    assert not preview.isError and not sent.isError
    assert preview.structuredContent["sent"] is False
    assert sent.structuredContent["sent"] is True
