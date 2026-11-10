from __future__ import annotations

from argparse import Namespace
from io import StringIO

import httpx
import pytest
from mcp_fakes import FakeClient, use_client
from rich.console import Console
from test_outlook_client_capabilities import RecordingClient

from outlook_draft import teams_new
from outlook_draft.commands import teams
from outlook_draft.errors import CliError, OutlookAPIError
from outlook_draft.mcp import sending

ME = {"id": "me-id", "displayName": "Ross Meyer", "mail": "ross@example.com"}
ANA = {"id": "11111111-1111-1111-1111-111111111111", "displayName": "Ana Diaz", "mail": "ana@example.com"}
BEN = {"id": "22222222-2222-2222-2222-222222222222", "displayName": "Ben Ode", "mail": "ben@example.com"}
DIRECTORY = {"ana@example.com": ANA, "ben@example.com": BEN, "ross@example.com": ME, ANA["id"]: ANA}


def _fake(**overrides) -> FakeClient:
    returns = {
        "get_current_user": ME,
        "get_directory_user": lambda ref: DIRECTORY.get(ref),
        "search_people": [],
        "create_teams_chat": {"id": "19:new@thread.v2", "webUrl": "https://teams/chat"},
        "send_teams_message": {"id": "msg-1"},
    }
    returns.update(overrides)
    return FakeClient(**returns)


# ── client payloads ──────────────────────────────────────────────────

def test_create_one_on_one_chat_payload() -> None:
    client = RecordingClient()
    client.create_teams_chat("me-id", ["ana-id"])
    call = client.calls[0]
    assert (call["method"], call["path"], call["max_retries"]) == ("POST", "/chats", 0)
    body = call["json_body"]
    assert body["chatType"] == "oneOnOne"
    assert "topic" not in body
    binds = [member["user@odata.bind"] for member in body["members"]]
    assert binds == [
        "https://graph.microsoft.com/v1.0/users('me-id')",
        "https://graph.microsoft.com/v1.0/users('ana-id')",
    ]
    assert all(member["roles"] == ["owner"] for member in body["members"])


def test_create_group_chat_payload_with_topic() -> None:
    client = RecordingClient()
    client.create_teams_chat("me-id", ["a", "b"], topic="Launch")
    body = client.calls[0]["json_body"]
    assert body["chatType"] == "group"
    assert body["topic"] == "Launch"
    assert len(body["members"]) == 3


def test_create_chat_rejects_topic_on_one_on_one_before_calling_graph() -> None:
    client = RecordingClient()
    with pytest.raises(ValueError, match="topic"):
        client.create_teams_chat("me-id", ["a"], topic="x")
    assert client.calls == []


def test_directory_user_falls_back_to_mail_filter_on_404() -> None:
    class Client(RecordingClient):
        def _request(self, method, path, **kwargs):
            super()._request(method, path, **kwargs)
            if path.startswith("/users/"):
                raise OutlookAPIError(404, "not found")
            return httpx.Response(200, json={"value": [ANA]})

    client = Client()
    assert client.get_directory_user("o'neil@example.com") == ANA
    assert client.calls[0]["path"] == "/users/o%27neil@example.com"
    assert client.calls[1]["params"]["$filter"] == "mail eq 'o''neil@example.com'"


# ── recipient resolution ─────────────────────────────────────────────

def test_split_refs_handles_commas_and_repeats() -> None:
    assert teams_new.split_refs(["a@x.com, b@x.com", " c@x.com "]) == ["a@x.com", "b@x.com", "c@x.com"]


def test_resolve_drops_self_and_duplicates() -> None:
    fake = _fake()
    me, users = teams_new.resolve_recipients(
        fake, ["ana@example.com", ANA["id"], "ross@example.com", "ben@example.com"],
    )
    assert me == ME
    assert users == [ANA, BEN]


def test_resolve_only_self_is_an_error() -> None:
    with pytest.raises(ValueError, match="only you"):
        teams_new.resolve_recipients(_fake(), ["ross@example.com"])


def test_unknown_email_is_an_error() -> None:
    with pytest.raises(ValueError, match="not a user"):
        teams_new.resolve_recipients(_fake(), ["guest@other.com"])


def test_unique_name_resolves_through_people_search() -> None:
    people = [
        {"displayName": "Ana Diaz", "userPrincipalName": "ana@example.com"},
        {"displayName": "Banana Team", "scoredEmailAddresses": [{"address": "x@example.com"}]},
    ]
    _, users = teams_new.resolve_recipients(_fake(search_people=people), ["ana diaz"])
    assert users == [ANA]


def test_personal_contacts_do_not_make_a_name_ambiguous() -> None:
    people = [
        {"displayName": "Ana Diaz", "userPrincipalName": "ana@example.com",
         "personType": {"class": "Person", "subclass": "OrganizationUser"}},
        {"displayName": "Ana Diaz", "scoredEmailAddresses": [{"address": "ana@gmail.com"}],
         "personType": {"class": "Person", "subclass": "PersonalContact"}},
    ]
    _, users = teams_new.resolve_recipients(_fake(search_people=people), ["Ana Diaz"])
    assert users == [ANA]


def test_ambiguous_name_lists_candidates() -> None:
    people = [
        {"displayName": "Ana Diaz", "userPrincipalName": "ana@example.com"},
        {"displayName": "Ana Smith", "userPrincipalName": "ana.smith@example.com"},
    ]
    with pytest.raises(ValueError, match="more than one.*Ana Smith <ana.smith@example.com>"):
        teams_new.resolve_recipients(_fake(search_people=people), ["Ana"])


def test_send_failure_after_create_reports_chat_id() -> None:
    def fail(*args, **kwargs):
        raise OutlookAPIError(503, "uncertain")

    fake = _fake(send_teams_message=fail)
    with pytest.raises(teams_new.ChatCreatedSendFailed, match="19:new@thread.v2"):
        teams_new.start_chat(fake, ME, [ANA], "Hi")
    assert fake.names().count("create_teams_chat") == 1
    assert fake.names().count("send_teams_message") == 1


# ── CLI ──────────────────────────────────────────────────────────────

def _run_cli(monkeypatch, fake, **kwargs) -> str:
    output = StringIO()
    monkeypatch.setattr(teams.runtime, "get_graph_client", lambda: fake)
    monkeypatch.setattr(teams.runtime, "console", Console(file=output, width=200))
    args = {"to": ["ana@example.com"], "message": "Hi", "topic": None, "html": False, "yes": False}
    args.update(kwargs)
    teams.cmd_teams_new(Namespace(**args))
    return output.getvalue()


def test_cli_preview_sends_nothing(monkeypatch) -> None:
    fake = _fake()
    with pytest.raises(SystemExit) as exit_info:
        _run_cli(monkeypatch, fake, to=["ana@example.com,ben@example.com"], topic="Launch")
    assert exit_info.value.code == 2
    assert "create_teams_chat" not in fake.names()
    assert "send_teams_message" not in fake.names()
    assert "close" in fake.names()


def test_cli_yes_creates_chat_then_sends(monkeypatch) -> None:
    fake = _fake()
    rendered = _run_cli(monkeypatch, fake, yes=True, html=True)
    assert fake.call("create_teams_chat") == (("me-id", [ANA["id"]]), {"topic": None})
    assert fake.call("send_teams_message") == (("19:new@thread.v2", "Hi"), {"content_type": "html"})
    assert "Teams message sent" in rendered
    assert "19:new@thread.v2" in rendered


def test_cli_topic_on_one_on_one_is_an_error(monkeypatch) -> None:
    fake = _fake()
    with pytest.raises(CliError, match="topic"):
        _run_cli(monkeypatch, fake, topic="x", yes=True)
    assert "create_teams_chat" not in fake.names()


# ── MCP ──────────────────────────────────────────────────────────────

@pytest.mark.parametrize("confirm", [False, True])
def test_mcp_teams_new(monkeypatch, confirm) -> None:
    fake = use_client(monkeypatch, sending, "graph_client", _fake())
    result = sending.teams_new(["ana@example.com", "ben@example.com"], "Hi", topic="Launch", confirm=confirm)
    assert result["sent"] is confirm
    assert result["chatType"] == "group"
    assert [member["mail"] for member in result["members"]] == ["ana@example.com", "ben@example.com"]
    assert ("create_teams_chat" in fake.names()) is confirm
    assert ("send_teams_message" in fake.names()) is confirm
    if confirm:
        assert result["chat_id"] == "19:new@thread.v2"


def test_mcp_teams_new_rejects_empty_message(monkeypatch) -> None:
    fake = use_client(monkeypatch, sending, "graph_client", _fake())
    with pytest.raises(ValueError, match="empty"):
        sending.teams_new(["ana@example.com"], "  ", confirm=True)
    assert "create_teams_chat" not in fake.names()
