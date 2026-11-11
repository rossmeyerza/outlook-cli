from __future__ import annotations

import asyncio
import json
from unittest.mock import Mock

import httpx
import pytest
from mcp_fakes import FakeClient, use_client

from outlook_draft import cli
from outlook_draft.commands import mail_ops as commands
from outlook_draft.errors import CliError, OutlookAPIError
from outlook_draft.mail_query import MailQueryMixin, validate_filters
from outlook_draft.mcp import mail_ops
from outlook_draft.mcp.server import build_server
from outlook_draft.outlook_client import OutlookClient
from outlook_draft.safety import SafetyError, policy

BASE = "https://outlook.office.com/api/v2.0"
NEXT = BASE + "/me/messages?$skip=2"


class Client(MailQueryMixin):
    _base_url = BASE

    def __init__(self, *responses):
        self.responses = iter(responses)
        self.calls = []

    def _request(self, method, path, **kwargs):
        self.calls.append((method, path, kwargs))
        value = next(self.responses)
        if isinstance(value, Exception):
            raise value
        return httpx.Response(200, json=value)


def message(mid="m1", **fields):
    return {"Id": mid, "From": {"EmailAddress": {"Address": "a@example.test"}},
            "ToRecipients": [{"EmailAddress": {"Address": "b@example.test"}}],
            "IsRead": False, "HasAttachments": False, "ReceivedDateTime": "2026-11-01T12:00:00Z", **fields}


def test_query_all_pages_exact_filters():
    client = Client({"value": [message(), message("m2", IsRead=True)], "@odata.nextLink": NEXT},
                    {"value": [message("m3")]})
    report = client.query_mail(filters={"unread": True, "sender": "A@example.test",
                                       "recipient": "b@example.test", "after": "2026-11-01", "before": "2026-11-02"}, all_pages=True)
    assert report["complete"] and report["countExact"]
    assert report["matchedCount"] == 2 and report["scannedCount"] == 3
    assert client.calls[1] == ("GET", NEXT, {"params": None})


def test_query_single_page_does_not_claim_total():
    report = Client({"value": [message()], "@odata.nextLink": NEXT}).query_mail()
    assert report["matchedCount"] == 1
    assert not report["countExact"] and not report["complete"]
    assert report["nextPage"] == NEXT


def test_query_keyword_limit_is_explicit():
    report = Client({"value": [message()]}).query_mail("hello", all_pages=True)
    assert not report["countExact"] and "server-side" in report["warning"]


def test_query_continuation_count_is_not_whole_query():
    report = Client({"value": [message()]}).query_mail(page=NEXT)
    assert not report["countExact"]


@pytest.mark.parametrize("filters", [{"wat": 1}, {"sender": " "}, {"unread": "yes"},
                                         {"after": "bad"}, {"before": 5},
                                         {"after": "2026-11-02", "before": "2026-11-01"}])
def test_query_filter_validation(filters):
    with pytest.raises(ValueError):
        validate_filters(filters)


@pytest.mark.parametrize("url", ["https://attacker.test/me/messages", BASE + "/users/a/messages",
                                    "http://outlook.office.com/api/v2.0/me/messages", "/not-mail", BASE + "/me/messages#x"])
def test_query_rejects_untrusted_pages(url):
    client = Client()
    with pytest.raises(ValueError):
        client.query_mail(page=url)
    assert not client.calls


def test_query_cycle_and_page_size():
    client = Client({"value": [], "@odata.nextLink": NEXT}, {"value": [], "@odata.nextLink": NEXT})
    with pytest.raises(ValueError, match="repeated"):
        client.query_mail(all_pages=True)
    with pytest.raises(ValueError):
        Client().query_mail(page_size=0)


def test_query_folder_is_encoded():
    client = Client({"value": []})
    client.query_mail(folder="a/b+")
    assert client.calls[0][1] == "/me/mailfolders/a%2Fb%2B/messages"


def test_conversation_complete_and_attachment_pages():
    client = Client({"ConversationId": "c'1"},
                    {"value": [message("m2", ReceivedDateTime="2026-11-01T12:00:00Z")], "@odata.nextLink": NEXT},
                    {"value": [message("m1", HasAttachments=True, ReceivedDateTime="2026-11-01T13:00:00+02:00")]},
                    {"value": [{"Id": "a1"}], "@odata.nextLink": BASE + "/me/messages/m1/attachments?$skip=1"},
                    {"value": [{"Id": "a2"}]}, {"value": []})
    report = client.mail_conversation("m1")
    assert report["complete"] and report["count"] == 2
    assert [item["Id"] for item in report["messages"]] == ["m1", "m2"]
    assert len(report["messages"][0]["Attachments"]) == 2
    assert client.calls[1][2]["params"]["$filter"] == "ConversationId eq 'c''1'"
    assert client.calls[-2][2]["params"] is None


def test_conversation_missing_id_and_cycle():
    with pytest.raises(ValueError, match="no conversation"):
        Client({}).mail_conversation("m")
    client = Client({"ConversationId": "c"}, {"value": [], "@odata.nextLink": NEXT},
                    {"value": [], "@odata.nextLink": NEXT})
    with pytest.raises(ValueError, match="repeated"):
        client.mail_conversation("m")


@pytest.mark.parametrize("action", ["archive", "trash", "move", "mark-read", "mark-unread"])
def test_bulk_preview_never_writes(action):
    client = Client()
    assert not client.bulk_mail(["m1", "m1"], action, folder="f")["confirmed"]
    assert not client.calls


@pytest.mark.parametrize("action,verb,body", [
    ("trash", "POST", {"DestinationId": "deleteditems"}),
    ("archive", "POST", {"DestinationId": "archive"}),
    ("move", "POST", {"DestinationId": "f"}),
    ("mark-read", "PATCH", {"IsRead": True}),
    ("mark-unread", "PATCH", {"IsRead": False}),
])
def test_bulk_confirmed_calls(action, verb, body):
    client = Client({"Id": "m-new"})
    report = client.bulk_mail(["m/1", "m/1"], action, folder="f", confirm=True)
    assert report["results"][0]["status"] == "applied"
    assert len(client.calls) == 1
    assert client.calls[0][0] == verb and "m%2F1" in client.calls[0][1]
    assert client.calls[0][2] == {"json_body": body, "max_retries": 0}


@pytest.mark.parametrize("status, outcome", [(0, "uncertain"), (503, "uncertain"), (403, "failed")])
def test_bulk_failure_stops_without_retry(status, outcome):
    client = Client({"Id": "new"}, OutlookAPIError(status, "failure"))
    report = client.bulk_mail(["a", "b", "c"], "trash", confirm=True)
    assert [r["status"] for r in report["results"]] == ["applied", outcome, "not_attempted"]
    assert len(client.calls) == 2


@pytest.mark.parametrize("ids, action, folder", [([], "trash", None), ([" "], "trash", None),
                                                  (["x"], "delete", None), (["x"], "move", None)])
def test_bulk_validation(ids, action, folder):
    with pytest.raises(ValueError):
        Client().bulk_mail(ids, action, folder=folder)


def test_mcp_bulk_preview_has_no_auth(monkeypatch):
    monkeypatch.setattr(mail_ops, "outlook_client", Mock(side_effect=AssertionError("No auth")))
    assert not mail_ops.mail_trash(["m1"])["confirmed"]


def test_mcp_query_and_conversation_shape(monkeypatch):
    fake = FakeClient(query_mail={"messages": [message()], "complete": True},
                      mail_conversation={"messages": [message(Body={"ContentType": "HTML", "Content": "<p>Hi</p>"})]})
    use_client(monkeypatch, mail_ops, "outlook_client", fake)
    assert mail_ops.mail_query()["messages"][0]["id"] == "m1"
    assert mail_ops.mail_conversation("m1")["messages"][0]["body"] == "Hi"


def test_cli_bulk_preview_and_partial_failure(monkeypatch, capsys):
    parser = cli.build_parser()
    args = parser.parse_args(["mail", "trash", "full-id"])
    monkeypatch.setattr(commands.runtime, "get_client", Mock(side_effect=AssertionError("No auth")))
    args.func(args)
    assert json.loads(capsys.readouterr().out)["confirmed"] is False
    args = parser.parse_args(["mail", "trash", "full-id", "--yes"])
    fake = Mock()
    fake.bulk_mail.return_value = {"results": [{"status": "uncertain"}]}
    monkeypatch.setattr(commands.runtime, "get_client", lambda: fake)
    with pytest.raises(CliError, match="stopped"):
        args.func(args)
    fake.close.assert_called_once()


def test_readonly_guards_real_client_before_network():
    client = OutlookClient(Mock())
    client._ensure_client = Mock(side_effect=AssertionError("No network"))
    with policy(readonly=True), pytest.raises(SafetyError):
        client._request("POST", "/me/messages", json_body={})
    client._ensure_client.assert_not_called()


@pytest.mark.parametrize("env, missing", [("OUTLOOK_CLI_READONLY", {"mail_bulk", "draft_create", "copilot_ask"}),
                                             ("OUTLOOK_CLI_NO_SEND", {"mail_send", "teams_send", "copilot_app"})])
def test_restricted_mcp_tools(monkeypatch, env, missing):
    monkeypatch.setenv(env, "1")
    names = {tool.name for tool in asyncio.run(build_server().list_tools())}
    assert not missing & names
    assert "mail_query" in names


def test_mcp_tool_allowlist_and_context_policy(monkeypatch):
    monkeypatch.setenv("OUTLOOK_CLI_ENABLE_TOOLS", "mail_query,mail_bulk")
    with policy(readonly=True):
        assert [tool.name for tool in asyncio.run(build_server().list_tools())] == ["mail_query"]
