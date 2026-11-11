from __future__ import annotations

import sys
from unittest.mock import Mock

import httpx
import pytest
from test_mail_organise import FakeClient as GraphClient
from test_mail_organise import rule
from test_mail_query import BASE, Client, message

from outlook_draft import cli
from outlook_draft.mail_organise import GRAPH_ROOT
from outlook_draft.outlook_client import OutlookClient


@pytest.mark.parametrize("path", [
    "/me/../../../../owa/service.svc", "/me/%2e%2e/%2e%2e/%2e%2e/owa/service.svc",
    "/me/%252e%252e/%252e%252e/beta", "/me/..\\..\\owa", "/me/messages\n",
])
def test_mail_continuation_rejects_normalisation_escape(path):
    requests = []
    client = OutlookClient(Mock())
    client._client = httpx.Client(transport=httpx.MockTransport(lambda request: requests.append(request)))
    try:
        with pytest.raises(ValueError):
            client.query_mail(page=BASE + path)
        assert not requests
    finally:
        client.close()


@pytest.mark.parametrize("url", [
    GRAPH_ROOT + "/../beta/me/messages", GRAPH_ROOT + "/me/../../beta/me/messages",
    GRAPH_ROOT + "/me/%2e%2e/%2e%2e/beta/me/messages",
])
def test_graph_continuation_rejects_normalisation_escape(url):
    client = GraphClient([{"value": [], "@odata.nextLink": url}])
    with pytest.raises(ValueError, match="pagination URL"):
        client.list_mail_categories()
    assert len(client.calls) == 1


@pytest.mark.parametrize("changes", [
    {"conditions": {"subjectContains": ["a"]}}, {"displayName": "Broader rule"},
])
@pytest.mark.parametrize("actions", [{"forwardTo": [{"emailAddress": {"address": "external@example.test"}}]},
                                      {"delete": True}, {"permanentDelete": True}])
def test_active_rule_patch_cannot_preserve_dangerous_actions(changes, actions):
    client = GraphClient([rule(isEnabled=True, actions=actions)])
    with pytest.raises(ValueError):
        client.update_mail_rule("rule-id", changes, confirm=True)
    assert [call[0] for call in client.calls] == ["GET"]


def test_active_safe_rule_patch_checks_effective_rule():
    client = GraphClient([rule(isEnabled=True), {}])
    client.update_mail_rule("rule-id", {"conditions": {"subjectContains": ["Client"]}}, confirm=True)
    assert [call[0] for call in client.calls] == ["GET", "PATCH"]


def test_inline_only_attachments_are_included():
    client = Client({"ConversationId": "c1"}, {"value": [message(HasAttachments=False)]},
                    {"value": [{"Id": "inline", "IsInline": True, "Name": "image.png"}]})
    result = client.mail_conversation("m1")
    assert result["complete"]
    assert result["messages"][0]["Attachments"][0]["IsInline"] is True


@pytest.mark.parametrize("restriction", ["--readonly", "--no-send", "OUTLOOK_CLI_READONLY", "OUTLOOK_CLI_NO_SEND"])
def test_restricted_gateway_cannot_spawn_unrestricted_children(monkeypatch, capsys, restriction):
    forbidden = Mock(side_effect=AssertionError("No gateway start"))
    monkeypatch.setattr(cli, "_dispatch", forbidden)
    flags = []
    if restriction.startswith("--"):
        flags.append(restriction)
    else:
        monkeypatch.setenv(restriction, "1")
    monkeypatch.setattr(sys, "argv", ["outlook-cli", *flags, "gateway", "start", "--self-chat"])
    with pytest.raises(SystemExit) as error:
        cli.main()
    assert error.value.code == 1
    assert "blocks gateway start" in capsys.readouterr().out
    forbidden.assert_not_called()
