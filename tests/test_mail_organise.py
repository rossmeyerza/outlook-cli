"""Fake-service checks for Graph mail categories and inbox rules."""

from __future__ import annotations

import argparse
from contextlib import contextmanager
from unittest.mock import Mock

import httpx
import pytest

from outlook_draft.commands import mail_organise as cli
from outlook_draft.errors import CliError, OutlookAPIError
from outlook_draft.mail_organise import CATEGORY_PATH, GRAPH_ROOT, RULE_PATH, MailOrganiseMixin, validate_rule
from outlook_draft.mcp import mail_organise as mcp


class FakeClient(MailOrganiseMixin):
    def __init__(self, responses=(), base_url=GRAPH_ROOT):
        self._base_url = base_url
        self.responses = list(responses)
        self.calls = []
        self.closed = False

    def _request(self, method, path, **kwargs):
        self.calls.append((method, path, kwargs))
        response = self.responses.pop(0) if self.responses else {}
        if isinstance(response, Exception):
            raise response
        if response is None:
            return httpx.Response(204)
        return httpx.Response(200, json=response)

    def close(self):
        self.closed = True


def rule(**changes):
    payload = {"displayName": "Clients", "sequence": 1,
               "conditions": {"fromAddresses": [{"emailAddress": {"address": "person@example.com"}}],
                              "subjectContains": ["Project"]},
               "actions": {"moveToFolder": "folder-id", "assignCategories": ["Clients"],
                           "markAsRead": True, "stopProcessingRules": True}}
    return {**payload, **changes}


@pytest.mark.parametrize("value", ["none", "None", "preset0", "Preset24"])
def test_category_colour_contract(value):
    client = FakeClient()
    result = client.create_mail_category("Clients", value)
    assert result["preview"] is True
    assert result["path"] == "/me/outlook/masterCategories"
    assert result["payload"] == {"displayName": "Clients", "color": value.lower()}
    assert not client.calls


@pytest.mark.parametrize("value", ["preset25", "preset-1", "blue", "", None, 2])
def test_category_invalid_colour(value):
    client = FakeClient()
    with pytest.raises(ValueError):
        client.create_mail_category("Clients", value, confirm=True)
    assert not client.calls


@pytest.mark.parametrize("method,args,verb", [
    ("create_mail_category", ("Clients", "preset1"), "POST"),
    ("update_mail_category", ("a/b+?", "none"), "PATCH"),
    ("delete_mail_category", ("a/b+?",), "DELETE"),
    ("create_mail_rule", (rule(),), "POST"),
    ("update_mail_rule", ("a/b+?", {"isEnabled": False}), "PATCH"),
    ("delete_mail_rule", ("a/b+?",), "DELETE"),
])
def test_mutations_preview_and_no_retries(method, args, verb):
    client = FakeClient([None])
    preview = getattr(client, method)(*args)
    assert preview["preview"]
    assert not client.calls
    result = getattr(client, method)(*args, confirm=True)
    assert result == {"preview": False, "result": None}
    assert len(client.calls) == 1
    assert client.calls[0][0] == verb
    assert client.calls[0][2]["max_retries"] == 0
    if method.startswith(("update", "delete")):
        assert client.calls[0][1].endswith("/a%2Fb%2B%3F")


def test_create_rule_disabled_by_default_and_explicit_enabled():
    client = FakeClient()
    assert client.create_mail_rule(rule())["payload"]["isEnabled"] is False
    assert client.create_mail_rule(rule(isEnabled=True))["payload"]["isEnabled"] is True


@pytest.mark.parametrize("payload", [
    rule(actions={"forwardTo": []}), rule(actions={"delete": False}),
    rule(actions={"permanentDelete": True}), rule(actions={"redirectTo": []}),
    rule(actions={"copyToFolder": "id"}), rule(conditions={"bodyContains": ["x"]}),
    rule(conditions={}), rule(actions={}), rule(actions={"markAsRead": False}),
    rule(actions={"markAsRead": 1}), rule(actions={"moveToFolder": ""}),
    rule(actions={"assignCategories": "not-a-list"}), rule(sequence=True),
    rule(sequence=-1), rule(sequence=2147483648), rule(isEnabled="false"),
    rule(unknown=True), rule(displayName="\n"), rule(conditions={"subjectContains": []}),
    rule(conditions={"fromAddresses": ["person@example.com"]}),
    rule(conditions={"fromAddresses": [{"emailAddress": {"address": "bad-address"}}]}),
    rule(conditions={"fromAddresses": [{"emailAddress": {"address": "p@example.com", "other": "x"}}]}),
])
def test_rule_validation_rejects_unknown_or_invalid_fields(payload):
    client = FakeClient()
    with pytest.raises(ValueError):
        client.create_mail_rule(payload, confirm=True)
    assert not client.calls


def test_rule_update_rejects_empty_and_create_requires_fields():
    with pytest.raises(ValueError, match="at least one"):
        validate_rule({})
    with pytest.raises(ValueError, match="Missing rule fields"):
        validate_rule({"displayName": "Clients"}, create=True)


@pytest.mark.parametrize("extra", [
    {"actions": {"forwardTo": [{"emailAddress": {"address": "p@example.com"}}]}},
    {"isReadOnly": True}, {"hasError": True}, {"exceptions": {"subjectContains": ["x"]}},
])
def test_enable_existing_unsafe_rule_fails_before_write(extra):
    client = FakeClient([{**rule(), **extra}])
    with pytest.raises(ValueError):
        client.update_mail_rule("id", {"isEnabled": True}, confirm=True)
    assert [call[0] for call in client.calls] == ["GET"]


def test_enable_existing_safe_rule_validates_and_previews():
    client = FakeClient([rule(), rule(), {}])
    preview = client.update_mail_rule("id", {"isEnabled": True})
    assert preview["preview"]
    client.update_mail_rule("id", {"isEnabled": True}, confirm=True)
    assert [call[0] for call in client.calls] == ["GET", "GET", "PATCH"]


@pytest.mark.parametrize("method,path", [("list_mail_categories", CATEGORY_PATH), ("list_mail_rules", RULE_PATH)])
def test_lists_follow_graph_pagination(method, path):
    next_url = GRAPH_ROOT + path + "?$skip=1"
    client = FakeClient([{"value": [{"id": "first"}], "@odata.nextLink": next_url},
                         {"value": [{"id": "second"}]}])
    assert getattr(client, method)() == [{"id": "first"}, {"id": "second"}]
    assert [call[1] for call in client.calls] == [path, next_url]


@pytest.mark.parametrize("url", ["https://evil.example/v1.0/items", "/me/items", "https://graph.microsoft.com/beta/me"])
def test_pagination_cannot_escape_graph(url):
    client = FakeClient([{"value": [], "@odata.nextLink": url}])
    with pytest.raises(ValueError, match="pagination URL"):
        client.list_mail_categories()
    assert len(client.calls) == 1


def test_pagination_cycle_fails_explicitly():
    url = GRAPH_ROOT + CATEGORY_PATH
    client = FakeClient([{"@odata.nextLink": url}, {"@odata.nextLink": url}])
    with pytest.raises(ValueError, match="repeated"):
        client.list_mail_categories()


@pytest.mark.parametrize("action", [
    lambda c: c.list_mail_categories(), lambda c: c.create_mail_category("Clients"),
    lambda c: c.categorise_messages(["id"], ["Clients"]), lambda c: c.list_mail_rules(),
])
def test_outlook_client_base_is_rejected(action):
    client = FakeClient(base_url="https://outlook.office.com/api/v2.0")
    with pytest.raises(ValueError, match="Graph v1.0"):
        action(client)
    assert not client.calls


def test_categorise_preview_preserves_other_categories():
    client = FakeClient([{"categories": ["Keep", "Remove"]}])
    result = client.categorise_messages(["a/b+?"], ["New"], ["Remove"])
    assert result["preview"] and not result["failed"]
    assert result["results"][0]["categories"] == ["Keep", "New"]
    assert result["results"][0]["status"] == "preview"
    assert len(client.calls) == 1
    assert client.calls[0][1] == "/me/messages/a%2Fb%2B%3F"


def test_categorise_write_uses_graph_contract_and_etag():
    client = FakeClient([{"categories": ["Keep"], "@odata.etag": 'W/"123"'}, {}])
    result = client.categorise_messages(["message-id"], ["New"], confirm=True)
    assert result["results"][0]["status"] == "updated"
    assert client.calls[1] == ("PATCH", "/me/messages/message-id", {
        "json_body": {"categories": ["Keep", "New"]}, "max_retries": 0,
        "extra_headers": {"If-Match": 'W/"123"'}})


def test_categorise_unchanged_skips_write_and_deduplicates_ids():
    client = FakeClient([{"categories": ["Keep"]}])
    result = client.categorise_messages(["id", "id"], ["Keep"], confirm=True)
    assert result["results"][0]["status"] == "unchanged"
    assert len(result["results"]) == len(client.calls) == 1


@pytest.mark.parametrize("status,uncertain", [(0, True), (503, True), (412, False), (403, False)])
def test_categorise_stops_after_write_error(status, uncertain):
    client = FakeClient([{"categories": []}, {}, {"categories": []}, OutlookAPIError(status, "failure")])
    result = client.categorise_messages(["first", "second", "third"], ["Clients"], confirm=True)
    assert result["failed"] and result["stopped"]
    assert [item["status"] for item in result["results"]] == ["updated", "write_failed", "unattempted"]
    assert result["results"][1]["uncertain"] is uncertain
    assert len(client.calls) == 4
    if status == 403:
        assert "Mail.ReadWrite" in result["results"][1]["error"]


def test_categorise_read_error_reports_failure_and_continues():
    client = FakeClient([OutlookAPIError(404, "missing"), {"categories": []}, {}])
    result = client.categorise_messages(["missing-id", "present-id"], ["Clients"], confirm=True)
    assert [item["status"] for item in result["results"]] == ["read_failed", "updated"]
    assert result["failed"] and not result["stopped"]


@pytest.mark.parametrize("ids,add,remove", [([], ["x"], []), (["1"], ["x"], []),
    (["id"], [], []), (["id"], ["x"], ["x"]), (["id"], "x", []), (["id"], ["\n"], [])])
def test_categorise_rejects_invalid_input_before_reads(ids, add, remove):
    client = FakeClient()
    with pytest.raises(ValueError):
        client.categorise_messages(ids, add, remove)
    assert not client.calls


@pytest.mark.parametrize("method,scope", [("GET", "MailboxSettings.Read"), ("POST", "MailboxSettings.ReadWrite")])
def test_403_names_scope_and_has_no_fallback(method, scope):
    client = FakeClient([OutlookAPIError(403, "forbidden")])
    with pytest.raises(OutlookAPIError, match=scope) as error:
        client._organise_request(method, CATEGORY_PATH, max_retries=0)
    assert error.value.status == 403
    assert len(client.calls) == 1


def test_create_uncertain_post_is_not_retried():
    client = FakeClient([OutlookAPIError(0, "timeout")])
    with pytest.raises(OutlookAPIError):
        client.create_mail_rule(rule(), confirm=True)
    assert len(client.calls) == 1
    assert client.calls[0][2]["max_retries"] == 0


def parse_cli(arguments):
    parser = argparse.ArgumentParser()
    cli.register(parser.add_subparsers(required=True))
    return parser.parse_args(arguments)


@pytest.mark.parametrize("arguments", [
    ["categories", "create", "Clients"], ["categories", "update", "id", "--colour", "preset1"],
    ["categories", "delete", "id"], ["rules", "delete", "id"],
    ["rules", "update", "id", "--disabled"], ["categorise", "id", "--add", "Clients"],
])
def test_cli_mutation_defaults_preview(arguments):
    assert parse_cli(arguments).yes is False


def test_cli_rule_create_disabled_and_json_validation():
    args = parse_cli(["rules", "create", "Clients", "--conditions", '{"subjectContains":["x"]}',
                      "--actions", '{"markAsRead":true}'])
    assert cli._rule_payload(args)["isEnabled"] is False
    with pytest.raises(ValueError, match="Invalid JSON"):
        cli._json_object("{")
    with pytest.raises(ValueError, match="object"):
        cli._json_object("[]")


def test_cli_uses_graph_factory_and_closes(monkeypatch):
    client = FakeClient()
    factory, output = Mock(return_value=client), Mock()
    monkeypatch.setattr(cli.runtime, "get_graph_client", factory)
    monkeypatch.setattr(cli.runtime, "print_json", output)
    args = parse_cli(["categories", "create", "Clients", "--json"])
    args.func(args)
    factory.assert_called_once_with()
    assert output.call_args.args[0]["preview"] is True
    assert client.closed and not client.calls


def test_cli_partial_failure_has_nonzero_exit_after_result(monkeypatch):
    client = FakeClient([{"categories": []}, OutlookAPIError(0, "timeout")])
    output = Mock()
    monkeypatch.setattr(cli.runtime, "get_graph_client", lambda: client)
    monkeypatch.setattr(cli.runtime, "print_json", output)
    with pytest.raises(CliError, match="Some messages failed"):
        cli.cmd_mail_organise(parse_cli(["categorise", "id", "--add", "Clients", "--yes"]))
    assert output.call_args.args[0]["failed"]
    assert client.closed


def test_mcp_uses_graph_factory_and_confirmation(monkeypatch):
    client = FakeClient()

    @contextmanager
    def factory():
        yield client

    monkeypatch.setattr(mcp, "graph_client", factory)
    assert mcp.mail_category_create("Clients")["preview"]
    assert not client.calls
    assert mcp.mail_category_create("Clients", confirm=True)["preview"] is False
    assert client.calls[0][0] == "POST"
    result = mcp.mail_rule_create("Clients", rule()["conditions"], rule()["actions"])
    assert result["payload"]["isEnabled"] is False
    assert len(mcp.TOOLS) == 9
