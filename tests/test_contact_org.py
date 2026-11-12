"""CLI and MCP expose the same read-only organisation view."""
from __future__ import annotations

import argparse
import asyncio
import json
from unittest.mock import Mock

import pytest
from mcp_fakes import use_client
from rich.console import Console
from test_directory import PERSON, fake_view

from outlook_draft import cli, runtime
from outlook_draft.commands import contacts as contact_commands
from outlook_draft.errors import CliError, OutlookAPIError
from outlook_draft.mcp import contacts
from outlook_draft.mcp.server import build_server
from outlook_draft.safety import policy


def org_args(**extra):
    return argparse.Namespace(**{"user": "me", "chain": False, "max_depth": None, "json": True, "no_spinner": True, **extra})


def test_cli_json_matches_mcp_and_uses_graph(monkeypatch, capsys):
    fake = fake_view()
    monkeypatch.setattr(runtime, "get_graph_client", lambda: fake)
    monkeypatch.setattr(runtime, "get_client", Mock(side_effect=AssertionError("No Outlook REST client")))
    contact_commands.cmd_contact_org(org_args())
    actual = json.loads(capsys.readouterr().out)
    assert fake.names().count("close") == 1
    use_client(monkeypatch, contacts, "graph_client", fake_view())
    assert actual == contacts.contact_org()


def test_cli_table_preserves_directory_text_without_markup(monkeypatch, capsys):
    fake = fake_view(get_organisation_profile={**PERSON, "displayName": "[red]Person[/]"})
    monkeypatch.setattr(runtime, "get_graph_client", lambda: fake)
    monkeypatch.setattr(runtime, "console", Console(width=100, force_terminal=False, color_system=None))
    contact_commands.cmd_contact_org(org_args(json=False))
    output = capsys.readouterr().out
    assert "[red]Person[/]" in output and "Direct reports: 1" in output
    assert "Department" in output and "Manager" in output


def test_partial_list_is_not_rendered_as_an_exact_total(monkeypatch, capsys):
    fake = fake_view(list_organisation_reports={"items": [], "complete": False, "error": "HTTP 403"})
    monkeypatch.setattr(runtime, "get_graph_client", lambda: fake)
    monkeypatch.setattr(runtime, "console", Console(width=100, force_terminal=False, color_system=None))
    contact_commands.cmd_contact_org(org_args(json=False))
    output = capsys.readouterr().out
    assert "0 returned; the list is incomplete" in output and "403" in output


@pytest.mark.parametrize("options", [{"user": "1"}, {"user": ".."}, {"user": "name\n"}, {"max_depth": 0}, {"max_depth": 2}])
def test_invalid_options_never_authenticate(monkeypatch, options):
    factory = Mock(side_effect=AssertionError("No authentication"))
    monkeypatch.setattr(contacts, "graph_client", factory)
    monkeypatch.setattr(runtime, "get_graph_client", factory)
    mcp_options = {"include_chain" if key == "chain" else key: value for key, value in options.items()}
    with pytest.raises(ValueError):
        contacts.contact_org(**mcp_options)
    with pytest.raises(CliError):
        contact_commands.cmd_contact_org(org_args(**options))
    factory.assert_not_called()


def test_cli_api_error_closes_client_and_is_explicit(monkeypatch):
    def forbidden(*args):
        raise OutlookAPIError(403, "Denied")
    fake = fake_view(get_organisation_profile=forbidden)
    monkeypatch.setattr(runtime, "get_graph_client", lambda: fake)
    with pytest.raises(CliError, match="403"):
        contact_commands.cmd_contact_org(org_args())
    assert fake.names().count("close") == 1


def test_parser_org_defaults_and_chain_options():
    parser = cli.build_parser()
    default = parser.parse_args(["contact", "org"])
    assert default.user == "me" and not default.chain and default.max_depth is None
    args = parser.parse_args(["contact", "org", "person@example.test", "--chain", "--max-depth", "3", "--json"])
    assert args.user == "person@example.test" and args.chain and args.max_depth == 3 and args.json
    assert args.func is contact_commands.cmd_contact_org


@pytest.mark.parametrize("restriction", ["readonly", "no_send"])
def test_mcp_org_remains_available_and_declares_read_only(restriction):
    with policy(**{restriction: True}):
        tools = {tool.name: tool for tool in asyncio.run(build_server().list_tools())}
    tool = tools["contact_org"]
    assert tool.annotations.readOnlyHint is True
    assert tool.annotations.destructiveHint is False and tool.annotations.idempotentHint is True
    assert tool.inputSchema["properties"]["user"]["default"] == "me"
    assert tool.inputSchema["properties"]["include_chain"]["default"] is False


def test_mcp_dispatch_reads_without_any_write(monkeypatch):
    fake = use_client(monkeypatch, contacts, "graph_client", fake_view())
    with policy(readonly=True):
        result = asyncio.run(build_server().call_tool("contact_org", {"user": "me"}))
    assert result
    assert set(fake.names()) == {"get_organisation_profile", "get_organisation_manager", "list_organisation_reports"}
