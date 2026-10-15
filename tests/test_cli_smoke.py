"""Recorded CLI contracts. All clients and account paths are isolated."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from unittest.mock import Mock

import pytest
from rich.console import Console

from outlook_draft import cli, runtime
from outlook_draft.commands import files
from outlook_draft.errors import OutlookAPIError

SNAPSHOT = Path(__file__).with_name("cli_smoke_outputs.json")
COMMANDS = [
    ["mail", "unread", "--json"],
    ["mail", "search", "sample", "--json"],
    ["cal", "agenda", "--json"],
    ["task", "list", "--json"],
    ["--json", "teams", "list"],
    ["files", "sites", "--json"],
    ["draft", "send", "sample"],
    ["mail", "send", "--to", "sample@example.test", "-s", "Sample", "-b", "Body"],
    ["teams", "send", "sample", "Body"],
]


def capture_parser(monkeypatch):
    parsers = []

    def capture(self, *args, **kwargs):
        parsers.append(self)
        raise SystemExit(0)

    with monkeypatch.context() as patch:
        patch.setattr(argparse.ArgumentParser, "parse_args", capture)
        with pytest.raises(SystemExit):
            cli.main()
    return parsers[0]


def help_outputs(parser, prefix=()):
    outputs = {" ".join(prefix): parser.format_help()}
    for action in parser._actions:
        if isinstance(action, argparse._SubParsersAction):
            for name, child in action.choices.items():
                outputs.update(help_outputs(child, (*prefix, name)))
    return outputs


@pytest.fixture
def isolated_cli(monkeypatch, tmp_path):
    client = Mock()
    client.list_unread.return_value = [{"Id": "message", "Subject": "Sample"}]
    client.search_messages.return_value = client.list_unread.return_value
    client.get_agenda.return_value = [{"Id": "event", "Subject": "Sample"}]
    client.list_tasks.return_value = [{"Id": "task", "Subject": "Sample"}]
    client.list_teams_chats.return_value = [{"id": "chat", "topic": "Sample", "chatType": "group"}]
    client.sp_list_sites.return_value = [{"id": "site", "displayName": "Sample"}]
    console = Console(width=100, force_terminal=False, color_system=None)
    monkeypatch.setattr(cli, "console", console)
    monkeypatch.setattr(runtime, "console", console)
    monkeypatch.setattr(runtime, "get_client", lambda **kwargs: client)
    monkeypatch.setattr(runtime, "get_graph_client", lambda: client)
    monkeypatch.setattr(cli, "_get_client", lambda **kwargs: client)
    monkeypatch.setattr(cli, "_get_graph_client", lambda: client)
    monkeypatch.setattr(files, "_GraphClient", lambda: client)
    monkeypatch.setattr(cli, "_refresh_account_paths", lambda: None)
    monkeypatch.setattr(files, "console", console)
    for module in (cli.calendar_commands, cli.tasks_commands, cli.teams_commands):
        for name in ("CAL_CACHE", "TASK_CACHE", "TEAMS_CACHE"):
            if hasattr(module, name):
                monkeypatch.setattr(module, name, tmp_path / name)
    return client


def run_command(monkeypatch, capsys, argv):
    monkeypatch.setattr(sys, "argv", ["outlook-cli", "--no-spinner", *argv])
    code = 0
    try:
        cli.main()
    except SystemExit as error:
        code = error.code
    captured = capsys.readouterr()
    return {"stdout": captured.out, "stderr": captured.err, "exit": code}


def test_all_help_matches_recording(monkeypatch):
    parser = capture_parser(monkeypatch)
    expected = json.loads(SNAPSHOT.read_text())["help"]
    assert help_outputs(parser) == expected
    for path in expected:
        with pytest.raises(SystemExit) as error:
            parser.parse_args([*path.split(), "--help"])
        assert error.value.code == 0


@pytest.mark.parametrize("argv", COMMANDS)
def test_outputs_match_recording(monkeypatch, capsys, isolated_cli, argv):
    expected = json.loads(SNAPSHOT.read_text())["commands"][" ".join(argv)]
    assert run_command(monkeypatch, capsys, argv) == expected
    isolated_cli.send_mail.assert_not_called()
    isolated_cli.send_message.assert_not_called()
    isolated_cli.send_teams_message.assert_not_called()


def test_api_error_matches_recording(monkeypatch, capsys, isolated_cli):
    isolated_cli.list_unread.side_effect = OutlookAPIError(503, "Unavailable")
    actual = run_command(monkeypatch, capsys, ["mail", "unread"])
    assert actual == json.loads(SNAPSHOT.read_text())["error"]
