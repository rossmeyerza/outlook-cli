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
from outlook_draft.commands import account, calendar, drafts, files, tasks, teams
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
    handler = parser.get_default("func")
    assert handler is None or callable(handler)
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
    monkeypatch.setattr(drafts, "console", console)
    monkeypatch.setattr(files, "_GraphClient", lambda: client)
    monkeypatch.setattr(account, "refresh_account_paths", lambda: None)
    monkeypatch.setattr(files, "console", console)
    for module in (calendar, tasks, teams):
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


@pytest.mark.parametrize("exit_code", [1, 2, 7])
def test_central_cli_error_exit_code(monkeypatch, capsys, isolated_cli, exit_code):
    from outlook_draft.errors import CliError

    isolated_cli.list_unread.side_effect = CliError("Expected failure", exit_code)
    result = run_command(monkeypatch, capsys, ["mail", "unread"])
    assert result == {"stdout": "Expected failure\n", "stderr": "", "exit": exit_code}


def test_api_error_renders_red(monkeypatch, capsys, isolated_cli):
    console = Console(width=100, force_terminal=True, color_system="standard")
    monkeypatch.setattr(cli, "console", console)
    isolated_cli.list_unread.side_effect = OutlookAPIError(503, "Unavailable")
    result = run_command(monkeypatch, capsys, ["mail", "unread"])
    assert result["exit"] == 1
    assert "\x1b[31m" in result["stdout"]


@pytest.mark.parametrize("argv", [
    ["teams", "list", "-n", "12"],
    ["teams", "list", "-n", "12", "--sort-received"],
    ["teams", "search", "Match", "--scan", "12", "-n", "3"],
])
def test_cli_parallel_teams_matches_sequential(monkeypatch, capsys, isolated_cli, argv):
    from contextlib import contextmanager

    isolated_cli.list_teams_chats.side_effect = lambda **kwargs: [
        {"id": str(index), "chatType": "group", "lastUpdatedDateTime": str(index)}
        for index in range(12)
    ]
    isolated_cli.get_current_user.return_value = {"id": "self"}
    isolated_cli.list_teams_message_metadata.return_value = []

    def members(chat_id, **kwargs):
        if chat_id == "8":
            raise OutlookAPIError(403, "Denied")
        return [{"displayName": "Match " + chat_id, "email": "other@example.test"}]

    @contextmanager
    def sequential(function, items):
        yield map(function, items)

    isolated_cli.list_teams_chat_members.side_effect = members
    actual = run_command(monkeypatch, capsys, argv)
    monkeypatch.setattr(teams, "ordered_map", sequential)
    assert actual == run_command(monkeypatch, capsys, argv)
