"""Synthetic checks for request policies and side-effect-free CLI plans."""
from __future__ import annotations

import argparse
import json
import sys
from unittest.mock import Mock

import pytest

from outlook_draft import cli
from outlook_draft.errors import OutlookDraftError
from outlook_draft.safety import SafetyError, check_request, command_allowed, policy


@pytest.fixture(autouse=True)
def clean_safety_env(monkeypatch):
    for name in ("OUTLOOK_CLI_READONLY", "OUTLOOK_CLI_NO_SEND"):
        monkeypatch.delenv(name, raising=False)


@pytest.mark.parametrize("method,path", [
    ("GET", "/messages"), ("HEAD", "/messages"), ("OPTIONS", "/messages"),
    ("POST", "/me/calendar/getSchedule"), ("POST", "/me/findMeetingTimes?x=1"),
])
def test_readonly_permits_reads(method, path):
    with policy(readonly=True):
        check_request(method, "https://example.test" + path)


@pytest.mark.parametrize("method,path", [
    ("POST", "/me/sendMail"), ("PATCH", "/messages/123"),
    ("DELETE", "/messages/123"), ("PUT", "/files/123/content"),
    ("POST", "/me/getSchedule/other"), ("PATCH", "/me/findMeetingTimes"),
])
def test_readonly_blocks_mutations(method, path):
    with policy(readonly=True), pytest.raises(OutlookDraftError):
        check_request(method, "https://example.test" + path)


@pytest.mark.parametrize("path", [
    "/me/sendMail", "/messages/123/send", "/chats/123/messages",
    "/teams/123/channels/456/messages", "/me/events", "/me/events/123/accept",
])
def test_no_send_blocks_message_and_invitation_requests(path):
    with policy(no_send=True), pytest.raises(SafetyError):
        check_request("POST", "https://example.test" + path)


def test_no_send_allows_drafts_and_reads():
    with policy(no_send=True):
        check_request("POST", "https://example.test/me/messages")
        check_request("GET", "https://example.test/chats/123/messages")


@pytest.mark.parametrize("name,path", [
    ("OUTLOOK_CLI_READONLY", "/messages"), ("OUTLOOK_CLI_NO_SEND", "/me/sendMail"),
])
def test_environment_restrictions_cannot_be_disabled(monkeypatch, name, path):
    monkeypatch.setenv(name, "1")
    with policy(readonly=False, no_send=False), pytest.raises(SafetyError):
        check_request("POST", "https://example.test" + path)


@pytest.mark.parametrize("fail", [False, True])
def test_policy_resets_and_nested_context_cannot_weaken_it(fail):
    try:
        with policy(readonly=True):
            with policy(), pytest.raises(SafetyError):
                check_request("POST", "/messages")
            if fail:
                raise ValueError("expected")
    except ValueError:
        pass
    check_request("POST", "/messages")


@pytest.mark.parametrize("name,enabled,disabled,expected", [
    ("mail.rules.create", "mail rules", None, True),
    ("mail.rules.create", "mail", "mail.rules", False),
    ("mail.rules.create", "mail.rule", None, False),
    ("mail.read", "mail,teams", "mail.send", True),
    ("teams.send", None, "teams", False),
    ("mail_send", ["mail_send"], None, True),
])
def test_command_prefixes(name, enabled, disabled, expected):
    assert command_allowed(name, enabled, disabled) is expected


def test_invalid_prefix_fails_closed():
    with pytest.raises(SafetyError):
        command_allowed("mail.read", "mail,,teams")


@pytest.mark.parametrize("position", ["before", "after"])
def test_dry_run_skips_dispatch_and_omits_private_content(monkeypatch, capsys, position):
    forbidden = Mock(side_effect=AssertionError("No dispatch or account access is permitted"))
    monkeypatch.setattr(cli, "_dispatch", forbidden)
    argv = ["mail", "send", "--to", "private@example.test", "-s", "Secret subject",
            "-b", "Secret body", "--yes"]
    flags = ["--dry-run", "--readonly", "--no-send", "--account", "missing"]
    monkeypatch.setattr(sys, "argv", ["outlook-cli", "--json", *(flags + argv if position == "before" else argv + flags)])
    cli.main()
    output = capsys.readouterr().out
    assert json.loads(output) == {"kind": "plan", "command": "mail.send", "targets": {}}
    forbidden.assert_not_called()


@pytest.mark.parametrize("flags", [["--disable-commands", "mail"], ["--enable-commands", "teams"]])
def test_command_policy_precedes_dispatch_and_dry_run(monkeypatch, capsys, flags):
    forbidden = Mock(side_effect=AssertionError("No dispatch is permitted"))
    monkeypatch.setattr(cli, "_dispatch", forbidden)
    monkeypatch.setattr(sys, "argv", ["outlook-cli", "mail", "read", "target", "--dry-run", *flags])
    with pytest.raises(SystemExit) as error:
        cli.main()
    assert error.value.code == 1
    assert "Command policy blocks" in capsys.readouterr().out
    forbidden.assert_not_called()


def test_nested_selected_command_uses_parser_tree():
    parser = argparse.ArgumentParser()
    mail = parser.add_subparsers(dest="domain").add_parser("mail")
    rules = mail.add_subparsers(dest="command").add_parser("rules")
    rules.add_subparsers(dest="rule_action").add_parser("create")
    args = parser.parse_args(["mail", "rules", "create"])
    assert cli._selected_command(parser, args) == "mail.rules.create"


@pytest.mark.parametrize("fail", [False, True])
def test_cli_request_policy_resets_after_dispatch(monkeypatch, fail):
    def dispatch(args, parser):
        with pytest.raises(SafetyError):
            check_request("POST", "/me/sendMail")
        if fail:
            raise OutlookDraftError("expected")

    monkeypatch.setattr(cli, "_dispatch", dispatch)
    monkeypatch.setattr(sys, "argv", ["outlook-cli", "--readonly", "draft", "send", "target", "--yes"])
    if fail:
        with pytest.raises(SystemExit):
            cli.main()
    else:
        cli.main()
    check_request("POST", "/me/sendMail")


@pytest.mark.parametrize("restriction", ["--readonly", "--no-send", "OUTLOOK_CLI_READONLY", "OUTLOOK_CLI_NO_SEND"])
@pytest.mark.parametrize("command", [
    ["copilot", "ask", "Secret prompt"],
    ["copilot", "app", "ask", "Secret prompt", "--file", "target"],
])
def test_copilot_prompts_cannot_bypass_safety(monkeypatch, capsys, restriction, command):
    forbidden = Mock(side_effect=AssertionError("Copilot must not start"))
    monkeypatch.setattr(cli, "_dispatch", forbidden)
    flags = []
    if restriction.startswith("--"):
        flags.append(restriction)
    else:
        monkeypatch.setenv(restriction, "1")
    monkeypatch.setattr(sys, "argv", ["outlook-cli", *command, *flags])
    with pytest.raises(SystemExit) as error:
        cli.main()
    assert error.value.code == 1
    assert "blocks Copilot prompts" in capsys.readouterr().out
    forbidden.assert_not_called()
    check_request("GET", "/me/messages")


@pytest.mark.parametrize("command", [["status"], ["inspect", "--tab-id", "1"]])
def test_restricted_copilot_diagnostics_remain_available(monkeypatch, command):
    dispatch = Mock()
    monkeypatch.setattr(cli, "_dispatch", dispatch)
    monkeypatch.setenv("OUTLOOK_CLI_READONLY", "1")
    monkeypatch.setenv("OUTLOOK_CLI_NO_SEND", "1")
    monkeypatch.setattr(sys, "argv", ["outlook-cli", "--readonly", "--no-send", "copilot", "app", *command])
    cli.main()
    dispatch.assert_called_once()
