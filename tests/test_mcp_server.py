from __future__ import annotations

import asyncio
import json
import time
from pathlib import Path

import jwt
import pytest
from mcp.shared.memory import create_connected_server_and_client_session
from mcp_fakes import FakeClient, use_client

from outlook_draft import config
from outlook_draft.errors import OutlookAPIError
from outlook_draft.mcp import auth, calendar, clients, contacts, drafts, files, mail, tasks, teams
from outlook_draft.mcp.server import INSTRUCTIONS, MODULES, build_server
from outlook_draft.token_manager import TokenManager

EXPECTED_TOOLS = {
    "mail_search", "mail_unread", "mail_read", "mail_attachments", "mail_download_attachments",
    "mail_links", "mail_folders", "mail_mark_read", "mail_move", "mail_archive",
    "draft_create", "draft_reply", "draft_list", "draft_get", "draft_delete",
    "cal_agenda", "cal_get", "cal_availability", "cal_find_time", "cal_create", "cal_respond",
    "contact_search", "task_list", "task_create", "task_complete",
    "teams_list", "teams_search", "teams_messages", "teams_attachments", "teams_download_attachments",
    "files_sites", "files_libraries", "files_list", "files_search", "files_download",
    "auth_status", "mail_send", "draft_send", "teams_send",
}


def _tool_names() -> set[str]:
    return {tool.__name__ for module in MODULES for tool in module.TOOLS}


def test_full_tool_contract_matches_pre_refactor_snapshot():
    tools = asyncio.run(build_server().list_tools())
    expected = json.loads(Path(__file__).with_name("mcp_tool_contracts.json").read_text())
    assert [tool.model_dump(mode="json") for tool in tools if tool.name not in {"mail_send", "draft_send", "teams_send"}] == expected


def test_registers_exactly_the_phase_one_tools():
    assert _tool_names() == EXPECTED_TOOLS


def test_only_approved_send_and_destructive_tools_are_registered():
    forbidden = ("send", "delete_event", "cancel", "update", "upload", "mkdir", "rename", "contact_create")
    for name in _tool_names() - {"draft_delete", "mail_send", "draft_send", "teams_send"}:
        assert not any(word in name for word in forbidden), name
    tools = asyncio.run(build_server().list_tools())
    for tool in tools:
        if tool.name.endswith("_send"):
            assert tool.inputSchema["properties"]["confirm"]["default"] is False
            assert tool.annotations.idempotentHint is False
            assert tool.annotations.destructiveHint is True


def test_instructions_first_line_is_the_description():
    assert INSTRUCTIONS.splitlines()[0].endswith("Send tools preview unless confirm=true.")


def test_in_memory_session_lists_tools_and_calls_one(monkeypatch):
    use_client(monkeypatch, mail, "outlook_client", FakeClient(list_unread=[{"Id": "AAMk-1", "Subject": "Hi"}]))

    async def scenario():
        async with create_connected_server_and_client_session(build_server()) as session:
            listed = await session.list_tools()
            result = await session.call_tool("mail_unread", {"count": 3})
            return listed, result

    listed, result = asyncio.run(scenario())

    assert {t.name for t in listed.tools} == EXPECTED_TOOLS
    assert result.isError is False
    assert result.structuredContent["result"][0]["id"] == "AAMk-1"


def test_in_memory_session_reports_tool_errors(monkeypatch):
    def boom(*args, **kwargs):
        raise OutlookAPIError(500, "server down")

    use_client(monkeypatch, mail, "outlook_client", FakeClient(list_unread=boom))

    async def scenario():
        async with create_connected_server_and_client_session(build_server()) as session:
            return await session.call_tool("mail_unread", {})

    result = asyncio.run(scenario())

    assert result.isError is True
    assert "server down" in result.content[0].text


def _expired_tokens(tmp_path: Path) -> Path:
    token = jwt.encode({"exp": time.time() - 60}, "k" * 32, algorithm="HS256")
    path = tmp_path / "tokens.json"
    path.write_text(json.dumps({"tokens": {config.OUTLOOK_TOKEN_DOMAIN: token, config.GRAPH_TOKEN_DOMAIN: token}}))
    return path


@pytest.fixture
def never_reauth(monkeypatch):
    def fail(self, *args, **kwargs):
        raise AssertionError("run_reauth must not be called")

    monkeypatch.setattr(TokenManager, "run_reauth", fail)


@pytest.mark.parametrize(
    ("call", "label"),
    [
        (lambda: mail.mail_unread(), "Outlook API"),
        (lambda: teams.teams_list(), "Microsoft Graph"),
        (lambda: files.files_sites(), "Microsoft Graph"),
    ],
)
def test_expired_token_raises_auth_error_without_reauth(monkeypatch, tmp_path, never_reauth, call, label):
    monkeypatch.setattr(config, "TOKENS_FILE", _expired_tokens(tmp_path))

    with pytest.raises(clients.McpAuthError) as excinfo:
        call()

    assert str(excinfo.value) == f"{label} token expired or missing. Run 'outlook-cli auth' in a terminal, then retry."


def test_missing_token_file_raises_auth_error(monkeypatch, tmp_path, never_reauth):
    monkeypatch.setattr(config, "TOKENS_FILE", tmp_path / "absent.json")

    with pytest.raises(clients.McpAuthError, match="Run 'outlook-cli auth'"):
        mail.mail_unread()


def test_token_manager_never_reauths_even_when_asked(monkeypatch, tmp_path, never_reauth):
    manager = clients.NoReauthTokenManager(tokens_file=_expired_tokens(tmp_path))

    with pytest.raises(Exception, match="expired"):
        manager.get_token(auto_reauth=True)
    assert manager.run_reauth() is False


def test_token_read_per_call_picks_up_refreshed_tokens(monkeypatch, tmp_path):
    tokens = tmp_path / "tokens.json"
    monkeypatch.setattr(config, "TOKENS_FILE", tokens)
    tokens.write_text(json.dumps({"tokens": {}}))
    with pytest.raises(clients.McpAuthError):
        mail.mail_unread()

    fresh = jwt.encode({"exp": time.time() + 3600}, "k" * 32, algorithm="HS256")
    tokens.write_text(json.dumps({"tokens": {config.OUTLOOK_TOKEN_DOMAIN: fresh}}))
    seen = {}
    monkeypatch.setattr(
        clients.OutlookClient, "list_unread", lambda self, top=20: seen.setdefault("token", self._tm.get_token()) and []
    )

    assert mail.mail_unread() == []
    assert seen["token"] == fresh


def test_tool_modules_never_call_sys_exit():
    import inspect

    for module in (*MODULES, clients):
        assert "sys.exit" not in inspect.getsource(module), module.__name__


@pytest.mark.parametrize(
    ("module", "attr", "returns", "call"),
    [
        (mail, "outlook_client", {"list_unread": [{"Id": "1"}]}, lambda: mail.mail_unread()),
        (drafts, "outlook_client", {"list_drafts": [{"Id": "1"}]}, lambda: drafts.draft_list()),
        (calendar, "outlook_client", {"get_agenda": [{"Id": "1"}]}, lambda: calendar.cal_agenda()),
        (contacts, "outlook_client", {"search_people": [{"Id": "1"}]}, lambda: contacts.contact_search("a")),
        (tasks, "outlook_client", {"list_tasks": [{"Id": "1"}]}, lambda: tasks.task_list()),
        (teams, "graph_client", {"list_teams_chats": [{"id": "c", "topic": "t"}]}, lambda: teams.teams_list()),
        (files, "files_client", {"sp_list_sites": [{"id": "g", "displayName": "S"}]}, lambda: files.files_sites()),
    ],
)
def test_tools_write_nothing_to_stdout(monkeypatch, capsys, module, attr, returns, call):
    use_client(monkeypatch, module, attr, FakeClient(**returns))

    call()

    captured = capsys.readouterr()
    assert captured.out == ""


def test_auth_status_writes_nothing_to_stdout(monkeypatch, tmp_path, capsys):
    monkeypatch.setattr(config, "TOKENS_FILE", tmp_path / "none.json")

    auth.auth_status()

    assert capsys.readouterr().out == ""


def test_stdio_server_emits_only_protocol_frames_on_stdout():
    import subprocess
    import sys

    requests = [
        {"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {
            "protocolVersion": "2025-06-18", "capabilities": {}, "clientInfo": {"name": "t", "version": "0"}}},
        {"jsonrpc": "2.0", "method": "notifications/initialized"},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
    ]
    proc = subprocess.Popen(
        [sys.executable, "-m", "outlook_draft", "mcp"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.DEVNULL,
        text=True,
    )
    try:
        for request in requests:
            proc.stdin.write(json.dumps(request) + "\n")
            proc.stdin.flush()
        # Every line on stdout must be a JSON-RPC frame; json.loads fails otherwise.
        frames = [json.loads(proc.stdout.readline()) for _ in range(2)]
    finally:
        proc.stdin.close()
        proc.wait(timeout=30)

    assert [f["id"] for f in frames] == [1, 2]
    assert all(f["jsonrpc"] == "2.0" for f in frames)
    assert len(frames[1]["result"]["tools"]) == len(EXPECTED_TOOLS)
    assert proc.stdout.read().strip() == ""
