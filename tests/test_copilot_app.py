from __future__ import annotations

import argparse
import asyncio
import json
from contextlib import contextmanager
from types import SimpleNamespace

import pytest

from outlook_draft.commands import copilot_app as command
from outlook_draft.copilot_bridge import app
from outlook_draft.errors import CliError, CopilotError
from outlook_draft.mcp import copilot
from outlook_draft.mcp.server import build_server


@pytest.mark.parametrize(("status", "exit_code"), [("completed", 0), ("approval_required", 3), ("timeout", 1), ("error", 1)])
def test_cli_json_and_exit_codes(monkeypatch, capsys, status, exit_code):
    monkeypatch.setattr(command, "_ask", lambda args: {"status": status, "text": "Hello"})
    args = argparse.Namespace(json=True)
    if exit_code:
        with pytest.raises(CliError) as error:
            command.cmd_ask(args)
        assert error.value.exit_code == exit_code
        assert str(error.value) == ""  # Keep stdout as one JSON document.
    else:
        command.cmd_ask(args)
    assert json.loads(capsys.readouterr().out) == {"status": status, "text": "Hello"}


def test_cli_bridge_error_as_json(monkeypatch, capsys):
    def fail(args):
        raise CopilotError("No Chrome")
    monkeypatch.setattr(command, "_ask", fail)
    with pytest.raises(CliError):
        command.cmd_ask(argparse.Namespace(json=True))
    result = json.loads(capsys.readouterr().out)
    assert result["status"] == "error"
    assert result["applied"] is None


def test_mcp_app_uses_bridge_and_has_destructive_annotations(monkeypatch):
    @contextmanager
    def client():
        yield object()
    monkeypatch.setattr(copilot, "files_client", client)
    monkeypatch.setattr(copilot, "resolve_app_file", lambda *a, **kw: {"app": "excel", "shared": True})
    calls = []
    def ask(*args, **kwargs):
        calls.append((args, kwargs))
        return {"status": "completed", "applied": True}
    monkeypatch.setattr(copilot, "ask_app", ask)
    result = copilot.copilot_app("Add a table", "Documents/Test.xlsx", open_tab=True)
    assert result["status"] == "completed"
    assert "outside" in result["warning"]
    assert calls[0][1]["open_tab"] is True
    tool = next(t for t in asyncio.run(build_server().list_tools()) if t.name == "copilot_app")
    assert tool.annotations.readOnlyHint is False
    assert tool.annotations.destructiveHint is True
    assert tool.annotations.idempotentHint is False


def test_file_resolution_uses_unique_id_and_drive_ownership(monkeypatch):
    resolved = SimpleNamespace(name="test.pptx", item_id="item", drive_id="drive", web_url="https://t.sharepoint.com/file")
    monkeypatch.setattr(app, "resolve_copilot_file", lambda *a, **kw: resolved)
    class Graph:
        def drive_item_with_sharepoint_ids(self, *args):
            return {"sharepointIds": {"listItemUniqueId": "file-guid"}}
        def _request(self, *args, **kwargs):
            return SimpleNamespace(json=lambda: {"id": "drive"})
    file = app.resolve_app_file(Graph(), "test")
    assert file["app"] == "powerpoint"
    assert file["uniqueId"] == "file-guid"
    assert file["shared"] is False


def test_excel_script_source_is_not_proof_of_an_applied_edit():
    event = {"type": "action", "app": "excel", "name": "executeScript", "writes": True}
    result = {"status": "completed", "applied": True, "actions": [event, {"type": "action_result", "ok": True}]}
    app._normalise_result(result, {"app": "excel"})
    assert result["applied"] is None
    assert result["actions"][0]["mayWrite"] is True
    assert "writes" not in result["actions"][0]
    assert event["writes"] is True  # Do not mutate the source event.
    assert "Not verified" in result["appliedEvidence"]


def test_wait_saved_false_when_no_later_modification(monkeypatch):
    from datetime import UTC, datetime
    monkeypatch.setattr(app, "_modified_at", lambda *args: None)
    assert app.wait_for_save(None, {}, datetime.now(UTC), 0) is False


def test_status_text_shows_busy_age_and_uncertain_state():
    assert command._busy_text({"busy": False, "uncertain": False}) == "busy: False"
    assert command._busy_text({"busy": True, "busyAgeSeconds": 42, "phase": "wait_answer"}) == "busy: True (wait_answer, 42 s)"
    text = command._busy_text({"busy": False, "uncertain": True, "uncertainReason": "timeout", "uncertainAgeSeconds": 300})
    assert text == "busy: False | uncertain: True (timeout, 300 s ago)"


def test_cli_warns_when_a_failed_request_was_sent(monkeypatch, capsys):
    monkeypatch.setattr(command, "_ask", lambda args: {"status": "timeout", "sent": True, "phase": "wait_answer", "text": ""})
    with pytest.raises(CliError):
        command.cmd_ask(argparse.Namespace(json=False))
    err = capsys.readouterr().err
    assert "Last phase: wait_answer." in err
    assert "can already have changed the document" in err


def test_cli_reset(monkeypatch, capsys):
    class Client:
        def reset(self, tab_id):
            return {"tabId": tab_id, "cleared": {"request": False, "uncertain": True}}
    monkeypatch.setattr(command, "BridgeClient", Client)
    command.cmd_reset(argparse.Namespace(tab_id=7, json=False))
    out = capsys.readouterr()
    assert "cleared uncertain state" in out.out
    assert "Check it" in out.err
