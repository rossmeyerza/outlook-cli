from __future__ import annotations

from datetime import datetime, timedelta, timezone

from outlook_draft.commands.gateway import (
    _chunk_response,
    _gateway_command_help,
    _heartbeat_is_stale,
    _handle_gateway_command,
    _list_pi_models,
    _model_help,
    _parse_graph_time,
)
from outlook_draft.commands.pi_session import _summarize_progress_event


def test_gateway_command_parser_ignores_normal_prompts() -> None:
    assert _handle_gateway_command("chat", "what is on today?", "@Marlow", 30) is None


def test_gateway_help_command_lists_session_controls() -> None:
    response = _handle_gateway_command("chat", "!help", "@Marlow", 30)

    assert response is not None
    assert "!new" in response
    assert "!reset" in response
    assert "!status" in response
    assert "!pause" in response


def test_gateway_help_supports_command_topics() -> None:
    assert "!model commands" in _gateway_command_help("model")
    assert "!model list" in _model_help()


def test_graph_time_parser_normalizes_zulu_timestamps() -> None:
    parsed = _parse_graph_time("2026-06-01T12:34:56Z")

    assert parsed.tzinfo == timezone.utc
    assert parsed.isoformat() == "2026-06-01T12:34:56+00:00"


def test_heartbeat_stale_only_when_running_and_old() -> None:
    old = (datetime.now(timezone.utc) - timedelta(minutes=10)).isoformat()

    assert _heartbeat_is_stale({"last_poll_at": old, "poll_interval": 30}, running=True)
    assert not _heartbeat_is_stale({"last_poll_at": old, "poll_interval": 30}, running=False)


def test_chunk_response_splits_long_messages_on_boundaries() -> None:
    response = "alpha beta gamma delta"

    assert _chunk_response(response, max_chars=12) == ["alpha beta", "gamma delta"]


def test_summarize_progress_event_for_tool_call_command() -> None:
    assert _summarize_progress_event(
        {"type": "tool_call", "name": "bash", "arguments": {"command": "outlook-cli cal agenda"}}
    ) == "Using bash: outlook-cli cal agenda"


def test_summarize_progress_event_for_tool_result() -> None:
    assert _summarize_progress_event({"type": "tool_result", "toolName": "bash"}) == "bash finished"


def test_summarize_progress_event_for_message_tool_call() -> None:
    assert _summarize_progress_event(
        {
            "type": "message",
            "message": {
                "content": [
                    {"type": "toolCall", "name": "bash", "arguments": {"command": "outlook-cli mail unread"}}
                ]
            },
        }
    ) == "Using bash: outlook-cli mail unread"


def test_list_pi_models_uses_optional_search(monkeypatch) -> None:
    calls = []

    class Result:
        returncode = 0
        stdout = "wpp/claude-sonnet-4.5\nwpp/claude-opus-4.8"
        stderr = ""

    monkeypatch.setattr("outlook_draft.commands.gateway.shutil.which", lambda name: "/usr/bin/pi")

    def fake_run(cmd, **kwargs):
        calls.append((cmd, kwargs))
        return Result()

    monkeypatch.setattr("outlook_draft.commands.gateway.subprocess.run", fake_run)

    response = _list_pi_models("sonnet")

    assert calls[0][0] == ["/usr/bin/pi", "--list-models", "sonnet"]
    assert "Pi models matching 'sonnet':" in response
    assert "wpp/claude-sonnet-4.5" in response
