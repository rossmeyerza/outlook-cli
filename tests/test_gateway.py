from __future__ import annotations

from datetime import timezone

from outlook_draft.commands.gateway import (
    _chunk_response,
    _handle_gateway_command,
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


def test_graph_time_parser_normalizes_zulu_timestamps() -> None:
    parsed = _parse_graph_time("2026-06-01T12:34:56Z")

    assert parsed.tzinfo == timezone.utc
    assert parsed.isoformat() == "2026-06-01T12:34:56+00:00"


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
