from __future__ import annotations

from datetime import timezone

from outlook_draft.commands.gateway import (
    _chunk_response,
    _handle_gateway_command,
    _parse_graph_time,
)


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
