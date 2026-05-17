from __future__ import annotations

from outlook_draft.commands.teams import _is_received_user_message


def test_received_message_filter_ignores_system_deleted_and_self_messages() -> None:
    self_id = "self"

    assert not _is_received_user_message({"messageType": "systemEventMessage"}, self_id)
    assert not _is_received_user_message({"messageType": "message", "deletedDateTime": "x"}, self_id)
    assert not _is_received_user_message(
        {"messageType": "message", "from": {"user": {"id": "self"}}},
        self_id,
    )
    assert _is_received_user_message(
        {"messageType": "message", "from": {"user": {"id": "other"}}},
        self_id,
    )
