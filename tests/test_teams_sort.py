from __future__ import annotations

from outlook_draft.commands.teams import _chat_title, _is_received_user_message, find_self_chat


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


def test_chat_title_labels_self_chat(monkeypatch) -> None:
    monkeypatch.setattr("outlook_draft.commands.teams.config.MS_EMAIL", "ross@example.com")

    assert _chat_title(
        {"chatType": "oneOnOne", "id": "19:self@unq.gbl.spaces"},
        [{"displayName": "Ross Meyer", "email": "ross@example.com"}],
    ) == "Self chat"


def test_find_self_chat_detects_single_current_user_member() -> None:
    class Client:
        def get_current_user(self):
            return {"id": "self", "mail": "ross@example.com"}

        def list_teams_chats(self, top=500):
            return [
                {"id": "other", "chatType": "oneOnOne", "lastUpdatedDateTime": "2026-06-01T09:00:00Z"},
                {"id": "self-chat", "chatType": "oneOnOne", "lastUpdatedDateTime": "2026-06-01T10:00:00Z"},
            ]

        def list_teams_chat_members(self, chat_id, top=10):
            if chat_id == "other":
                return [{"userId": "other", "email": "other@example.com"}]
            return [{"userId": "self", "email": "ross@example.com", "displayName": "Ross Meyer"}]

    found = find_self_chat(Client())

    assert found is not None
    assert found[0]["id"] == "self-chat"
