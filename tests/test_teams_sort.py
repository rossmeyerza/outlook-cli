from __future__ import annotations

from outlook_draft.commands.teams import (
    SELF_CHAT_ID,
    _chat_match_reasons,
    _chat_title,
    _is_received_user_message,
    find_self_chat,
)
from outlook_draft.errors import OutlookAPIError


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

    assert _chat_title({"id": SELF_CHAT_ID}) == "Self chat"
    assert _chat_title(
        {"chatType": "oneOnOne", "id": "19:self@unq.gbl.spaces"},
        [{"displayName": "Ross Meyer", "email": "ross@example.com"}],
    ) == "Self chat"


def test_chat_match_reasons_matches_topic() -> None:
    reasons = _chat_match_reasons(
        {"topic": "Q2 budget review"},
        [],
        "budget",
    )

    assert reasons == ["topic"]


def test_chat_match_reasons_matches_member_name_and_email_terms() -> None:
    reasons = _chat_match_reasons(
        {"topic": ""},
        [
            {
                "displayName": "Tarik Patel",
                "email": "tarik.patel@example.com",
                "userId": "user-1",
            },
            {
                "displayName": "Ada Lovelace",
                "email": "ada@example.com",
                "userId": "user-2",
            },
        ],
        "tarik example",
    )

    assert reasons == ["member: Tarik Patel"]


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

        def list_teams_message_metadata(self, chat_id, top=10):
            raise OutlookAPIError(404, "not found")

    found = find_self_chat(Client())

    assert found is not None
    assert found[0]["id"] == "self-chat"


def test_find_self_chat_prefers_special_notes_thread() -> None:
    class Client:
        def get_current_user(self):
            return {"id": "self", "mail": "ross@example.com", "displayName": "Ross Meyer"}

        def list_teams_message_metadata(self, chat_id, top=10):
            assert chat_id == SELF_CHAT_ID
            return [{"createdDateTime": "2026-06-01T19:17:16.87Z"}]

    found = find_self_chat(Client())

    assert found is not None
    assert found[0]["id"] == SELF_CHAT_ID
    assert found[0]["lastUpdatedDateTime"] == "2026-06-01T19:17:16.87Z"
