from __future__ import annotations

import base64
import json
import time
from argparse import Namespace
from urllib.parse import parse_qs, urlsplit

import jwt
import pytest
from mcp_fakes import FakeClient, use_client
from websockets.datastructures import Headers
from websockets.exceptions import InvalidStatus
from websockets.http11 import Response

from outlook_draft import cache, config
from outlook_draft.commands import copilot as copilot_commands
from outlook_draft.copilot import (
    RECORD_SEPARATOR,
    AnswerCollector,
    CopilotClient,
    build_file_annotation,
    build_stream_url,
    clean_answer_text,
    copilot_file_from_drive_item,
    parse_frames,
)
from outlook_draft.errors import CliError, CopilotError
from outlook_draft.mcp import copilot as copilot_tools

OID = "11111111-1111-1111-1111-111111111111"
TID = "22222222-2222-2222-2222-222222222222"

DRIVE_ITEM = {
    "id": "01ABCDEFGHIJKLMNOPQRSTUVWXYZ",
    "name": "Plan.pptx",
    "webUrl": "https://contoso.sharepoint.com/sites/team/Shared%20Documents/Plan.pptx",
    "parentReference": {"driveId": "b!drive"},
    "sharepointIds": {"siteId": "site-guid", "webId": "web-guid", "listId": "list-guid"},
}


def _token() -> str:
    claims = {"aud": "https://substrate.office.com/sydney", "oid": OID, "tid": TID, "exp": int(time.time()) + 3600}
    return jwt.encode(claims, "x" * 32, algorithm="HS256")


def _part(*messages: dict, **argument) -> dict:
    return {"type": 2, "invocationId": "1", "item": {
        "type": "message_part", "target": "update", "argument": {"messages": list(messages), **argument},
    }}


def _turn_frames(request_id: str) -> list[dict]:
    """A file answer as StreamHub sends it: echo, progress, streamed text, citations, end."""
    citations = {"searchResultsMap": {
        "1": {"id": "a", "sourceName": "Plan", "sourceUrl": "https://contoso.sharepoint.com/Plan.pptx"},
        "2": {"id": "b", "sourceUrl": "Page_1_thumbnail"},
    }}
    return [
        _part({"author": "user", "text": "Summarise", "requestId": request_id}),
        _part({"author": "bot", "messageType": "Progress", "text": "Digging in…", "requestId": request_id}),
        _part({"author": "bot", "messageType": "ChatName", "text": "Plan summary"}),
        _part({"author": "bot", "contentOrigin": "DeepLeo", "text": "The plan has **two** goals \ue200cite\ue202turn1sea"}),
        _part({"author": "bot", "contentOrigin": "DeepLeo",
               "text": "The plan has **two** goals \ue200cite\ue202turn1search1\ue201 for <Person>Contoso</Pers"}),
        _part({"author": "bot", "contentOrigin": "DeepLeo", "requestId": request_id,
               "text": "The plan has **two** goals \ue200cite\ue202turn1search1\ue201 for <Person>Contoso</Person>.  \n"}),
        _part({"author": "bot", "messageType": "Internal", "contentOrigin": "citation-formatter",
               "requestId": request_id, "text": json.dumps(citations)}),
        _part({"author": None, "messageType": "EndOfRequest"}),
    ]


class FakeSocket:
    def __init__(self, replies: list[str]):
        self.replies = list(replies)
        self.sent: list[dict] = []

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def send(self, text: str) -> None:
        assert text.endswith(RECORD_SEPARATOR)
        self.sent.append(json.loads(text[:-1]))
        request = self.chat_request()
        if request and not getattr(self, "_answered", False):
            self._answered = True
            frames = _turn_frames(request["requestId"])
            self.replies.extend(json.dumps(f) + RECORD_SEPARATOR for f in frames)

    def recv(self, timeout: float | None = None) -> str:
        if not self.replies:
            raise TimeoutError
        return self.replies.pop(0)

    def chat_request(self) -> dict | None:
        for frame in self.sent:
            item = frame.get("item") or {}
            if item.get("type") == "chat_request":
                return item["chatStreamRequest"]
        return None


class FakeTokenManager:
    def __init__(self, reauth: bool = True):
        self.reauth_calls = 0
        self._reauth = reauth

    def get_token(self, *, auto_reauth: bool = False, headless: bool = True) -> str:
        return _token()

    def run_reauth(self, *, headless: bool = True) -> bool:
        self.reauth_calls += 1
        return self._reauth


def test_file_annotation_uses_substrate_spo_id():
    file = copilot_file_from_drive_item(DRIVE_ITEM)
    triple = base64.urlsafe_b64encode(b"site-guid,web-guid,list-guid").decode().rstrip("=")

    annotation = build_file_annotation(file)

    assert file.spo_id == f"SPO_{triple}_01ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    assert annotation["messageAnnotationType"] == "File"
    assert annotation["id"] == annotation["messageAnnotationMetadata"]["id"] == file.spo_id
    assert annotation["messageAnnotationMetadata"]["fileType"] == "pptx"
    assert annotation["messageAnnotationMetadata"]["fileInfo"] == {"fileUrl": DRIVE_ITEM["webUrl"], "fileStoreType": "ODSP"}


def test_drive_item_without_sharepoint_ids_or_as_folder_is_rejected():
    with pytest.raises(CopilotError, match="siteId"):
        copilot_file_from_drive_item({**DRIVE_ITEM, "sharepointIds": {}})
    with pytest.raises(CopilotError, match="folder"):
        copilot_file_from_drive_item({**DRIVE_ITEM, "folder": {}})


def test_stream_url_routes_to_the_token_user():
    token = _token()
    url = build_stream_url(token, session_id="s-1", client_session_id="c-1")
    parts = urlsplit(url)
    query = parse_qs(parts.query)

    assert parts.path == f"/m365Copilot/StreamHub/{OID}@{TID}"
    assert query["XRoutingParameterSessionKey"] == [OID]
    assert query["chatsessionid"] == query["clientrequestid"] == ["s-1"]
    assert query["source"] == ['"officeweb"']
    assert query["access_token"] == [token]


def test_parse_frames_splits_on_record_separator():
    raw = '{"type":6}' + RECORD_SEPARATOR + "not json" + RECORD_SEPARATOR + '{"type":2}' + RECORD_SEPARATOR

    assert parse_frames(raw) == [{"type": 6}, {"type": 2}]


def test_clean_answer_text_removes_markers_and_entity_tags():
    text = "Goals \ue200cite\ue202turn1search1L3\ue201 for <Person>Contoso</Person>.  \n\ue200image\ue202x\ue201Next"

    assert clean_answer_text(text) == "Goals  for Contoso.\nNext"
    assert clean_answer_text("Done \ue200cite\ue202tur", final=False) == "Done "
    assert clean_answer_text("For <Person>Con", final=False) == "For Con"
    assert clean_answer_text("For <Pers", final=False) == "For "


def test_collector_streams_clean_text_and_collects_sources():
    collector = AnswerCollector("req-1")
    deltas = [collector.feed(frame) for frame in _turn_frames("req-1")]

    assert collector.done is True
    assert collector.text == "The plan has **two** goals  for Contoso."
    assert "".join(deltas).strip() == "The plan has **two** goals  for Contoso."
    assert collector.chat_name == "Plan summary"
    assert collector.progress == "Digging in…"
    assert collector.sources == [{"title": "Plan", "url": "https://contoso.sharepoint.com/Plan.pptx"}]


def test_collector_ignores_replay_of_earlier_turn():
    collector = AnswerCollector("req-new")
    for frame in _turn_frames("req-old"):
        collector.feed(frame)

    assert collector.done is False
    assert collector.text == ""


def test_collector_records_close_frame():
    collector = AnswerCollector("req-1")
    collector.feed({"type": 7, "error": "Connection closed with an error."})

    assert collector.done is True
    assert collector.errors == ["Connection closed with an error."]


def test_ask_sends_handshake_subscribe_and_file_request():
    socket = FakeSocket(["{}" + RECORD_SEPARATOR])
    seen: dict = {}

    def connector(url, **kwargs):
        seen.update(url=url, **kwargs)
        return socket

    streamed: list[str] = []
    answer = CopilotClient(FakeTokenManager(), connector=connector).ask(
        "Summarise", files=[copilot_file_from_drive_item(DRIVE_ITEM)], on_text=streamed.append,
    )

    assert seen["origin"] == "https://copilot.cloud.microsoft"
    assert socket.sent[0] == {"protocol": "json", "version": 1}
    subscribe = socket.sent[3]["item"]
    assert subscribe == {"type": "subscribe", "topicId": f"/conversations/{answer.conversation_id}", "action": "create"}
    request = socket.chat_request()
    assert request["isStartOfSession"] is True
    assert request["conversationId"] == answer.conversation_id
    assert request["message"]["text"] == "Summarise"
    assert request["message"]["messageAnnotations"][0]["messageAnnotationType"] == "File"
    assert answer.text == "The plan has **two** goals  for Contoso."
    assert answer.complete is True
    assert "".join(streamed).strip() == answer.text


def test_follow_up_subscribes_without_create_and_waits_for_replay():
    replay = json.dumps(_part({"author": None, "messageType": "EndOfRequest"})) + RECORD_SEPARATOR
    socket = FakeSocket(["{}" + RECORD_SEPARATOR, replay])

    answer = CopilotClient(FakeTokenManager(), connector=lambda url, **kw: socket).ask(
        "And then?", conversation_id="conv-1",
    )

    assert socket.sent[3]["item"] == {"type": "subscribe", "topicId": "/conversations/conv-1"}
    request = socket.chat_request()
    assert request["isStartOfSession"] is False
    assert "messageAnnotations" not in request["message"]
    assert answer.conversation_id == "conv-1"
    assert answer.text


def test_ask_refreshes_token_once_after_401():
    socket = FakeSocket(["{}" + RECORD_SEPARATOR])
    calls = {"n": 0}

    def connector(url, **kwargs):
        calls["n"] += 1
        if calls["n"] == 1:
            raise InvalidStatus(Response(401, "Unauthorized", Headers()))
        return socket

    manager = FakeTokenManager()
    answer = CopilotClient(manager, connector=connector).ask("Hi")

    assert manager.reauth_calls == 1
    assert calls["n"] == 2
    assert answer.complete is True


def test_ask_reports_auth_failure_when_refresh_fails():
    def connector(url, **kwargs):
        raise InvalidStatus(Response(401, "Unauthorized", Headers()))

    with pytest.raises(CopilotError, match="outlook-cli auth"):
        CopilotClient(FakeTokenManager(reauth=False), connector=connector).ask("Hi")


@pytest.mark.parametrize(
    ("ref", "site", "lookup"),
    [
        ("https://contoso.sharepoint.com/:p:/r/Plan.pptx", None, "get_share_drive_item"),
        ("Documents/Plan.pptx", None, "od_item_by_path"),
        ("01ABCDEFGHIJKLMNOPQRSTUVWXYZ", None, "drive_item_with_sharepoint_ids"),
    ],
)
def test_resolve_copilot_file_routes_each_reference(ref, site, lookup):
    plain = {key: value for key, value in DRIVE_ITEM.items() if key != "sharepointIds"}
    gc = FakeClient(
        get_share_drive_item=plain,
        od_item_by_path=plain,
        drive_item_with_sharepoint_ids=DRIVE_ITEM,
    )

    file = copilot_commands.resolve_copilot_file(gc, ref, site=site)

    assert gc.names()[0] == lookup
    # The last lookup always reads sharepointIds from the item's own drive.
    assert gc.calls[-1] == ("drive_item_with_sharepoint_ids", ("01ABCDEFGHIJKLMNOPQRSTUVWXYZ", "b!drive"), {})
    assert file.name == "Plan.pptx"


def test_resolve_copilot_file_uses_site_lookup(monkeypatch):
    found = {}

    def fake_find(gc, site, path, library):
        found.update(site=site, path=path, library=library)
        return {"id": "b!drive"}, {key: value for key, value in DRIVE_ITEM.items() if key != "sharepointIds"}

    monkeypatch.setattr(copilot_commands, "_find_sp_item_by_path", fake_find)
    gc = FakeClient(drive_item_with_sharepoint_ids=DRIVE_ITEM)

    copilot_commands.resolve_copilot_file(gc, "General/Plan.pptx", site="Team", library="Documents")

    assert found == {"site": "Team", "path": "General/Plan.pptx", "library": "Documents"}


def test_continue_without_earlier_conversation_fails(monkeypatch, tmp_path):
    monkeypatch.setattr(cache, "COPILOT_CACHE", tmp_path / "copilot-conversation.json")
    args = Namespace(conversation=None, continue_last=True)

    with pytest.raises(CliError, match="No earlier Copilot conversation"):
        copilot_commands._conversation_id(args)

    cache.save_cache(cache.COPILOT_CACHE, [{"Id": "conv-9"}])
    assert copilot_commands._conversation_id(args) == "conv-9"


def test_mcp_copilot_ask_resolves_files_and_returns_answer(monkeypatch):
    use_client(monkeypatch, copilot_tools, "files_client", FakeClient(
        get_share_drive_item=DRIVE_ITEM, drive_item_with_sharepoint_ids=DRIVE_ITEM,
    ))
    socket = FakeSocket(["{}" + RECORD_SEPARATOR])
    use_client(monkeypatch, copilot_tools, "copilot_client",
               CopilotClient(FakeTokenManager(), connector=lambda url, **kw: socket))

    result = copilot_tools.copilot_ask("Summarise", files=["https://contoso.sharepoint.com/Plan.pptx"])

    assert result["text"] == "The plan has **two** goals  for Contoso."
    assert result["files"] == [{"name": "Plan.pptx", "id": DRIVE_ITEM["id"], "driveId": "b!drive", "webUrl": DRIVE_ITEM["webUrl"]}]
    assert result["sources"][0]["title"] == "Plan"
    assert config.COPILOT_TOKEN_DOMAIN == "substrate.office.com/sydney"
