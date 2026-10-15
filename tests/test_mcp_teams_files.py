from __future__ import annotations

from pathlib import Path

import pytest
from mcp_fakes import FakeClient, use_client

from outlook_draft.errors import OutlookAPIError
from outlook_draft.mcp import files, teams

CHAT = {"id": "19:abc@thread.v2", "topic": "Project X", "chatType": "group", "lastUpdatedDateTime": "2026-10-15T10:00:00Z"}
ONE_ON_ONE = {"id": "19:def@unq.gbl.spaces", "chatType": "oneOnOne", "lastUpdatedDateTime": "2026-10-14T10:00:00Z"}
MEMBERS = [{"displayName": "Dana Smith", "email": "dana@example.com", "userId": "u1"}]


def test_teams_list_sorts_and_resolves_titles(monkeypatch):
    fake = FakeClient(
        list_teams_chats=[ONE_ON_ONE, CHAT],
        list_teams_chat_members=MEMBERS,
    )
    use_client(monkeypatch, teams, "graph_client", fake)

    result = teams.teams_list(count=5)

    assert [c["title"] for c in result] == ["Project X", "Dana Smith"]
    assert fake.call("list_teams_chats")[1] == {"top": 5}


def test_teams_list_survives_member_lookup_failure(monkeypatch):
    def boom(*args, **kwargs):
        raise OutlookAPIError(403, "no")

    use_client(monkeypatch, teams, "graph_client", FakeClient(list_teams_chats=[ONE_ON_ONE], list_teams_chat_members=boom))

    assert teams.teams_list()[0]["title"].startswith("oneOnOne")


def test_teams_search_matches_topic_and_member(monkeypatch):
    fake = FakeClient(list_teams_chats=[ONE_ON_ONE, CHAT], list_teams_chat_members=MEMBERS)
    use_client(monkeypatch, teams, "graph_client", fake)

    assert [c["id"] for c in teams.teams_search("project")] == [CHAT["id"]]
    assert [c["match"] for c in teams.teams_search("dana")] == ["member: Dana Smith"]
    with pytest.raises(ValueError):
        teams.teams_search("  ")


def test_teams_messages_orders_and_strips_html(monkeypatch):
    messages = [
        {"id": "m2", "createdDateTime": "2026-10-15T10:05:00Z", "messageType": "message",
         "from": {"user": {"displayName": "Dana"}}, "body": {"contentType": "html", "content": "<p>second</p>"}},
        {"id": "m1", "createdDateTime": "2026-10-15T10:00:00Z", "messageType": "message",
         "from": {"user": {"displayName": "Ross"}}, "body": {"contentType": "text", "content": "first"}},
        {"id": "m0", "createdDateTime": "2026-10-15T09:00:00Z", "messageType": "systemEventMessage", "body": {}},
    ]
    fake = FakeClient(get_teams_chat=CHAT, list_teams_chat_members=[], list_teams_messages=messages)
    use_client(monkeypatch, teams, "graph_client", fake)

    result = teams.teams_messages(CHAT["id"], count=10)

    assert [m["text"] for m in result["messages"]] == ["[system event]", "first", "second"]
    assert result["messages"][2]["sender"] == "Dana"
    assert fake.call("list_teams_messages") == ((CHAT["id"],), {"top": 10})


def test_teams_messages_self_chat_skips_chat_lookup(monkeypatch):
    fake = FakeClient(list_teams_messages=[])
    use_client(monkeypatch, teams, "graph_client", fake)

    assert teams.teams_messages("48:notes")["chat"]["title"] == "Self chat"
    assert "get_teams_chat" not in fake.names()


def _attachment_message():
    return {
        "messageType": "message",
        "createdDateTime": "2026-10-15T10:00:00Z",
        "from": {"user": {"displayName": "Dana"}},
        "attachments": [{"contentType": "reference", "name": "deck.pptx", "contentUrl": "https://a.sharepoint.com/deck"}],
        "body": {"contentType": "text", "content": ""},
    }


def test_teams_attachments_and_download(monkeypatch, tmp_path: Path):
    fake = FakeClient(
        list_teams_messages=[_attachment_message()],
        get_share_drive_item={"name": "deck.pptx"},
        download_share_url=b"bytes",
    )
    use_client(monkeypatch, teams, "graph_client", fake)

    items = teams.teams_attachments(CHAT["id"])
    written = teams.teams_download_attachments(CHAT["id"], str(tmp_path / "dl"), attachment="1")

    assert items[0]["name"] == "deck.pptx"
    assert written == [str(tmp_path / "dl" / "deck.pptx")]
    assert (tmp_path / "dl" / "deck.pptx").read_bytes() == b"bytes"


def test_teams_download_requires_absolute_dir(monkeypatch):
    use_client(monkeypatch, teams, "graph_client", FakeClient())

    with pytest.raises(ValueError, match="absolute"):
        teams.teams_download_attachments(CHAT["id"], "dl")


GROUP = {"id": "g1", "displayName": "Tesco", "mail": "tesco@example.com"}
DRIVE = {"id": "d1", "name": "Documents", "driveType": "documentLibrary", "webUrl": "https://sp/docs"}
FILE = {"id": "ITEM1", "name": "plan.xlsx", "size": 5, "webUrl": "https://sp/plan",
        "file": {}, "parentReference": {"path": "/drives/d1/root:/General"}}


def _sp_fake(**extra):
    return FakeClient(
        sp_list_sites=[GROUP],
        sp_get_site={"id": "s1"},
        sp_list_drives=[DRIVE],
        **extra,
    )


def test_files_sites_and_libraries(monkeypatch):
    use_client(monkeypatch, files, "files_client", _sp_fake())

    assert files.files_sites()[0]["name"] == "Tesco"
    assert files.files_libraries("tes") == [
        {"id": "d1", "name": "Documents", "type": "documentLibrary", "webUrl": "https://sp/docs"}
    ]


def test_files_list_site_without_path_lists_libraries(monkeypatch):
    use_client(monkeypatch, files, "files_client", _sp_fake())

    result = files.files_list(site="Tesco")

    assert result["kind"] == "libraries"
    assert result["items"][0]["name"] == "Documents"


def test_files_list_onedrive_root(monkeypatch):
    fake = FakeClient(od_item_by_path={"id": "root"}, od_list_children=[FILE])
    use_client(monkeypatch, files, "files_client", fake)

    result = files.files_list()

    assert result["location"] == "OneDrive"
    assert result["items"][0]["id"] == "ITEM1"
    assert fake.call("od_item_by_path")[0] == ("",)


def test_files_search_requires_site_and_maps_count(monkeypatch):
    fake = _sp_fake(sp_search=[FILE])
    use_client(monkeypatch, files, "files_client", fake)

    with pytest.raises(ValueError, match="requires 'site'"):
        files.files_search("plan")
    result = files.files_search("plan", site="Tesco", library="Documents", count=7)

    assert fake.call("sp_search")[0] == ("d1", "plan", 7)
    assert result[0]["library"] == "Documents"
    assert result[0]["webUrl"] == "https://sp/plan"


def test_files_download_onedrive_writes_file(monkeypatch, tmp_path: Path):
    fake = FakeClient(od_item_by_path={"id": "i1", "name": "a.txt"}, od_download=b"hello")
    use_client(monkeypatch, files, "files_client", fake)

    result = files.files_download("Docs/a.txt", str(tmp_path) + "/")

    assert result == {"name": "a.txt", "path": str(tmp_path / "a.txt"), "size": 5}
    with pytest.raises(FileExistsError):
        files.files_download("Docs/a.txt", str(tmp_path / "a.txt"))
    assert files.files_download("Docs/a.txt", str(tmp_path / "a.txt"), overwrite=True)["size"] == 5


def test_files_download_rejects_folders_and_relative_dest(monkeypatch, tmp_path: Path):
    use_client(monkeypatch, files, "files_client", FakeClient(od_item_by_path={"id": "i", "name": "d", "folder": {}}))

    with pytest.raises(ValueError, match="not folders"):
        files.files_download("Docs", str(tmp_path))
    with pytest.raises(ValueError, match="absolute"):
        files.files_download("Docs/a.txt", "relative")
