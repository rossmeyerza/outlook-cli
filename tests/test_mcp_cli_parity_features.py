"""CLI parity tests use fake clients and never change live Microsoft data."""
from __future__ import annotations

import asyncio
from unittest.mock import Mock

import httpx
import pytest
from mcp_fakes import FakeClient, use_client

from outlook_draft.errors import OutlookAPIError
from outlook_draft.mcp import calendar, contacts, drafts, files, mail, sending, tasks, teams
from outlook_draft.mcp.server import READ_ONLY_TOOLS, build_server
from outlook_draft.outlook_client import OutlookClient
from outlook_draft.safety import SafetyError, policy

CONTACT = {"Id": "contact-id", "DisplayName": "Person", "EmailAddresses": [{"Address": "person@example.test"}]}
TASK = {"Id": "task-id", "Subject": "Check", "Status": "NotStarted"}
EVENT = {"Id": "event-id", "Subject": "Check", "Type": "SingleInstance", "IsOrganizer": True,
         "ChangeKey": "version-1", "Attendees": [{"EmailAddress": {"Address": "person@example.test"}}]}
NEW_TOOLS = {"mail_download_links", "cal_rooms", "cal_update", "contact_create", "contact_update",
             "task_update", "task_delete", "teams_get", "teams_self", "files_upload", "files_mkdir",
             "files_rename", "files_move"}
MUTATIONS = NEW_TOOLS - READ_ONLY_TOOLS


@pytest.mark.parametrize("confirmed", [False, True])
def test_contact_changes_preview_and_apply(monkeypatch, confirmed):
    fake = use_client(monkeypatch, contacts, "outlook_client", FakeClient(
        get_contact=CONTACT, create_contact=CONTACT, update_contact=CONTACT))
    created = contacts.contact_create("Person", "person@example.test", company="Company", confirm=confirmed)
    updated = contacts.contact_update("contact-id", company="", confirm=confirmed)
    assert created["created"] is confirmed and updated["updated"] is confirmed
    assert ("create_contact" in fake.names()) is confirmed
    assert ("update_contact" in fake.names()) is confirmed
    if confirmed:
        assert fake.call("update_contact")[1] == {"company": ""}


@pytest.mark.parametrize("method, kwargs", [
    (contacts.contact_create, {"display_name": "", "email": "person@example.test"}),
    (contacts.contact_create, {"display_name": "Person", "email": "bad\naddress"}),
    (contacts.contact_update, {"contact_id": "1", "company": "Company"}),
    (contacts.contact_update, {"contact_id": "contact-id"}),
    (tasks.task_update, {"task_id": "task-id", "importance": "invalid"}),
    (tasks.task_update, {"task_id": "task-id", "status": "invalid"}),
    (tasks.task_update, {"task_id": "task-id", "due": "not a date"}),
    (tasks.task_update, {"task_id": "task-id"}),
    (tasks.task_delete, {"task_id": "task-id\n"}),
])
def test_invalid_contact_and_task_input_has_no_client(monkeypatch, method, kwargs):
    for module in (contacts, tasks):
        monkeypatch.setattr(module, "outlook_client", Mock(side_effect=AssertionError("No authentication")))
    with pytest.raises(ValueError):
        method(**kwargs, confirm=True)


@pytest.mark.parametrize("confirmed", [False, True])
def test_task_update_and_delete(monkeypatch, confirmed):
    fake = use_client(monkeypatch, tasks, "outlook_client", FakeClient(get_task=TASK, update_task=TASK))
    updated = tasks.task_update("task-id", due="2026-11-12 10:00", importance="High", status="Deferred", confirm=confirmed)
    deleted = tasks.task_delete("task-id", confirm=confirmed)
    assert updated["updated"] is confirmed and deleted["deleted"] is confirmed
    assert ("update_task" in fake.names()) is confirmed
    assert ("delete_task" in fake.names()) is confirmed
    assert updated["fields"]["due_dt"].startswith("2026-11-12T10:00")


@pytest.mark.parametrize("confirmed", [False, True])
def test_calendar_update_notifies_and_preserves_explicit_empty_values(monkeypatch, confirmed):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event=EVENT, update_event=EVENT))
    result = calendar.cal_update("event-id", location="", body="", online_meeting=False, confirm=confirmed)
    assert result["willNotifyAttendees"] and result["updated"] is confirmed
    assert result["confirmed"] is confirmed
    assert ("update_event" in fake.names()) is confirmed
    if confirmed:
        assert fake.call("update_event")[1] == {"location": "", "body": "", "is_online_meeting": False}


def test_calendar_update_series_and_version_checks(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event={**EVENT, "Type": "SeriesMaster"}, update_event=EVENT))
    with pytest.raises(ValueError, match="whole series"):
        calendar.cal_update("event-id", subject="Changed", confirm=True)
    with pytest.raises(ValueError, match="changed after"):
        calendar.cal_update("event-id", subject="Changed", confirm=True, allow_series=True, expected_change_key="old")
    assert "update_event" not in fake.names()
    assert calendar.cal_update("event-id", subject="Changed", confirm=True, allow_series=True)["updated"]


@pytest.mark.parametrize("changes", [{}, {"start": "2026-11-12 10:00"},
                                     {"start": "2026-11-12 11:00", "end": "2026-11-12 10:00"}])
def test_invalid_calendar_update_has_no_client(monkeypatch, changes):
    monkeypatch.setattr(calendar, "outlook_client", Mock(side_effect=AssertionError("No authentication")))
    with pytest.raises(ValueError):
        calendar.cal_update("event-id", confirm=True, **changes)


def test_calendar_rooms_and_native_suggestions(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(find_rooms=[{"Address": "room@example.test"}], find_meeting_times=[{"Confidence": 100}]))
    assert calendar.cal_rooms("room-list")[0]["Address"] == "room@example.test"
    assert fake.call("find_rooms")[1] == {"room_list": "room-list"}
    result = calendar.cal_find_time(["person@example.test"], "2026-11-12 10:00", "2026-11-12 18:00", outlook=True, count=75)
    assert result == {"mode": "outlook", "suggestions": [{"Confidence": 100}]}
    assert fake.call("find_meeting_times")[1]["max_candidates"] == 75
    assert "get_schedule" not in fake.names()


@pytest.mark.parametrize("method", [drafts.draft_reply, drafts.draft_forward])
def test_reply_and_forward_importance(monkeypatch, method):
    quoted = {"Id": "draft-id", "Body": {"ContentType": "HTML", "Content": "<p>Original</p>"}}
    fake = use_client(monkeypatch, drafts, "outlook_client", FakeClient(create_reply_draft=quoted, create_forward_draft=quoted, update_draft=quoted))
    options = {"to": ["person@example.test"]} if method is drafts.draft_forward else {}
    method("message-id", body="Reply", importance="High", **options)
    assert fake.call("update_draft")[1]["importance"] == "High"
    assert "Original" in fake.call("update_draft")[1]["body"]
    with pytest.raises(ValueError):
        method("message-id", body="Reply", importance="invalid", **options)
    assert fake.names().count("create_reply_draft" if method is drafts.draft_reply else "create_forward_draft") == 1


@pytest.mark.parametrize("signature", ["none", "new", "reply"])
def test_email_signature_and_no_save(monkeypatch, signature):
    monkeypatch.setattr(sending, "load_signature", lambda path: "SIGNATURE")
    fake = use_client(monkeypatch, sending, "outlook_client", FakeClient())
    preview = sending.mail_send(["person@example.test"], "Subject", "Body", signature=signature, save_to_sent_items=False)
    assert not fake.names() and not preview["saveToSentItems"]
    assert ("SIGNATURE" in preview["bodyHtml"]) is (signature != "none")
    sending.mail_send(["person@example.test"], "Subject", "Body", signature=signature, save_to_sent_items=False, confirm=True)
    assert fake.call("send_mail")[1]["save_to_sent_items"] is False


def test_mail_link_download_reports_partial_failure(monkeypatch, tmp_path):
    monkeypatch.setattr(mail, "mail_links", lambda *args, **kwargs: [{"url": "https://a.sharepoint.com/one"}, {"url": "https://a.sharepoint.com/two"}])
    def metadata(url):
        if url.endswith("two"):
            raise OutlookAPIError(403, "No access")
        return {"name": "../report.txt"}
    use_client(monkeypatch, mail, "graph_client", FakeClient(get_share_drive_item=metadata, download_share_url=b"report"))
    result = mail.mail_download_links("message-id", str(tmp_path))
    assert len(result["files"]) == 1 and len(result["errors"]) == 1
    assert (tmp_path / ".._report.txt").read_bytes() == b"report"
    assert "403" in result["errors"][0]["error"]


def file_context(path, drive):
    item = {"id": (path or "root") + "-id", "name": path or "Root", "folder": {}, "parentReference": {"id": "parent-id"}}
    if path == "sample.txt":
        item.pop("folder")
        item["file"] = {}
    return {"drive_id": drive, "site": "Site" if drive else None, "library": "Documents" if drive else None, "item": item}


@pytest.mark.parametrize("drive", [None, "drive-id"])
@pytest.mark.parametrize("operation", ["upload", "mkdir", "rename", "move"])
@pytest.mark.parametrize("confirmed", [False, True])
def test_all_file_changes_preview_and_apply_on_both_drives(monkeypatch, tmp_path, drive, operation, confirmed):
    prefix = "sp_" if drive else "od_"
    fake = use_client(monkeypatch, files, "files_client", FakeClient(**{
        prefix + "list_children": [], prefix + operation: {"id": "result-id", "name": "Result", "file": {}}}))
    monkeypatch.setattr(files, "_file_context", lambda client, path, site, library, **kwargs: file_context(path, drive))
    source = tmp_path / "sample.txt"
    source.write_text("Content")
    options = {"confirm": confirmed, "site": "Site" if drive else None}
    arguments = {"upload": (str(source), "Archive"), "mkdir": ("Archive/New",),
                 "rename": ("sample.txt", "renamed.txt"), "move": ("sample.txt", "Archive")}
    result = getattr(files, "files_" + operation)(*arguments[operation], **options)
    assert result["confirmed"] is confirmed
    assert (prefix + operation in fake.names()) is confirmed
    if confirmed and drive:
        assert fake.call(prefix + operation)[0][0] == drive
    if confirmed and operation == "upload":
        assert fake.call(prefix + operation)[1] == {"overwrite": False}


def test_upload_requires_explicit_overwrite(monkeypatch, tmp_path):
    source = tmp_path / "sample.txt"
    source.write_text("Content")
    fake = use_client(monkeypatch, files, "files_client", FakeClient(od_item_by_path={"id": "folder-id", "folder": {}},
        od_list_children=[{"id": "existing-id", "name": "sample.txt", "file": {}}], od_upload={"id": "existing-id", "name": "sample.txt"}))
    assert not files.files_upload(str(source))["canApply"]
    with pytest.raises(ValueError, match="overwrite"):
        files.files_upload(str(source), confirm=True)
    assert "od_upload" not in fake.names()
    assert files.files_upload(str(source), confirm=True, overwrite=True)["uploaded"]


@pytest.mark.parametrize("name", ["../escape", "bad\\name", "..", "", "bad\nname"])
def test_bad_file_names_have_no_client(monkeypatch, name):
    monkeypatch.setattr(files, "files_client", Mock(side_effect=AssertionError("No authentication")))
    with pytest.raises(ValueError):
        files.files_rename("sample.txt", name, confirm=True)


def test_file_destination_must_exist_and_be_a_folder(monkeypatch, tmp_path):
    source = tmp_path / "sample.txt"
    source.write_text("Content")
    fake = use_client(monkeypatch, files, "files_client", FakeClient(od_item_by_path={"id": "file-id", "file": {}}))
    with pytest.raises(ValueError, match="existing folder"):
        files.files_upload(str(source), confirm=True)
    assert "od_upload" not in fake.names()
    with pytest.raises(ValueError, match="explicit SharePoint site"):
        files.files_mkdir("New", library="Documents", confirm=True)


def test_file_download_accepts_urls_and_onedrive_ids(monkeypatch, tmp_path):
    remote = {"id": "remote-id", "name": "report.txt", "file": {}}
    fake = use_client(monkeypatch, files, "files_client", FakeClient(get_share_drive_item=remote, download_share_url=b"URL", od_item_by_id=remote, od_download=b"ID"))
    files.files_download("https://a.sharepoint.com/report.txt", str(tmp_path))
    assert (tmp_path / "report.txt").read_bytes() == b"URL"
    files.files_download("01ABCDEFGHIJKLMN0123456789", str(tmp_path), overwrite=True)
    assert (tmp_path / "report.txt").read_bytes() == b"ID"
    assert "od_item_by_path" not in fake.names()


def test_teams_large_counts_and_member_option(monkeypatch):
    chats = [{"id": f"chat-{i}", "chatType": "oneOnOne"} for i in range(80)]
    fake = use_client(monkeypatch, teams, "graph_client", FakeClient(list_teams_chats=chats,
        list_teams_chat_members=[{"displayName": "Person"}], get_teams_chat=chats[0], list_teams_messages=[]))
    assert len(teams.teams_list(80)) == 80
    fake.calls.clear()
    assert len(teams.teams_search("Person", scan=800, count=80, member_count=250)) == 80
    assert fake.call("list_teams_chats")[1] == {"top": 800}
    assert fake.call("list_teams_chat_members")[1] == {"top": 250}
    teams.teams_messages("chat-0", count=250)
    teams.teams_attachments("chat-0", count=250)
    assert fake.call("list_teams_messages")[1] == {"top": 250}


def test_teams_details_self_and_received_sort(monkeypatch):
    chats = [{"id": "chat-a", "topic": "A", "createdDateTime": "2026-01-01"}, {"id": "chat-b", "topic": "B"}]
    fake = use_client(monkeypatch, teams, "graph_client", FakeClient(list_teams_chats=chats, get_current_user={"id": "me"},
        get_teams_chat=chats[0], list_teams_chat_members=[{"displayName": "Person"}]))
    monkeypatch.setattr(teams, "_latest_received_message_time", lambda client, chat, user: "2" if chat["id"] == "chat-b" else "1")
    assert teams.teams_list(sort_received=True)[0]["id"] == "chat-b"
    details = teams.teams_get("chat-a")
    assert details["created"] == "2026-01-01" and details["members"][0]["displayName"] == "Person"
    assert "list_teams_messages" not in fake.names()
    monkeypatch.setattr(teams, "find_self_chat", lambda client: ({"id": "48:notes", "topic": "Self chat"}, []))
    assert teams.teams_self()["id"] == "48:notes"
    assert teams.teams_get("48:notes")["title"] == "Self chat"


def test_new_tools_annotations_and_readonly_registration():
    registered = {tool.name: tool for tool in asyncio.run(build_server().list_tools())}
    assert registered.keys() >= NEW_TOOLS
    for name in MUTATIONS:
        tool = registered[name]
        assert tool.inputSchema["properties"]["confirm"]["default"] is False
        assert tool.annotations.destructiveHint is True and tool.annotations.idempotentHint is False
    with policy(readonly=True):
        readonly = {tool.name for tool in asyncio.run(build_server().list_tools())}
    assert NEW_TOOLS & readonly == NEW_TOOLS - MUTATIONS


@pytest.mark.parametrize("name, args, kwargs", [
    ("create_contact", (), {"display_name": "Person", "email": "person@example.test"}),
    ("update_contact", ("contact/id",), {"company": "Company"}),
    ("update_task", ("task/id",), {"subject": "Changed"}),
    ("delete_task", ("task/id",), {}),
    ("update_event", ("event/id",), {"subject": "Changed"}),
    ("od_mkdir", ("parent/id", "New"), {}),
    ("od_rename", ("item/id", "New"), {}),
    ("od_move", ("item/id", "parent-id"), {}),
    ("sp_mkdir", ("drive/id", "parent-id", "New"), {}),
    ("sp_rename", ("drive/id", "item-id", "New"), {}),
    ("sp_move", ("drive/id", "item-id", "parent-id"), {}),
])
def test_new_client_mutations_have_zero_retries(name, args, kwargs):
    client = OutlookClient(Mock())
    client._request = Mock(return_value=httpx.Response(200, json={"id": "result"}))
    getattr(client, name)(*args, **kwargs)
    assert client._request.call_args.kwargs["max_retries"] == 0
    assert "%2F" in client._request.call_args.args[1] or name == "create_contact"


@pytest.mark.parametrize("sharepoint", [False, True])
@pytest.mark.parametrize("large", [False, True])
def test_upload_conflict_policy_and_zero_retries(tmp_path, sharepoint, large):
    source = tmp_path / "sample.txt"
    source.write_bytes(b"x" * (4 * 1024 * 1024 if large else 1))
    client = OutlookClient(Mock())
    client._request = Mock(return_value=httpx.Response(200, json={"id": "result", "uploadUrl": "https://upload.example.test/session"}))
    client._upload_chunks = Mock(return_value={"id": "result"})
    args = ("drive-id", "parent-id", "sample.txt", source) if sharepoint else ("parent-id", "sample.txt", source)
    getattr(client, "sp_upload" if sharepoint else "od_upload")(*args, overwrite=False)
    call = client._request.call_args.kwargs
    assert call["max_retries"] == 0
    if large:
        assert call["json_body"]["item"]["@microsoft.graph.conflictBehavior"] == "fail"
    else:
        assert call["params"]["@microsoft.graph.conflictBehavior"] == "fail"


def test_chunk_upload_obeys_readonly_before_network():
    client = OutlookClient(Mock())
    client._ensure_client = Mock(side_effect=AssertionError("No network"))
    with policy(readonly=True), pytest.raises(SafetyError):
        client._upload_chunks("https://upload.example.test/session", b"bytes")
    client._ensure_client.assert_not_called()


def test_mcp_dispatch_contact_creation_still_requires_confirmation(monkeypatch):
    fake = use_client(monkeypatch, contacts, "outlook_client", FakeClient(create_contact=CONTACT))
    server = build_server()
    args = {"display_name": "Person", "email": "person@example.test"}
    asyncio.run(server.call_tool("contact_create", args))
    assert "create_contact" not in fake.names()
    asyncio.run(server.call_tool("contact_create", {**args, "confirm": True}))
    assert fake.names().count("create_contact") == 1


def test_unknown_organiser_cannot_hide_calendar_notifications(monkeypatch):
    fake = use_client(monkeypatch, calendar, "outlook_client", FakeClient(get_event={**EVENT, "IsOrganizer": None}))
    preview = calendar.cal_update("event-id", subject="Changed")
    assert not preview["canApply"]
    with pytest.raises(ValueError, match="notifies attendees"):
        calendar.cal_update("event-id", subject="Changed", confirm=True)
    assert "update_event" not in fake.names()


@pytest.mark.parametrize("operation", ["rename", "move"])
def test_file_roots_cannot_change(monkeypatch, operation):
    fake = use_client(monkeypatch, files, "files_client", FakeClient(od_item_by_path={"id": "root-id", "name": "Root", "folder": {}}, od_list_children=[]))
    with pytest.raises(ValueError, match="root"):
        getattr(files, "files_" + operation)("Root", "New", confirm=True)
    assert "od_" + operation not in fake.names()


def test_file_move_never_replaces_a_collision(monkeypatch):
    fake = use_client(monkeypatch, files, "files_client", FakeClient(od_item_by_path=lambda path: file_context(path, None)["item"],
        od_list_children=[{"id": "different-id", "name": "sample.txt", "file": {}}]))
    with pytest.raises(ValueError, match="already exists"):
        files.files_move("sample.txt", "Archive", confirm=True)
    assert "od_move" not in fake.names()


def test_inline_mail_download_and_folder_count(monkeypatch, tmp_path):
    fake = use_client(monkeypatch, mail, "outlook_client", FakeClient(list_mail_folders=[],
        list_message_attachments=[{"Id": "inline-id", "IsInline": True}],
        get_message_attachment={"Name": "logo.png", "ContentBytes": "eA=="}))
    mail.mail_folders(count=750)
    assert fake.call("list_mail_folders")[1] == {"top": 750}
    result = mail.mail_download_attachments("message-id", str(tmp_path), include_inline=True)
    assert result == [str(tmp_path / "logo.png")]
    assert (tmp_path / "logo.png").read_bytes() == b"x"


@pytest.mark.parametrize("module, method, client_method, kwargs, expected", [
    (calendar, "cal_agenda", "get_agenda", {"days": 120, "count": 500}, {"days": 120, "top": 500}),
    (drafts, "draft_list", "list_drafts", {"count": 500}, {"top": 500}),
    (contacts, "contact_search", "search_people", {"query": "Person", "count": 500}, {"top": 500}),
    (tasks, "task_list", "list_tasks", {"count": 500}, {"top": 500}),
])
def test_requested_counts_have_no_extra_mcp_upper_limit(monkeypatch, module, method, client_method, kwargs, expected):
    fake = use_client(monkeypatch, module, "outlook_client", FakeClient(**{client_method: []}))
    getattr(module, method)(**kwargs)
    assert fake.call(client_method)[1] == expected


def test_remote_paths_encode_reserved_url_characters():
    client = OutlookClient(Mock())
    client._request = Mock(return_value=httpx.Response(200, json={}))
    client.od_item_by_path("Reports/Q4 #1?.txt")
    assert client._request.call_args.args[1].endswith("/Reports/Q4%20%231%3F.txt")
