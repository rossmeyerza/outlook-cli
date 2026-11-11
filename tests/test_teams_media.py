"""Synthetic Graph checks for Teams media and its CLI and MCP entry points."""
from __future__ import annotations

import base64
import json
import os
import re
from argparse import ArgumentParser, Namespace
from contextlib import nullcontext
from uuid import UUID

import httpx
import pytest

from outlook_draft import teams_media as media
from outlook_draft import teams_new
from outlook_draft.commands import teams
from outlook_draft.commands import teams_media as commands
from outlook_draft.errors import CliError, OutlookAPIError
from outlook_draft.mcp import sending
from outlook_draft.safety import SafetyError, policy

ME = "00000000-0000-0000-0000-000000000001"
ANA = "00000000-0000-0000-0000-000000000002"
BEN = "00000000-0000-0000-0000-000000000003"
GUID = "12345678-1234-1234-1234-123456789abc"
URL = "https://example.sharepoint.com/personal/test/file.pdf"
PNG = b"\x89PNG\r\n\x1a\nprivate-image-bytes"


def member(user_id, **overrides):
    return {"userId": user_id, "tenantId": "tenant", "roles": [], **overrides}


class Graph:
    def __init__(self):
        self.calls = []
        self.events = []
        self.members = [member(ME), member(ANA), member(BEN)]
        self.folder = {"id": "folder", "folder": {}}
        self.chat = {"id": "chat", "chatType": "group", "tenantId": "tenant"}
        self.metadata = {"eTag": f'"{{{GUID}}},1"', "webDavUrl": URL}
        self.grants = [ANA, BEN]
        self.permissions = None
        self.invite_status = 200
        self.fail = None
        self.puts = 0
        self.closed = False

    def _request(self, method, path, *, params=None, json_body=None, data=None,
                 extra_headers=None, max_retries=None):
        call = dict(method=method, path=path, params=params, json_body=json_body,
                    data=data, extra_headers=extra_headers, max_retries=max_retries)
        self.calls.append(call)
        self.events.append((method, path))
        if self.fail and self.fail(call, self):
            raise OutlookAPIError(503, "synthetic uncertain result")
        status, body = self._response(method, path)
        return httpx.Response(status, json=body, request=httpx.Request(method, "https://graph.test" + path))

    def _response(self, method, path):
        if path.startswith("/me/drive/root:"):
            return 200, self.folder
        if method == "PUT":
            self.puts += 1
            return 201, {"id": f"item-{self.puts}", "webUrl": URL}
        if path.endswith("/invite"):
            grants = [{"grantedToV2": {"user": {"id": value}}} for value in self.grants]
            return self.invite_status, {"value": self.permissions if self.permissions is not None else grants}
        if method == "GET" and path.startswith("/me/drive/items/"):
            return 200, self.metadata
        if path.endswith("/messages"):
            return 201, {"id": "message"}
        raise AssertionError((method, path))

    def get_current_user(self):
        self.events.append(("me",))
        return {"id": ME}

    def get_teams_chat(self, chat_id):
        self.events.append(("chat", chat_id))
        return self.chat

    def list_teams_chat_members(self, chat_id, *, top):
        self.events.append(("members", chat_id, top))
        return self.members

    def list_teams_message_metadata(self, chat_id, *, top):
        self.events.append(("self", chat_id, top))
        return []

    def get_directory_user(self, ref):
        self.events.append(("directory", ref))
        return {"id": ANA, "mail": "ana@example.test", "userType": "Member"}

    def create_teams_chat(self, me, users, *, topic):
        self.events.append(("create", me, users, topic))
        return {"id": "created-chat"}

    def send_teams_message(self, *args, **kwargs):
        raise AssertionError("Media must use the media message flow")

    def close(self):
        self.closed = True


@pytest.fixture
def graph():
    return Graph()


@pytest.fixture
def paths(tmp_path):
    document = tmp_path / "report.pdf"
    document.write_bytes(b"private-document-bytes")
    image = tmp_path / "image.png"
    image.write_bytes(PNG)
    return str(document), str(image)


def prepared(paths):
    return media.prepare_media([paths[0]], [paths[1]])


def writes(graph):
    return [call for call in graph.calls if call["method"] != "GET"]


def bind_mcp(monkeypatch, graph):
    monkeypatch.setattr(sending, "graph_client", lambda: nullcontext(graph))


def bind_cli(monkeypatch, graph):
    output = []
    monkeypatch.setattr(commands.runtime, "get_graph_client", lambda: graph)
    monkeypatch.setattr(commands.runtime, "print_json", output.append)
    monkeypatch.setattr(teams, "_resolve_teams_chat_id", lambda args, ref: ref)
    return output


def cli_args(paths, **overrides):
    return Namespace(**dict(dict(chat_id="chat", message="", attach=[paths[0]], image=[paths[1]],
                                 upload_folder=media.DEFAULT_UPLOAD_FOLDER, yes=False, html=False,
                                 to=["ana@example.test"], topic=None), **overrides))


def test_prepare_all_local_before_graph(monkeypatch, graph, paths, tmp_path):
    bind_mcp(monkeypatch, graph)
    with pytest.raises(ValueError, match="Cannot read"):
        sending.teams_send("chat", attachments=[paths[0], str(tmp_path / "absent")], confirm=True)
    assert graph.events == []
    assert graph.calls == []


@pytest.mark.parametrize("kind", ["fifo", "symlink", "directory", "broken-link"])
def test_only_regular_files(tmp_path, kind):
    target = tmp_path / "target.pdf"
    target.write_bytes(b"data")
    path = tmp_path / "input.pdf"
    if kind == "fifo":
        os.mkfifo(path)
    elif kind == "directory":
        path.mkdir()
    else:
        path.symlink_to(target if kind == "symlink" else tmp_path / "missing")
    with pytest.raises(ValueError):
        media.prepare_media([str(path)])


@pytest.mark.parametrize("path", ["", "   ", None, 123])
def test_bad_local_paths(path):
    with pytest.raises(ValueError, match="non-empty"):
        media.prepare_media([path])


def test_count_before_file_access(monkeypatch):
    monkeypatch.setattr(media, "_read_file", lambda *args: pytest.fail("Must check count first"))
    with pytest.raises(ValueError, match="At most 10"):
        media.prepare_media(["missing"] * 6, ["missing"] * 5)


@pytest.mark.parametrize("name,signature,mime", [("a.png", PNG, "image/png"),
                          ("a.jpg", b"\xff\xd8\xffx", "image/jpeg"),
                          ("a.jpeg", b"\xff\xd8\xffx", "image/jpeg"),
                          ("a.gif", b"GIF87ax", "image/gif"),
                          ("a.GIF", b"GIF89ax", "image/gif")])
def test_supported_inline_types(tmp_path, name, signature, mime):
    path = tmp_path / name
    path.write_bytes(signature)
    result = media.prepare_media(images=[str(path)])
    assert result.images[0].content_type == mime
    assert result.images[0].data == signature


@pytest.mark.parametrize("name,data", [("a.png", b"not png"), ("a.jpg", PNG),
                                      ("a.svg", b"<svg/>"), ("a.webp", b"RIFF")])
def test_inline_type_must_match_signature(tmp_path, name, data):
    path = tmp_path / name
    path.write_bytes(data)
    with pytest.raises(ValueError, match="matching extension"):
        media.prepare_media(images=[str(path)])


@pytest.mark.parametrize("image", [False, True])
def test_per_file_size_boundary(monkeypatch, tmp_path, image):
    monkeypatch.setattr(media, "MAX_IMAGE_BYTES" if image else "MAX_FILE_BYTES", 16)
    path = tmp_path / "a.png"
    path.write_bytes(PNG[:8] + b"x" * 8)
    kwargs = {"images" if image else "attachments": [str(path)]}
    assert media.prepare_media(**kwargs)
    path.write_bytes(path.read_bytes() + b"x")
    with pytest.raises(ValueError, match="exceeds"):
        media.prepare_media(**kwargs)


@pytest.mark.parametrize("limit,files,images", [("MAX_TOTAL_BYTES", 2, 0),
                         ("MAX_TOTAL_BYTES", 1, 1), ("MAX_IMAGE_TOTAL", 0, 2)])
def test_total_size_limits(monkeypatch, paths, limit, files, images):
    monkeypatch.setattr(media, limit, 30)
    with pytest.raises(ValueError, match="total size"):
        media.prepare_media([paths[0]] * files, [paths[1]] * images)


def test_summary_and_repr_hide_bytes(paths):
    result = prepared(paths)
    summary = json.dumps(result.summary())
    assert "private-" not in summary
    assert "private-" not in repr(result)
    assert result.summary()["attachments"][0]["size"] == len(b"private-document-bytes")


@pytest.mark.parametrize("mode", ["readonly", "no_send"])
def test_policy_before_any_graph_request(graph, paths, mode):
    with policy(**{mode: True}), pytest.raises(SafetyError):
        media.send_media(graph, "chat", "", prepared(paths))
    assert graph.events == []


@pytest.mark.parametrize("folder", [{}, {"id": "file"}, {"id": "folder", "folder": {}, "shared": {"scope": "users"}}])
def test_destination_must_exist_and_be_private(graph, paths, folder):
    graph.folder = folder
    with pytest.raises(ValueError):
        media.send_media(graph, "chat", "", prepared(paths))
    assert writes(graph) == []


@pytest.mark.parametrize("path", ["", "/root", "a//b", "../a", "a/./b", "a/"])
def test_invalid_destination_before_graph(graph, path):
    with pytest.raises(ValueError):
        media.resolve_upload_folder(graph, path)
    assert graph.calls == []


def test_private_nested_folder_is_read_only_plan(graph, paths):
    result = media.media_plan(graph, "chat", "", prepared(paths), "Private Files/Nested")
    assert graph.calls[0]["path"] == "/me/drive/root:/Private%20Files/Nested"
    assert result["accessUserIds"] == [ANA, BEN]
    assert writes(graph) == []
    assert "private-document-bytes" not in json.dumps(result)
    assert "private-image-bytes" not in json.dumps(result)


@pytest.mark.parametrize("change", [{"roles": ["guest"]}, {"tenantId": "external"},
                      {"tenantId": None}, {"userId": "unresolved"},
                      {"@odata.type": "#microsoft.graph.anonymousGuestConversationMember"}])
def test_untrusted_members_prevent_upload(graph, paths, change):
    graph.members[1].update(change)
    with pytest.raises(ValueError):
        media.send_media(graph, "chat", "", prepared(paths))
    assert writes(graph) == []


def test_member_limit_prevents_upload(graph, paths):
    graph.members = [member(ME)] + [member(ANA)] * 1000
    with pytest.raises(ValueError, match="1000"):
        media.send_media(graph, "chat", "", prepared(paths))
    assert writes(graph) == []


@pytest.mark.parametrize("case", ["absent-self", "wrong-own-tenant", "unknown-tenant", "channel"])
def test_membership_and_tenant_required(graph, paths, case):
    if case == "absent-self":
        graph.members = [member(ANA)]
    elif case == "wrong-own-tenant":
        graph.members[0]["tenantId"] = "other"
    elif case == "unknown-tenant":
        graph.chat.pop("tenantId")
        for entry in graph.members:
            entry.pop("tenantId")
    else:
        graph.chat["chatType"] = "channel"
    with pytest.raises(ValueError):
        media.send_media(graph, "chat", "", prepared(paths))
    assert writes(graph) == []


def test_current_members_are_rechecked_after_preview(graph, paths):
    media.media_plan(graph, "chat", "", prepared(paths))
    graph.members[1]["roles"] = ["guest"]
    with pytest.raises(ValueError):
        media.send_media(graph, "chat", "", prepared(paths))
    assert writes(graph) == []
    assert graph.events.count(("members", "chat", 1001)) == 2


def test_own_tenant_fallback_and_duplicate_members(graph):
    graph.chat.pop("tenantId")
    graph.members.append(member(ANA))
    assert media.file_audience(graph, "chat") == [ANA, BEN]


def test_self_notes_upload_without_invites(graph, paths):
    result = media.send_media(graph, "48:notes", "", prepared(paths))
    assert result["sent"] is True
    assert ("self", "48:notes", 1) in graph.events
    assert not any(call["path"].endswith("/invite") for call in graph.calls)
    assert not any(event[0] in {"chat", "members"} for event in graph.events)
    assert writes(graph)[-1]["path"] == "/chats/48%3Anotes/messages"


def test_upload_names_bytes_and_no_retries(graph, paths):
    result = media.send_media(graph, "chat", "", prepared(paths))
    upload = next(call for call in graph.calls if call["method"] == "PUT")
    artifact = result["uploads"][0]
    assert UUID(artifact["remoteName"].split("-")[0]).version == 4
    assert re.fullmatch(r"[0-9a-f]{32}-report.pdf", artifact["remoteName"])
    assert upload["data"] == b"private-document-bytes"
    assert upload["extra_headers"] == {"Content-Type": "application/pdf"}
    assert all(call["max_retries"] == 0 for call in writes(graph))
    assert artifact["status"] == "uploaded"


def test_documented_reference_and_hosted_payload(graph, paths):
    media.send_media(graph, "chat", "<text>\nnext", prepared(paths))
    payload = writes(graph)[-1]["json_body"]
    assert payload["attachments"] == [{"id": GUID, "contentType": "reference", "contentUrl": URL, "name": "report.pdf"}]
    assert f'<attachment id="{GUID}"></attachment>' in payload["body"]["content"]
    assert payload["body"]["contentType"] == "html"
    assert "&lt;text&gt;<br>next" in payload["body"]["content"]
    hosted = payload["hostedContents"][0]
    assert hosted["@microsoft.graph.temporaryId"] == "1"
    assert hosted["contentType"] == "image/png"
    assert base64.b64decode(hosted["contentBytes"]) == PNG
    assert '../hostedContents/1/$value' in payload["body"]["content"]


def test_multiple_hosted_ids_match_and_html_is_preserved(graph, paths):
    image_media = media.prepare_media(images=[paths[1], paths[1]])
    media.send_media(graph, "chat", "<b>Hello</b>", image_media, is_html=True)
    payload = writes(graph)[0]["json_body"]
    assert payload["body"]["content"].startswith("<b>Hello</b>")
    for entry in payload["hostedContents"]:
        assert f'../hostedContents/{entry["@microsoft.graph.temporaryId"]}/$value' in payload["body"]["content"]
    assert len(graph.calls) == 1


def test_per_user_read_grants_without_links(graph, paths):
    media.send_media(graph, "chat", "", prepared(paths))
    invite = next(call for call in graph.calls if call["path"].endswith("/invite"))
    assert invite["json_body"] == {"recipients": [{"objectId": ANA}, {"objectId": BEN}],
                                  "requireSignIn": True, "sendInvitation": False,
                                  "roles": ["read"], "retainInheritedPermissions": True}
    assert not any("createLink" in call["path"] for call in graph.calls)
    assert graph.events.index(("members", "chat", 1001)) < next(i for i, event in enumerate(graph.events) if event[0] == "PUT")


@pytest.mark.parametrize("status,grants,permissions", [(207, [ANA, BEN], None),
        (200, [ANA], None), (200, [], None),
        (200, [ANA, BEN], [{"error": {"code": "denied"}}])])
def test_every_grant_must_be_confirmed(graph, paths, status, grants, permissions):
    graph.invite_status, graph.grants, graph.permissions = status, grants, permissions
    with pytest.raises(media.MediaSendError) as failure:
        media.send_media(graph, "chat", "", prepared(paths))
    assert failure.value.phase == "file access"
    assert failure.value.uploads[0]["id"] == "item-1"
    assert not any(call["path"].startswith("/chats/") for call in graph.calls)
    assert len([call for call in graph.calls if call["path"].endswith("/invite")]) == 1


@pytest.mark.parametrize("identity_key", ["grantedToIdentitiesV2", "grantedToIdentities", "grantedToV2", "grantedTo"])
def test_grant_identity_response_shapes(graph, identity_key):
    identities = [{"user": {"id": value}} for value in [ANA, BEN]]
    graph.permissions = ([{identity_key: identities}] if "Identities" in identity_key
                         else [{identity_key: identity} for identity in identities])
    media._grant_access(graph, "item", [ANA, BEN])
    assert len(graph.calls) == 1


@pytest.mark.parametrize("phase", ["upload", "file access", "message send", "second upload"])
def test_partial_failures_keep_artifacts_without_retry_or_delete(graph, paths, phase):
    def fail(call, client):
        if phase == "upload":
            return call["method"] == "PUT"
        if phase == "second upload":
            return call["method"] == "PUT" and client.puts == 1
        return call["path"].endswith("/invite" if phase == "file access" else "/messages")
    graph.fail = fail
    files = media.prepare_media([paths[0]] * (2 if phase == "second upload" else 1))
    with pytest.raises(media.MediaSendError) as failure:
        media.send_media(graph, "chat", "", files)
    error = failure.value
    assert error.chat_id == "chat"
    assert error.phase == ("upload" if phase == "second upload" else phase)
    assert len(error.uploads) == (2 if phase == "second upload" else 1)
    assert error.uploads[-1]["status"] == ("uncertain" if "upload" in phase else "uploaded")
    assert all(item["remoteName"] and item["folderId"] == "folder" for item in error.uploads)
    assert "No automatic retry or cleanup" in str(error)
    assert not any(call["method"] == "DELETE" for call in graph.calls)
    assert all(call["max_retries"] == 0 for call in writes(graph))
    assert len(writes(graph)) == {"upload": 1, "file access": 2, "message send": 3, "second upload": 3}[phase]


@pytest.mark.parametrize("metadata", [{"eTag": "bad", "webDavUrl": URL},
    {"eTag": f'"{{{GUID}}},1"', "webDavUrl": "http://example.sharepoint.com/x"},
    {"eTag": f'"{{{GUID}}},1"', "webDavUrl": "https://other.sharepoint.com/x"}])
def test_bad_reference_metadata_keeps_uploaded_file(graph, paths, metadata):
    graph.metadata = metadata
    with pytest.raises(media.MediaSendError) as failure:
        media.send_media(graph, "chat", "", prepared(paths))
    assert failure.value.uploads[0]["status"] == "uploaded"
    assert failure.value.uploads[0]["webUrl"] == URL
    assert len(writes(graph)) == 1


@pytest.mark.parametrize("confirm", [False, True, 1, "true"])
def test_mcp_requires_explicit_boolean_confirmation(monkeypatch, graph, paths, confirm):
    bind_mcp(monkeypatch, graph)
    result = sending.teams_send("chat", attachments=[paths[0]], images=[paths[1]], confirm=confirm)
    assert result["sent"] is (confirm is True)
    assert bool(writes(graph)) is (confirm is True)
    if confirm is not True:
        assert "private-" not in json.dumps(result)


def test_mcp_image_preview_needs_no_client(monkeypatch, paths):
    monkeypatch.setattr(sending, "graph_client", lambda: pytest.fail("Preview must remain local"))
    assert sending.teams_send("chat", images=[paths[1]])["sent"] is False


@pytest.mark.parametrize("command", ["send", "new"])
def test_cli_repeatable_media_flags_and_optional_text(command):
    parser = ArgumentParser()
    teams.register(parser.add_subparsers(dest="group"))
    target = ["chat"] if command == "send" else ["--to", "ana@example.test"]
    args = parser.parse_args(["teams", command, *target, "--attach", "a.pdf", "--attach", "b.pdf",
                             "--image", "a.png", "--image", "b.png", "--upload-folder", "Private"])
    assert args.message == ""
    assert args.attach == ["a.pdf", "b.pdf"]
    assert args.image == ["a.png", "b.png"]
    assert args.upload_folder == "Private"
    assert args.yes is False


@pytest.mark.parametrize("new", [False, True])
def test_cli_preview_never_writes(monkeypatch, graph, paths, new):
    output = bind_cli(monkeypatch, graph)
    with pytest.raises(CliError) as failure:
        (teams.cmd_teams_new if new else teams.cmd_teams_send)(cli_args(paths))
    assert failure.value.exit_code == 2
    assert output[0]["sent"] is False
    assert "private-" not in json.dumps(output)
    assert writes(graph) == []
    assert not any(event[0] == "create" for event in graph.events)
    assert graph.closed


@pytest.mark.parametrize("new", [False, True])
def test_cli_confirmed_media_only(monkeypatch, graph, paths, new):
    output = bind_cli(monkeypatch, graph)
    (teams.cmd_teams_new if new else teams.cmd_teams_send)(cli_args(paths, yes=True))
    assert output[0]["sent"] is True
    assert graph.closed
    assert writes(graph)[-1]["path"].endswith("/messages")


@pytest.mark.parametrize("confirm", [False, True])
def test_new_chat_mcp_uses_existing_create_flow(monkeypatch, graph, paths, confirm):
    bind_mcp(monkeypatch, graph)
    result = sending.teams_new(["ana@example.test"], attachments=[paths[0]], confirm=confirm)
    assert result["sent"] is confirm
    creates = [event for event in graph.events if event[0] == "create"]
    assert creates == ([("create", ME, [ANA], None)] if confirm else [])
    if confirm:
        assert result["chat_id"] == "created-chat"
        assert graph.events.index(creates[0]) < next(i for i, event in enumerate(graph.events) if event[0] == "PUT")
    else:
        assert writes(graph) == []


@pytest.mark.parametrize("bad", ["folder", "content"])
def test_start_chat_prevalidates_before_create(graph, paths, bad):
    if bad == "folder":
        graph.folder = {"id": "not-folder"}
    with pytest.raises(ValueError):
        teams_new.start_chat(graph, {"id": ME}, [{"id": ANA}],
                             "x" * 20001 if bad == "content" else "",
                             media=prepared(paths))
    assert not any(event[0] == "create" for event in graph.events)
    assert writes(graph) == []


@pytest.mark.parametrize("command", [teams.cmd_teams_send, teams.cmd_teams_new])
def test_cli_without_text_or_media_rejects_before_client(monkeypatch, paths, command):
    monkeypatch.setattr(commands.runtime, "get_graph_client", lambda: pytest.fail("Must validate first"))
    with pytest.raises(CliError, match="message|text|empty"):
        command(cli_args(paths, attach=None, image=None))


@pytest.mark.parametrize("content", ["", "   ", "é" * 10001])
def test_invalid_content_has_no_remote_calls(graph, content):
    with pytest.raises(ValueError):
        media.send_media(graph, "chat", content, media.PreparedMedia())
    assert graph.calls == []


@pytest.mark.parametrize("kind", ["upload", "message"])
def test_incomplete_write_response_preserves_known_metadata(monkeypatch, graph, paths, kind):
    response = graph._response
    def incomplete(method, path):
        if kind == "upload" and method == "PUT":
            return 201, {"id": "known-upload-id"}
        if kind == "message" and path.endswith("/messages"):
            return 201, {}
        return response(method, path)
    monkeypatch.setattr(graph, "_response", incomplete)
    with pytest.raises(media.MediaSendError) as failure:
        media.send_media(graph, "chat", "", prepared(paths))
    assert failure.value.phase == ("upload" if kind == "upload" else "message send")
    assert failure.value.uploads[0]["id"] == ("known-upload-id" if kind == "upload" else "item-1")
    assert not any(call["method"] == "DELETE" for call in graph.calls)


def test_new_chat_media_failure_preserves_created_id(graph, paths):
    graph.fail = lambda call, client: call["path"].endswith("/messages")
    with pytest.raises(teams_new.ChatCreatedSendFailed, match="created-chat") as failure:
        teams_new.start_chat(graph, {"id": ME}, [{"id": ANA}], "", media=prepared(paths))
    assert failure.value.chat == {"id": "created-chat"}
    assert isinstance(failure.value.error, media.MediaSendError)
    assert failure.value.error.uploads[0]["id"] == "item-1"
    assert len([event for event in graph.events if event[0] == "create"]) == 1
    assert len([call for call in graph.calls if call["path"].endswith("/messages")]) == 1


@pytest.mark.parametrize("new", [False, True])
@pytest.mark.parametrize("mode", ["readonly", "no_send"])
def test_mcp_policy_prevents_media_writes(monkeypatch, graph, paths, new, mode):
    bind_mcp(monkeypatch, graph)
    with policy(**{mode: True}), pytest.raises(SafetyError):
        if new:
            sending.teams_new(["ana@example.test"], attachments=[paths[0]], confirm=True)
        else:
            sending.teams_send("chat", attachments=[paths[0]], confirm=True)
    assert writes(graph) == []
    assert not any(event[0] == "create" for event in graph.events)


@pytest.mark.parametrize("new", [False, True])
@pytest.mark.parametrize("content", ['"' * 20000, '&' * 6000])
def test_expanded_html_is_rejected_before_remote_access(monkeypatch, graph, paths, new, content):
    bind_mcp(monkeypatch, graph)
    with pytest.raises(ValueError, match="final Teams HTML"):
        if new:
            sending.teams_new(["ana@example.test"], content, attachments=[paths[0]], confirm=True)
        else:
            sending.teams_send("chat", content, attachments=[paths[0]], confirm=True)
    assert graph.events == []
    assert graph.calls == []


@pytest.mark.parametrize("new", [False, True])
def test_cli_expanded_html_is_rejected_before_client(monkeypatch, paths, new):
    monkeypatch.setattr(commands.runtime, "get_graph_client", lambda: pytest.fail("No client before validation"))
    with pytest.raises(CliError, match="final Teams HTML"):
        (teams.cmd_teams_new if new else teams.cmd_teams_send)(cli_args(paths, message='"' * 20000, yes=True))


def test_explicit_html_and_body_size_estimate_match_actual_payload(graph, paths):
    text = '"' * 20000
    result = media.media_plan(graph, "chat", text, prepared(paths), is_html=True)
    assert result["sent"] is False
    media.validate_content(text, prepared(paths), is_html=True)
    with pytest.raises(ValueError, match="final Teams HTML"):
        media.send_media(graph, "chat", text, prepared(paths))
    assert writes(graph) == []
