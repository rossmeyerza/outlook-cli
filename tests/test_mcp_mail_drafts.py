from __future__ import annotations

import base64
from pathlib import Path

import pytest
from mcp_fakes import FakeClient, use_client

from outlook_draft import config
from outlook_draft.mcp import drafts, mail


def _message(**overrides):
    base = {
        "Id": "AAMk-1",
        "Subject": "Quarterly plan",
        "From": {"EmailAddress": {"Name": "Ann", "Address": "ann@example.com"}},
        "ToRecipients": [{"EmailAddress": {"Name": "Ross", "Address": "ross@example.com"}}],
        "ReceivedDateTime": "2026-10-15T08:00:00Z",
        "IsRead": False,
        "HasAttachments": False,
        "BodyPreview": "x" * 500,
        "Body": {"ContentType": "HTML", "Content": "<p>Hello <b>there</b></p><a href=\"https://a.sharepoint.com/f\">deck</a>"},
    }
    base.update(overrides)
    return base


def test_mail_unread_returns_lean_summaries_with_full_ids(monkeypatch):
    fake = use_client(monkeypatch, mail, "outlook_client", FakeClient(list_unread=[_message()]))

    result = mail.mail_unread(count=3)

    assert fake.call("list_unread")[1] == {"top": 3}
    assert result == [
        {
            "id": "AAMk-1",
            "subject": "Quarterly plan",
            "from": {"name": "Ann", "address": "ann@example.com"},
            "received": "2026-10-15T08:00:00Z",
            "isRead": False,
            "hasAttachments": False,
            "preview": "x" * 300,
        }
    ]


def test_mail_unread_caps_count(monkeypatch):
    fake = use_client(monkeypatch, mail, "outlook_client", FakeClient(list_unread=[]))

    mail.mail_unread(count=100000)

    assert fake.call("list_unread")[1] == {"top": 100}


def test_mail_search_passes_query(monkeypatch):
    fake = use_client(monkeypatch, mail, "outlook_client", FakeClient(search_messages=[_message()]))

    mail.mail_search("budget", count=5)

    assert fake.call("search_messages") == (("budget",), {"top": 5})


def test_mail_read_converts_html_and_lists_links(monkeypatch):
    fake = FakeClient(
        get_message=_message(HasAttachments=True),
        list_message_attachments=[{"Id": "att-1", "Name": "a.pdf", "ContentType": "application/pdf", "Size": 10, "IsInline": False}],
    )
    use_client(monkeypatch, mail, "outlook_client", fake)

    result = mail.mail_read("AAMk-1")

    assert result["body"] == "Hello there\ndeck"
    assert "bodyHtml" not in result
    assert result["to"] == [{"name": "Ross", "address": "ross@example.com"}]
    assert result["attachments"][0]["name"] == "a.pdf"
    assert result["links"][0]["url"] == "https://a.sharepoint.com/f"


def test_mail_read_include_html_and_truncation(monkeypatch):
    use_client(monkeypatch, mail, "outlook_client", FakeClient(get_message=_message()))

    result = mail.mail_read("AAMk-1", include_html=True, max_chars=5)

    assert result["bodyHtml"].startswith("<p>Hello")
    assert result["body"] == "Hello"
    assert result["bodyTruncated"] is True


def test_mail_links_defaults_to_share_links_only(monkeypatch):
    body = '<a href="https://example.com/x">x</a><a href="https://a.sharepoint.com/f">f</a>'
    use_client(monkeypatch, mail, "outlook_client", FakeClient(get_message=_message(Body={"ContentType": "HTML", "Content": body})))

    assert [link["url"] for link in mail.mail_links("AAMk-1")] == ["https://a.sharepoint.com/f"]


def test_mail_download_attachments_writes_files(monkeypatch, tmp_path: Path):
    content = base64.b64encode(b"data").decode()
    fake = FakeClient(
        list_message_attachments=[
            {"Id": "1", "IsInline": False},
            {"Id": "2", "IsInline": True},
        ],
        get_message_attachment=lambda mid, aid: {"Name": f"file{aid}.txt", "ContentBytes": content},
    )
    use_client(monkeypatch, mail, "outlook_client", fake)

    written = mail.mail_download_attachments("AAMk-1", str(tmp_path / "out"))

    assert written == [str(tmp_path / "out" / "file1.txt")]
    assert (tmp_path / "out" / "file1.txt").read_bytes() == b"data"


def test_mail_download_attachments_requires_absolute_dir(monkeypatch):
    use_client(monkeypatch, mail, "outlook_client", FakeClient())

    with pytest.raises(ValueError, match="absolute"):
        mail.mail_download_attachments("AAMk-1", "relative/dir")


def test_mail_move_resolves_folder_name(monkeypatch):
    fake = FakeClient(
        list_mail_folders=[{"Id": "folder-9", "DisplayName": "Clients"}],
        move_message={"Id": "AAMk-2", "Subject": "s"},
    )
    use_client(monkeypatch, mail, "outlook_client", fake)

    result = mail.mail_move("AAMk-1", "clients")

    assert fake.call("move_message")[0] == ("AAMk-1", "folder-9")
    assert result["id"] == "AAMk-2"


def test_mail_move_unknown_folder_raises(monkeypatch):
    use_client(monkeypatch, mail, "outlook_client", FakeClient(list_mail_folders=[]))

    with pytest.raises(ValueError, match="No folder matching"):
        mail.mail_move("AAMk-1", "nope")


def test_mail_mark_read_and_archive(monkeypatch):
    fake = FakeClient(archive_message={"Id": "AAMk-3", "Subject": "s"})
    use_client(monkeypatch, mail, "outlook_client", fake)

    assert mail.mail_mark_read("AAMk-1", is_read=False) == {"id": "AAMk-1", "isRead": False}
    assert fake.call("update_message_read_state") == (("AAMk-1",), {"is_read": False})
    assert mail.mail_archive("AAMk-1")["id"] == "AAMk-3"


@pytest.fixture
def signatures(tmp_path: Path, monkeypatch):
    new = tmp_path / "new.html"
    reply = tmp_path / "reply.html"
    new.write_text("<div>NEW-SIG</div>")
    reply.write_text("<div>REPLY-SIG</div>")
    monkeypatch.setattr(config, "SIGNATURE_NEW_FILE", new)
    monkeypatch.setattr(config, "SIGNATURE_REPLY_FILE", reply)


def test_draft_create_applies_aptos_and_new_signature(monkeypatch, signatures):
    fake = FakeClient(create_draft={"Id": "d1", "Subject": "Hi", "WebLink": "https://o/d1", "ToRecipients": []})
    use_client(monkeypatch, drafts, "outlook_client", fake)

    result = drafts.draft_create(
        to=["a@example.com, b@example.com"], subject="Hi", body="Line one\n\nLine two", cc=["c@example.com"]
    )

    kwargs = fake.call("create_draft")[1]
    assert kwargs["to"] == ["a@example.com", "b@example.com"]
    assert kwargs["cc"] == ["c@example.com"]
    assert kwargs["bcc"] is None
    assert kwargs["content_type"] == "HTML"
    assert "Aptos" in kwargs["body"]
    assert "<p>Line one</p><p>Line two</p>" in kwargs["body"]
    assert kwargs["body"].endswith("<div>NEW-SIG</div>")
    assert result["id"] == "d1"
    assert result["webLink"] == "https://o/d1"
    assert "send_message" not in fake.names()


def test_draft_create_html_body_is_not_escaped(monkeypatch, signatures):
    fake = FakeClient(create_draft={"Id": "d1"})
    use_client(monkeypatch, drafts, "outlook_client", fake)

    drafts.draft_create(to=["a@example.com"], subject="s", body="<p>raw</p>", html=True)

    assert "<p>raw</p>" in fake.call("create_draft")[1]["body"]


def test_draft_create_validates_input(monkeypatch, signatures):
    use_client(monkeypatch, drafts, "outlook_client", FakeClient())

    with pytest.raises(ValueError, match="recipient"):
        drafts.draft_create(to=[], subject="s", body="b")
    with pytest.raises(ValueError, match="importance"):
        drafts.draft_create(to=["a@example.com"], subject="s", body="b", importance="Urgent")


def test_draft_create_missing_signature_raises_not_exits(monkeypatch, tmp_path):
    monkeypatch.setattr(config, "SIGNATURE_NEW_FILE", tmp_path / "missing.html")
    use_client(monkeypatch, drafts, "outlook_client", FakeClient())

    with pytest.raises(FileNotFoundError):
        drafts.draft_create(to=["a@example.com"], subject="s", body="b")


def test_draft_reply_uses_reply_signature_and_quotes_original(monkeypatch, signatures):
    fake = FakeClient(
        create_reply_draft={"Id": "d2", "Body": {"ContentType": "Text", "Content": "orig <text>"}},
        update_draft={"Id": "d2", "Subject": "RE: x"},
    )
    use_client(monkeypatch, drafts, "outlook_client", fake)

    result = drafts.draft_reply("AAMk-1", "Thanks", reply_all=True)

    assert fake.call("create_reply_draft") == (("AAMk-1",), {"reply_all": True})
    body = fake.call("update_draft")[1]["body"]
    assert "REPLY-SIG" in body and "NEW-SIG" not in body
    assert body.endswith("<p>orig &lt;text&gt;</p>")
    assert result["id"] == "d2"


def test_draft_delete_refuses_non_drafts(monkeypatch):
    fake = FakeClient(get_draft={"Id": "m1", "IsDraft": False})
    use_client(monkeypatch, drafts, "outlook_client", fake)

    with pytest.raises(ValueError, match="not a draft"):
        drafts.draft_delete("m1")
    assert "delete_draft" not in fake.names()


def test_draft_delete_and_get(monkeypatch):
    fake = FakeClient(get_draft={"Id": "d1", "IsDraft": True, "Subject": "s", "Body": {"ContentType": "HTML", "Content": "<p>hi</p>"}})
    use_client(monkeypatch, drafts, "outlook_client", fake)

    assert drafts.draft_get("d1")["body"] == "hi"
    assert drafts.draft_delete("d1") == {"id": "d1", "deleted": True}
    assert fake.call("delete_draft")[0] == ("d1",)


def test_draft_list_maps_count(monkeypatch):
    fake = use_client(monkeypatch, drafts, "outlook_client", FakeClient(list_drafts=[{"Id": "d1", "Subject": "s"}]))

    assert drafts.draft_list(count=2)[0]["id"] == "d1"
    assert fake.call("list_drafts")[1] == {"top": 2}
