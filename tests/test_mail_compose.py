from __future__ import annotations

import argparse
import base64
from contextlib import contextmanager
from unittest.mock import Mock

import pytest

from outlook_draft import config
from outlook_draft.commands import drafts as cli
from outlook_draft.errors import CliError, OutlookAPIError
from outlook_draft.mail_compose import (
    MAX_ATTACHMENT_BYTES,
    MailComposeMixin,
    attach_to_draft,
    draft_update_fields,
    encode_attachment,
    encode_attachments,
    finish_quoted_draft,
)
from outlook_draft.mcp import drafts as mcp


class FakeClient(MailComposeMixin):
    def __init__(self):
        self.draft = {"Id": "d1", "IsDraft": True, "Body": {"ContentType": "Text", "Content": "original <text>"}}
        self.get_draft = Mock(side_effect=lambda _: self.draft)
        self.update_draft = Mock(side_effect=lambda *args, **kwargs: dict(self.draft))
        self.create_draft = Mock(side_effect=lambda **kwargs: dict(self.draft))
        self.create_reply_draft = Mock(side_effect=lambda *args, **kwargs: dict(self.draft))
        self._request = Mock(return_value=Mock(json=lambda: dict(self.draft)))
        self.close = Mock()
        self.send_mail = Mock()


@pytest.fixture
def client(monkeypatch, tmp_path):
    fake = FakeClient()
    signature = tmp_path / "signature.html"
    signature.write_text("<div>signature</div>")
    monkeypatch.setattr(config, "SIGNATURE_NEW_FILE", signature)
    monkeypatch.setattr(config, "SIGNATURE_REPLY_FILE", signature)
    monkeypatch.setattr(cli.runtime, "get_client", lambda: fake)
    monkeypatch.setattr(cli, "_resolve_message_id", lambda *args: "m1")
    monkeypatch.setattr(cli, "_resolve_draft_id", lambda *args: "d1")

    @contextmanager
    def factory():
        yield fake

    monkeypatch.setattr(mcp, "outlook_client", factory)
    return fake


@pytest.fixture
def attachment(tmp_path):
    path = tmp_path / "report.txt"
    path.write_bytes(b"hello\x00world")
    return path


def parse(*args):
    parser = argparse.ArgumentParser()
    cli.register(parser.add_subparsers(dest="group"))
    return parser.parse_args(args)


def test_encode_attachment(attachment):
    encoded = encode_attachment(attachment)
    assert encoded["@odata.type"] == "#Microsoft.OutlookServices.FileAttachment"
    assert encoded["Name"] == "report.txt"
    assert encoded["ContentType"] == "text/plain"
    assert base64.b64decode(encoded["ContentBytes"]) == b"hello\x00world"


@pytest.mark.parametrize("name", ["", " ", "missing", "."])
def test_reject_invalid_attachment(name, tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    with pytest.raises(ValueError):
        encode_attachment(name)


def test_reject_fifo_before_open(tmp_path):
    import os

    fifo = tmp_path / "pipe"
    os.mkfifo(fifo)
    with pytest.raises(ValueError, match="regular file"):
        encode_attachment(fifo)


def test_attachment_size_limit(tmp_path):
    path = tmp_path / "blob.unknown-extension"
    path.write_bytes(b"x" * MAX_ATTACHMENT_BYTES)
    assert encode_attachment(path)["ContentType"] == "application/octet-stream"
    with path.open("ab") as stream:
        stream.write(b"x")
    with pytest.raises(ValueError, match="3 MiB"):
        encode_attachment(path)


def test_empty_file_is_valid(tmp_path):
    path = tmp_path / "empty.txt"
    path.touch()
    assert encode_attachment(path)["ContentBytes"] == ""


def test_attachment_post_has_no_retry(client, attachment):
    client.add_draft_attachment("d1", attachment)
    args, kwargs = client._request.call_args
    assert args == ("POST", "/me/messages/d1/attachments")
    assert kwargs["max_retries"] == 0
    assert kwargs["json_body"] == encode_attachment(attachment)


def test_attachment_refuses_non_draft(client, attachment):
    client.draft["IsDraft"] = False
    with pytest.raises(ValueError, match="not a draft"):
        client.add_draft_attachment("d1", attachment)
    client._request.assert_not_called()


def test_forward_post_has_no_retry(client):
    assert client.create_forward_draft("m1")["Id"] == "d1"
    client._request.assert_called_once_with(
        "POST", "/me/messages/m1/createForward", json_body={}, max_retries=0,
    )


@pytest.mark.parametrize("is_draft", [False, None])
def test_update_refuses_non_drafts(client, is_draft):
    client.draft["IsDraft"] = is_draft
    with pytest.raises(ValueError, match="not a draft"):
        client.update_mail_draft("d1", subject="changed")
    client.update_draft.assert_not_called()


def test_update_omitted_fields_and_explicit_empty(client):
    client.update_mail_draft("d1", subject="", body="", to=[], cc=[], bcc=[])
    client.update_draft.assert_called_once_with("d1", subject="", body="", to=[], cc=[], bcc=[])
    client.update_draft.reset_mock()
    assert client.update_mail_draft("d1") == client.draft
    client.update_draft.assert_not_called()


def test_update_fields_preserve_omitted_body_and_signature(client):
    fields = draft_update_fields("new", None, None, [""], None, False, None, "reply")
    assert fields == {"subject": "new", "cc": []}


def test_update_fields_default_has_no_signature(client):
    fields = draft_update_fields(None, "", None, None, None, False, None, "none")
    assert fields == {"body": "", "content_type": "Text"}


def test_update_explicit_signature_has_no_duplicate(client):
    fields = draft_update_fields(None, "<p>hi</p><div>signature</div>", None, None, None, True, None, "reply")
    assert fields["body"].count("<div>signature</div>") == 1


def test_partial_attachment_error_retains_id_and_no_retry(client, attachment):
    client._request.side_effect = [Mock(json=lambda: {}), OutlookAPIError(503, "uncertain")]
    with pytest.raises(OutlookAPIError, match="Draft d1: attachment 2 failed after 1"):
        attach_to_draft(client, client.draft, encode_attachments([str(attachment)] * 2))
    assert client._request.call_count == 2


def test_quoted_update_error_includes_draft_id(client):
    client.update_draft.side_effect = OutlookAPIError(503, "uncertain")
    with pytest.raises(OutlookAPIError, match="Draft d1 remains saved"):
        finish_quoted_draft(client, client.draft, "hi", False, "sig")


@pytest.mark.parametrize("operation", ["create", "reply", "update", "forward"])
def test_mcp_validates_all_files_before_write(client, attachment, operation):
    options = {"attachments": [str(attachment), str(attachment) + ".missing"]}
    with pytest.raises(ValueError, match="Cannot read attachment"):
        if operation == "create":
            mcp.draft_create(["a@example.com"], "s", "b", **options)
        elif operation == "reply":
            mcp.draft_reply("m1", "b", **options)
        elif operation == "forward":
            mcp.draft_forward("m1", ["a@example.com"], **options)
        else:
            mcp.draft_update("d1", subject="s", **options)
    client.create_draft.assert_not_called()
    client.create_reply_draft.assert_not_called()
    client.update_draft.assert_not_called()
    client._request.assert_not_called()


def test_mcp_forward_quotes_original(client, attachment):
    result = mcp.draft_forward("m1", ["a@example.com,b@example.com"], "Intro", attachments=[str(attachment)])
    assert result["id"] == "d1"
    kwargs = client.update_draft.call_args.kwargs
    assert kwargs["to"] == ["a@example.com", "b@example.com"]
    assert kwargs["body"].endswith("<div>signature</div><p>original &lt;text&gt;</p>")
    assert kwargs["body"].count("<div>signature</div>") == 1
    assert client._request.call_count == 2


def test_mcp_update_recipients_only_preserves_body(client):
    mcp.draft_update("d1", to=[], subject="")
    client.update_draft.assert_called_once_with("d1", subject="", to=[])


@pytest.mark.parametrize("operation", ["create", "reply", "update", "forward"])
def test_cli_attach_is_repeatable_and_json_supported(operation):
    required = {"create": ["--to", "a@b", "-s", "s"], "reply": ["m1"],
                "update": ["d1"], "forward": ["m1", "--to", "a@b"]}
    args = parse("draft", operation, *required[operation], "--attach", "a", "--attach", "b", "--json")
    assert args.attach == ["a", "b"]
    assert args.json is True


def test_cli_update_empty_fields(client, capsys):
    args = parse("draft", "update", "d1", "--to", "", "--subject", "", "--body", "", "--json")
    args.func(args)
    client.update_draft.assert_called_once_with("d1", subject="", to=[], body="", content_type="Text")
    assert '"Id": "d1"' in capsys.readouterr().out


def test_cli_forward_json(client, capsys):
    args = parse("draft", "forward", "m1", "--to", "a@b", "--json")
    args.func(args)
    assert client.update_draft.call_args.kwargs["body"].endswith("<p>original &lt;text&gt;</p>")
    assert '"Id": "d1"' in capsys.readouterr().out


@pytest.mark.parametrize("operation", ["create", "reply", "update", "forward"])
def test_cli_attachment_invalid_before_any_write(client, operation):
    required = {"create": ["--to", "a@b", "-s", "s", "-b", "body"],
                "reply": ["m1", "-b", "body"], "update": ["d1"],
                "forward": ["m1", "--to", "a@b"]}
    args = parse("draft", operation, *required[operation], "--attach", "")
    with pytest.raises(CliError, match="empty"):
        args.func(args)
    client.create_draft.assert_not_called()
    client.create_reply_draft.assert_not_called()
    client.update_draft.assert_not_called()
    client._request.assert_not_called()


@pytest.mark.parametrize("attached", [False, True])
def test_mail_send_only_passes_attachments_when_present(client, attachment, attached):
    args = argparse.Namespace(body="text", body_file=None, to=["a@b"], cc=None, bcc=None,
                              subject="s", yes=True, signature="none", html=False,
                              no_save=False, importance="Normal", attach=[str(attachment)] if attached else None)
    cli.cmd_send_mail(args)
    kwargs = client.send_mail.call_args.kwargs
    assert ("attachments" in kwargs) == attached
    if attached:
        assert kwargs["attachments"] == [encode_attachment(attachment)]


@pytest.mark.parametrize("operation", ["create", "reply", "update"])
def test_mcp_adds_attachments(client, attachment, operation):
    options = {"attachments": [str(attachment)]}
    if operation == "create":
        result = mcp.draft_create(["a@b"], "s", "body", **options)
    elif operation == "reply":
        result = mcp.draft_reply("m1", "body", **options)
    else:
        result = mcp.draft_update("d1", subject="new", **options)
    assert result["id"] == "d1"
    assert client._request.call_args.args == ("POST", "/me/messages/d1/attachments")


@pytest.mark.parametrize("operation", ["create", "reply"])
def test_cli_create_reply_json_and_attachment(client, attachment, capsys, operation):
    required = ["--to", "a@b", "-s", "s"] if operation == "create" else ["m1"]
    args = parse("draft", operation, *required, "-b", "body", "--attach", str(attachment), "--json")
    args.func(args)
    assert '"Id": "d1"' in capsys.readouterr().out
    assert client._request.call_args.args == ("POST", "/me/messages/d1/attachments")


def test_send_rejects_all_invalid_files_before_write(client, attachment):
    args = argparse.Namespace(body="text", body_file=None, to=["a@b"], subject="s", yes=True,
                              attach=[str(attachment), str(attachment) + ".missing"])
    with pytest.raises(CliError, match="Cannot read attachment"):
        cli.cmd_send_mail(args)
    client.send_mail.assert_not_called()


def test_mail_send_preview_does_not_send(client):
    args = argparse.Namespace(body="text", body_file=None, to=["a@b"], subject="s", yes=False)
    with pytest.raises(SystemExit) as exc:
        cli.cmd_send_mail(args)
    assert exc.value.code == 2
    client.send_mail.assert_not_called()
