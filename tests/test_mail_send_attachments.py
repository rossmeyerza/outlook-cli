from __future__ import annotations

import base64
from unittest.mock import Mock

import pytest
from mcp_fakes import FakeClient, use_client
from test_outlook_client_mail import RecordingClient

from outlook_draft.cli import build_parser
from outlook_draft.commands import drafts
from outlook_draft.errors import CliError
from outlook_draft.mail_compose import encode_attachments
from outlook_draft.mcp import sending


@pytest.mark.parametrize("confirm", [False, True])
def test_mcp_send_attachment_preview_and_payload(monkeypatch, tmp_path, confirm):
    path = tmp_path / "file.txt"
    path.write_bytes(b"ab")
    monkeypatch.setattr(sending, "load_signature", lambda path: "<p>Signature</p>")
    fake = FakeClient(send_mail=None)
    if confirm:
        use_client(monkeypatch, sending, "outlook_client", fake)
    else:
        monkeypatch.setattr(sending, "outlook_client", Mock(side_effect=AssertionError("No authentication")))
    report = sending.mail_send(["a@example.test"], "Subject", "Body", attachments=[str(path)], confirm=confirm)
    assert report["sent"] is confirm
    assert report["attachments"] == [{"name": "file.txt", "size": 2}]
    assert "ContentBytes" not in str(report)
    if confirm:
        payload = fake.call("send_mail")[1]["attachments"]
        assert base64.b64decode(payload[0]["ContentBytes"]) == b"ab"


def test_client_send_mail_includes_encoded_attachments(tmp_path):
    path = tmp_path / "file.txt"
    path.write_text("test")
    client = RecordingClient()
    try:
        client.send_mail(subject="S", body="B", to=["a@example.test"], attachments=encode_attachments([str(path)]))
        call = client.calls[0]
        assert call["path"] == "/me/sendmail" and call["max_retries"] == 0
        assert call["json_body"]["Message"]["Attachments"][0]["Name"] == "file.txt"
    finally:
        client.close()


def test_mcp_send_invalid_attachment_does_not_authenticate(monkeypatch):
    monkeypatch.setattr(sending, "outlook_client", Mock(side_effect=AssertionError("No authentication")))
    with pytest.raises(ValueError):
        sending.mail_send(["a@example.test"], "Subject", "Body", attachments=[""], confirm=True)


@pytest.mark.parametrize("is_draft", [False, True])
def test_cli_draft_delete_refuses_received_mail(monkeypatch, is_draft):
    fake = Mock()
    fake.get_draft.return_value = {"IsDraft": is_draft}
    monkeypatch.setattr(drafts.runtime, "get_client", lambda: fake)
    args = build_parser().parse_args(["draft", "delete", "x" * 50])
    if is_draft:
        args.func(args)
        fake.delete_draft.assert_called_once_with("x" * 50)
    else:
        with pytest.raises(CliError, match="not a draft"):
            args.func(args)
        fake.delete_draft.assert_not_called()
    fake.close.assert_called_once()
