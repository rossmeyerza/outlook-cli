from __future__ import annotations

from unittest.mock import Mock

import pytest
from mcp_fakes import FakeClient, use_client
from test_cli_smoke import capture_parser
from test_outlook_client_mail import RecordingClient

from outlook_draft.commands import mail as command
from outlook_draft.errors import CliError, OutlookAPIError
from outlook_draft.mcp import mail


@pytest.mark.parametrize("parent, path", [
    (None, "/me/mailfolders"),
    ("inbox", "/me/mailfolders/inbox/childfolders"),
    ("folder/a+b=", "/me/mailfolders/folder%2Fa%2Bb%3D/childfolders"),
])
def test_create_folder_request(parent, path):
    client = RecordingClient()
    try:
        client.create_mail_folder(" Clients ", parent_id=parent)
        assert client.calls[0]["method"] == "POST"
        assert client.calls[0]["path"] == path
        assert client.calls[0]["json_body"] == {"DisplayName": "Clients"}
        assert client.calls[0]["max_retries"] == 0
    finally:
        client.close()


@pytest.mark.parametrize("name, parent", [("", None), ("  ", None), ("a\nb", None), ("a\x7fb", None), ("Clients", " ")])
def test_create_folder_rejects_invalid_input(name, parent):
    client = RecordingClient()
    try:
        with pytest.raises(ValueError):
            client.create_mail_folder(name, parent_id=parent)
        assert client.calls == []
    finally:
        client.close()


def test_mcp_create_folder(monkeypatch):
    fake = use_client(monkeypatch, mail, "outlook_client", FakeClient(
        create_mail_folder={"Id": "f1", "DisplayName": "Clients"},
    ))
    assert mail.mail_create_folder("Clients", "inbox") == {"id": "f1", "displayName": "Clients"}
    assert fake.call("create_mail_folder") == (("Clients",), {"parent_id": "inbox"})


def test_cli_create_folder(monkeypatch, capsys):
    parser = capture_parser(monkeypatch)
    args = parser.parse_args(["mail", "create-folder", "Clients", "--parent", "inbox", "--json"])
    client = Mock()
    client.create_mail_folder.return_value = {"Id": "f1", "DisplayName": "Clients"}
    monkeypatch.setattr(command.runtime, "get_client", lambda: client)
    args.func(args)
    client.create_mail_folder.assert_called_once_with("Clients", parent_id="inbox")
    client.close.assert_called_once()
    assert '"id": "f1"' in capsys.readouterr().out


@pytest.mark.parametrize("error", [ValueError("Invalid name"), OutlookAPIError(409, "Exists")])
def test_cli_create_folder_error_closes_client(monkeypatch, error):
    parser = capture_parser(monkeypatch)
    args = parser.parse_args(["mail", "create-folder", "Clients"])
    client = Mock()
    client.create_mail_folder.side_effect = error
    monkeypatch.setattr(command.runtime, "get_client", lambda: client)
    with pytest.raises(CliError, match="Failed to create folder"):
        args.func(args)
    client.close.assert_called_once()
