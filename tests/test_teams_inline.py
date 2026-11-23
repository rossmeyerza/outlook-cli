import json
from argparse import Namespace
from contextlib import nullcontext
from unittest.mock import Mock

import pytest

from outlook_draft.commands import teams as cli
from outlook_draft.errors import CliError, OutlookAPIError
from outlook_draft.mcp import teams as mcp
from outlook_draft.outlook_client import OutlookClient
from outlook_draft.teams_inline import inline_image_ids

CHAT = '48:notes'
MESSAGE = '12345'
IMAGE = 'a+b/c='
PNG = b'\x89PNG\r\n\x1a\nexample'


def message():
    prefix = 'https://graph.microsoft.com/v1.0/chats/48%3Anotes/messages/12345/hostedContents/'
    return {'id': MESSAGE, 'body': {'contentType': 'html', 'content': (
        f'<img src="{prefix}a%2Bb%2Fc%3D/$value">'
        '<img src="../hostedContents/a%2Bb%2Fc%3D/$value">'
        '<img src="https://example.test/hostedContents/foreign/$value">'
        '<img src="https://graph.microsoft.com/v1.0/chats/other/messages/12345/hostedContents/foreign/$value">'
    )}}


def test_discovers_only_hosted_images_for_the_message():
    assert inline_image_ids(message(), CHAT) == [IMAGE]
    assert inline_image_ids(dict(message(), deletedDateTime='now'), CHAT) == []


def test_mcp_messages_keep_attachments_separate(monkeypatch):
    client = Mock()
    client.list_teams_messages.return_value = [message()]
    monkeypatch.setattr(mcp, 'graph_client', lambda: nullcontext(client))
    result = mcp.teams_messages(CHAT)['messages'][0]
    assert result['inlineImageIds'] == [IMAGE]
    assert result['hasAttachments'] is False


def test_cli_gets_exact_message_with_full_image_ids(monkeypatch, capsys):
    client = Mock()
    client.get_teams_message.return_value = message()
    monkeypatch.setattr(cli.runtime, 'get_graph_client', lambda: client)
    cli.cmd_teams_messages(Namespace(chat_id=CHAT, message_id=MESSAGE, json=True, count=20, full=False))
    client.get_teams_message.assert_called_once_with(CHAT, MESSAGE)
    client.list_teams_messages.assert_not_called()
    assert json.loads(capsys.readouterr().out)[0]['inlineImageIds'] == [IMAGE]


def test_download_uses_fixed_graph_get_and_escapes_ids():
    client = OutlookClient(Mock())
    client._request = Mock(return_value=Mock(content=PNG))
    assert client.get_teams_inline_image(CHAT, MESSAGE, IMAGE) == PNG
    client._request.assert_called_once_with('GET', '/chats/48%3Anotes/messages/12345/hostedContents/a%2Bb%2Fc%3D/$value')
    with pytest.raises(CliError, match='Invalid Teams ID'):
        client.get_teams_inline_image(CHAT, MESSAGE, '..')


def test_download_refuses_overwrite_unless_requested(monkeypatch, tmp_path):
    client = Mock()
    client.get_teams_inline_image.return_value = PNG
    monkeypatch.setattr(mcp, 'graph_client', lambda: nullcontext(client))
    path = tmp_path / 'image.png'
    path.write_bytes(b'original')
    with pytest.raises(CliError, match='already exists'):
        mcp.teams_download_inline_image(CHAT, MESSAGE, IMAGE, str(path))
    client.get_teams_inline_image.assert_not_called()
    assert path.read_bytes() == b'original'
    assert mcp.teams_download_inline_image(CHAT, MESSAGE, IMAGE, str(path), overwrite=True) == str(path)
    assert path.read_bytes() == PNG


@pytest.mark.parametrize('failure', [None, OutlookAPIError(404, 'Image not found')])
def test_download_does_not_create_file_on_failure(tmp_path, failure):
    client = Mock()
    client.get_teams_inline_image.return_value = b'<html>not an image</html>'
    client.get_teams_inline_image.side_effect = failure
    path = tmp_path / 'image.png'
    with pytest.raises(OutlookAPIError if failure else CliError):
        cli.download_inline_image(client, CHAT, MESSAGE, IMAGE, path, overwrite=False)
    assert not path.exists()
