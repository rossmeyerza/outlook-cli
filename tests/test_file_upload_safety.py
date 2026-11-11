"""Upload failures must not expose credentials or trigger duplicate writes."""
from __future__ import annotations

import traceback
from unittest.mock import Mock

import httpx
import pytest

from outlook_draft.commands.files import _write_download
from outlook_draft.errors import OutlookAPIError
from outlook_draft.outlook_client import OutlookClient


def upload_client(response=None, failure=None):
    client = OutlookClient(Mock())
    transport = Mock()
    transport.put = Mock(return_value=response, side_effect=failure)
    client._ensure_client = Mock(return_value=transport)
    return client, transport


def test_upload_network_error_hides_the_session_credential():
    upload_url = "https://upload.example.test/?sig=TEST_CREDENTIAL"
    client, transport = upload_client(failure=httpx.ConnectError(upload_url))
    with pytest.raises(OutlookAPIError, match="uncertain") as error:
        client._upload_chunks(upload_url, b"data")
    assert transport.put.call_count == 1
    assert "TEST_CREDENTIAL" not in "".join(traceback.format_exception(error.value))


def test_upload_http_failure_hides_response_content():
    client, transport = upload_client(response=httpx.Response(503, text="TEST_CREDENTIAL"))
    with pytest.raises(OutlookAPIError, match="file chunk failed") as error:
        client._upload_chunks("https://upload.example.test/session", b"data")
    assert transport.put.call_count == 1
    assert "TEST_CREDENTIAL" not in str(error.value)


def test_incomplete_session_does_not_claim_success():
    client, transport = upload_client(response=httpx.Response(202, json={"nextExpectedRanges": ["0-"]}))
    with pytest.raises(OutlookAPIError, match="no final item ID"):
        client._upload_chunks("https://upload.example.test/session", b"data")
    assert transport.put.call_count == 1


def test_completed_session_returns_the_item():
    client, transport = upload_client(response=httpx.Response(201, json={"id": "remote-id", "name": "report.txt"}))
    assert client._upload_chunks("https://upload.example.test/session", b"data")["id"] == "remote-id"
    assert transport.put.call_count == 1


def test_download_never_truncates_an_existing_file_without_approval(tmp_path):
    target = tmp_path / "report.txt"
    target.write_bytes(b"Original")
    with pytest.raises(FileExistsError):
        _write_download(target, b"New", overwrite=False)
    assert target.read_bytes() == b"Original"
    _write_download(target, b"New", overwrite=True)
    assert target.read_bytes() == b"New"
