"""Retry contracts shared by Outlook and file operations."""
from unittest.mock import Mock

import httpx
import pytest

from outlook_draft.commands.files import _GraphClient
from outlook_draft.errors import OutlookAPIError
from outlook_draft.outlook_client import OutlookClient


@pytest.mark.parametrize("factory", [OutlookClient, _GraphClient])
@pytest.mark.parametrize("status", [429, 500, 503])
def test_retry_after_is_preserved(monkeypatch, factory, status):
    manager = Mock()
    manager.get_token.return_value = "fake"
    client = factory(manager)
    responses = [httpx.Response(status, headers={"Retry-After": "0.25"}), httpx.Response(200)]
    transport = Mock()
    transport.request.side_effect = responses
    monkeypatch.setattr(client, "_ensure_client", lambda: transport)
    sleep = Mock()
    monkeypatch.setattr("outlook_draft.outlook_client.time.sleep", sleep)
    assert client._request("GET", "https://example.test") is responses[1]
    sleep.assert_called_once_with(0.25)


@pytest.mark.parametrize("factory", [OutlookClient, _GraphClient])
def test_401_reauth_retry_and_failure(monkeypatch, factory):
    manager = Mock()
    manager.get_token.return_value = "fake"
    manager.run_reauth.return_value = True
    client = factory(manager)
    transport = Mock()
    transport.request.side_effect = [httpx.Response(401), httpx.Response(200)]
    monkeypatch.setattr(client, "_ensure_client", lambda: transport)
    assert client._request("GET", "https://example.test").status_code == 200
    manager.run_reauth.assert_called_once_with(headless=True)
    transport.request.side_effect = None
    transport.request.return_value = httpx.Response(401)
    manager.run_reauth.return_value = False
    with pytest.raises(OutlookAPIError, match="Authentication failed"):
        client._request("GET", "https://example.test")


@pytest.mark.parametrize("factory", [OutlookClient, _GraphClient])
def test_network_failure_retries_twice(monkeypatch, factory):
    client = factory(Mock())
    transport = Mock()
    transport.request.side_effect = httpx.ConnectError("offline")
    monkeypatch.setattr(client, "_ensure_client", lambda: transport)
    sleep = Mock()
    monkeypatch.setattr("outlook_draft.outlook_client.time.sleep", sleep)
    with pytest.raises(OutlookAPIError, match="offline"):
        client._request("GET", "https://example.test")
    assert transport.request.call_count == 3
    assert [call.args[0] for call in sleep.call_args_list] == [1, 2]
