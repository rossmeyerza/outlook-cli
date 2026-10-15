"""Outbound messages must not be duplicated by transport retries."""
from unittest.mock import Mock

import httpx
import pytest

from outlook_draft.errors import OutlookAPIError
from outlook_draft.outlook_client import OutlookClient


@pytest.mark.parametrize("send", [
    lambda client: client.send_message("draft-1"),
    lambda client: client.send_mail(subject="s", body="b", to=["a@example.com"]),
    lambda client: client.send_teams_message("chat-1", "hello"),
])
@pytest.mark.parametrize("failure", [429, 503, "timeout"])
def test_send_transport_does_not_retry(monkeypatch, send, failure):
    manager = Mock()
    manager.get_token.return_value = "fake"
    client = OutlookClient(manager)
    transport = Mock()
    if failure == "timeout":
        transport.request.side_effect = httpx.ReadTimeout("uncertain result")
    else:
        transport.request.return_value = httpx.Response(failure, text="send failed")
    monkeypatch.setattr(client, "_ensure_client", lambda: transport)
    sleep = Mock()
    monkeypatch.setattr("outlook_draft.outlook_client.time.sleep", sleep)
    with pytest.raises(OutlookAPIError):
        send(client)
    assert transport.request.call_count == 1
    sleep.assert_not_called()
