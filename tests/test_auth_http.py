"""Auth HTTP probes are offline and never open a browser."""
from unittest.mock import Mock

import httpx
import pytest

from outlook_draft import auth


@pytest.mark.parametrize("status, expected", [(200, True), (401, False), (500, False)])
def test_token_probe_uses_httpx(monkeypatch, status, expected):
    get = Mock(return_value=httpx.Response(status))
    monkeypatch.setattr(auth.httpx, "get", get)
    assert auth._test_token("fake", "https://example.test") is expected
    get.assert_called_once_with(
        "https://example.test", headers={"Authorization": "Bearer fake"},
        timeout=10, follow_redirects=True,
    )


def test_token_probe_network_failure_is_best_effort(monkeypatch):
    monkeypatch.setattr(auth.httpx, "get", Mock(side_effect=httpx.ConnectError("offline")))
    assert auth._test_token("fake", "https://example.test") is False
