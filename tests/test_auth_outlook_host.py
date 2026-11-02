"""Offline checks for Outlook Web login and token capture."""

import base64
import json
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from outlook_draft import auth


def jwt(audience, scopes="Mail.Read"):
    payload = base64.urlsafe_b64encode(json.dumps({"aud": audience, "scp": scopes}).encode()).decode().rstrip("=")
    return f"header.{payload}.signature"


@pytest.fixture
def tokens_file(tmp_path, monkeypatch):
    path = tmp_path / "tokens.json"
    monkeypatch.setattr(auth.config, "TOKENS_FILE", path)
    monkeypatch.setattr(auth.config, "ensure_dirs", lambda: None)
    return path


@pytest.mark.parametrize("url, expected", [
    ("https://outlook.office.com/mail/", True),
    ("https://outlook.cloud.microsoft/mail/", True),
    ("https://OUTLOOK.CLOUD.MICROSOFT:443/mail/", True),
    ("https://login.microsoftonline.com/?redirect_uri=https://outlook.cloud.microsoft/mail/", False),
    ("https://outlook.cloud.microsoft.example.test/", False),
    ("https://outlook.cloud.microsoft@example.test/", False),
    ("https://[invalid", False),
    ("not a url", False),
])
def test_on_outlook(url, expected):
    assert auth._on_outlook(url) is expected


def capture(captured, url, token):
    request = SimpleNamespace(url=url, headers={"authorization": f"Bearer {token}"})
    auth._make_request_interceptor(captured)(request)


@pytest.mark.parametrize("audience", [
    "https://outlook.office.com",
    "https://outlook.office365.com/",
    ["https://outlook.office.com"],
])
def test_new_host_maps_outlook_token_and_saves_before_login_finishes(tokens_file, audience):
    tokens_file.write_text(json.dumps({"tokens": {"graph.microsoft.com": "saved-graph", "substrate.office.com/sydney": "saved-copilot"}}))
    captured = {}
    token = jwt(audience)
    capture(captured, "https://outlook.cloud.microsoft/owa/service.svc", token)
    assert captured == {"outlook.office.com": token}
    assert json.loads(tokens_file.read_text())["tokens"] == {
        "graph.microsoft.com": "saved-graph",
        "substrate.office.com/sydney": "saved-copilot",
        "outlook.office.com": token,
    }


@pytest.mark.parametrize("token", [
    "unreadable",
    jwt("https://graph.microsoft.com"),
    jwt("https://outlook.office.com.example.test"),
    jwt(auth.COPILOT_AUDIENCE),
    jwt(None),
])
def test_new_host_rejects_other_audiences(tokens_file, token):
    captured = {}
    capture(captured, "https://outlook.cloud.microsoft/owa/service.svc", token)
    assert captured == {}
    assert not tokens_file.exists()


def test_request_host_does_not_match_redirect_parameter(tokens_file):
    captured = {}
    capture(captured, "https://example.test/?redirect=https://outlook.office.com", jwt("https://graph.microsoft.com"))
    assert captured == {}
    assert not tokens_file.exists()


@pytest.mark.parametrize("audience, accepted", [("https://outlook.office.com", True), (auth.COPILOT_AUDIENCE, False)])
def test_storage_fallback_uses_same_token_mapping(audience, accepted):
    token = jwt(audience)
    page = Mock()
    page.evaluate.return_value = [{"target": "https://outlook.cloud.microsoft/Mail.Read", "secret": token}]
    captured = {}
    auth._extract_storage_tokens(page, captured)
    assert captured == ({"outlook.office.com": token} if accepted else {})


def test_scope_upgrade_is_saved(tokens_file):
    captured = {}
    capture(captured, "https://outlook.office.com/api/v2.0/me", jwt("https://outlook.office.com"))
    broader = jwt("https://outlook.office.com", "Mail.Read Mail.Write")
    capture(captured, "https://outlook.cloud.microsoft/owa/service.svc", broader)
    assert json.loads(tokens_file.read_text())["tokens"]["outlook.office.com"] == broader


def test_save_failure_preserves_previous_cache(tokens_file, monkeypatch, capsys):
    auth._save_tokens({"graph.microsoft.com": "saved-graph"})
    before = tokens_file.read_bytes()
    monkeypatch.setattr(auth.os, "replace", Mock(side_effect=OSError("write failed")))
    captured = {}
    capture(captured, "https://outlook.cloud.microsoft/owa/service.svc", jwt("https://outlook.office.com"))
    assert "outlook.office.com" in captured
    assert tokens_file.read_bytes() == before
    assert list(tokens_file.parent.iterdir()) == [tokens_file]
    output = capsys.readouterr()
    assert output.out == ""
    assert "Could not save tokens" in output.err


def test_mfa_wait_accepts_redirect_to_new_host(monkeypatch):
    page = Mock()
    monkeypatch.setattr(auth, "_has_okta_password_error", lambda page: False)
    monkeypatch.setattr(auth, "_find_mfa_number", lambda page: "12")

    def redirect(predicate, timeout):
        assert timeout == 180000
        assert predicate("https://outlook.cloud.microsoft/mail/")
        assert not predicate("https://login.microsoftonline.com/?redirect_uri=https://outlook.cloud.microsoft/")

    page.wait_for_url.side_effect = redirect
    auth._wait_for_mfa(page)
    page.wait_for_url.assert_called_once()


def test_mfa_wait_returns_when_already_on_new_host(monkeypatch):
    page = Mock(url="https://outlook.cloud.microsoft/mail/")
    monkeypatch.setattr(auth, "_has_okta_password_error", lambda page: False)
    monkeypatch.setattr(auth, "_find_mfa_number", lambda page: None)
    auth._wait_for_mfa(page)
    page.wait_for_url.assert_not_called()


@pytest.mark.parametrize("close_fails", [False, True])
def test_login_saves_outlook_before_copilot_and_keeps_other_tokens(tokens_file, monkeypatch, close_fails):
    auth._save_tokens({"graph.microsoft.com": "saved-graph"})
    outlook_token = jwt("https://outlook.office.com")
    copilot_token = jwt(auth.COPILOT_AUDIENCE)
    playwright = Mock()
    browser = playwright.chromium.launch.return_value
    context = browser.new_context.return_value
    page = context.new_page.return_value

    def register_request(event, handler):
        assert event == "request"
        page.goto.side_effect = lambda *args, **kwargs: handler(SimpleNamespace(
            url="https://outlook.cloud.microsoft/owa/service.svc",
            headers={"authorization": f"Bearer {outlook_token}"},
        ))

    def capture_copilot(context):
        assert auth._load_tokens()["outlook.office.com"] == outlook_token
        return copilot_token

    page.on.side_effect = register_request
    manager = Mock()
    manager.__enter__ = Mock(return_value=playwright)
    manager.__exit__ = Mock(return_value=False)
    monkeypatch.setattr(auth, "sync_playwright", lambda: manager)
    monkeypatch.setattr(auth, "_wait_for_login_or_tokens", lambda *args: False)
    monkeypatch.setattr(auth, "_capture_copilot_in_context", capture_copilot)
    if close_fails:
        browser.close.side_effect = RuntimeError("browser stopped")
        with pytest.raises(RuntimeError, match="browser stopped"):
            auth.capture_tokens_via_browser()
    else:
        result = auth.capture_tokens_via_browser()
        assert result == {"outlook.office.com": outlook_token, auth.config.COPILOT_TOKEN_DOMAIN: copilot_token}

    page.goto.assert_called_once_with("https://outlook.cloud.microsoft/calendar", wait_until="domcontentloaded")
    assert auth._load_tokens() == {
        "graph.microsoft.com": "saved-graph",
        "outlook.office.com": outlook_token,
        auth.config.COPILOT_TOKEN_DOMAIN: copilot_token,
    }
