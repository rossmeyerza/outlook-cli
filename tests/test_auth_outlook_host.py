from __future__ import annotations

import base64
import json

import pytest

from outlook_draft import auth


def _jwt(claims: dict) -> str:
    def enc(obj: dict) -> str:
        return base64.urlsafe_b64encode(json.dumps(obj).encode()).decode().rstrip("=")

    return f"{enc({'alg': 'none'})}.{enc(claims)}.sig"


class _Request:
    def __init__(self, url: str, token: str) -> None:
        self.url = url
        self.headers = {"authorization": f"Bearer {token}"}


@pytest.fixture
def tokens_file(tmp_path, monkeypatch):
    path = tmp_path / "tokens.json"
    monkeypatch.setattr(auth.config, "TOKENS_FILE", path)
    monkeypatch.setattr(auth.config, "ensure_dirs", lambda: None)
    return path


@pytest.mark.parametrize(
    ("url", "expected"),
    [
        ("https://outlook.office.com/mail/", True),
        ("https://outlook.cloud.microsoft/mail/", True),
        ("https://login.microsoftonline.com/common/oauth2?redirect_uri=https://outlook.cloud.microsoft/mail/", False),
        ("https://login.okta.com/signin?next=outlook.office.com", False),
        ("not a url", False),
        (None, False),
    ],
)
def test_on_outlook_matches_host_only(url, expected) -> None:
    assert auth._on_outlook(url) is expected


def test_token_aud_reads_string_and_list() -> None:
    assert auth._token_aud(_jwt({"aud": "https://outlook.office.com"})) == "https://outlook.office.com"
    assert auth._token_aud(_jwt({"aud": ["a", "b"]})) == "a b"
    assert auth._token_aud("garbage") == ""


def test_new_host_outlook_token_stored_under_old_key(tokens_file) -> None:
    captured: dict[str, str] = {}
    token = _jwt({"aud": "https://outlook.office.com", "scp": "Mail.ReadWrite"})

    auth._make_request_interceptor(captured)(
        _Request("https://outlook.cloud.microsoft/owa/service.svc", token)
    )

    assert captured == {"outlook.office.com": token}


def test_new_host_non_outlook_token_ignored(tokens_file) -> None:
    captured: dict[str, str] = {}
    token = _jwt({"aud": "https://graph.microsoft.com", "scp": "User.Read"})

    auth._make_request_interceptor(captured)(
        _Request("https://outlook.cloud.microsoft/owa/service.svc", token)
    )

    assert captured == {}
    assert not tokens_file.exists()


def test_token_saved_as_soon_as_captured_and_merged(tokens_file) -> None:
    tokens_file.write_text(json.dumps({"tokens": {"graph.microsoft.com": "old-graph"}}))
    captured: dict[str, str] = {}
    token = _jwt({"aud": "https://outlook.office.com", "scp": "Mail.ReadWrite"})

    auth._make_request_interceptor(captured)(
        _Request("https://outlook.office.com/api/v2.0/me", token)
    )

    saved = json.loads(tokens_file.read_text())["tokens"]
    assert saved == {"graph.microsoft.com": "old-graph", "outlook.office.com": token}
