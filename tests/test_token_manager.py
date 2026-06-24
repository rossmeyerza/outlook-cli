from __future__ import annotations

import time

import pytest

from outlook_draft.errors import TokenExpiredError, TokenNotFoundError
from outlook_draft.token_manager import TokenManager


def test_get_token_auto_reauth_when_missing(tmp_path) -> None:
    tm = TokenManager(tokens_file=tmp_path / "missing-tokens.json", token_label="Graph")

    def run_reauth(*, headless: bool = True) -> bool:
        tm._token = "fresh-token"
        tm._expiry = time.time() + 3600
        return True

    tm.run_reauth = run_reauth  # type: ignore[method-assign]

    assert tm.get_token(auto_reauth=True) == "fresh-token"


def test_get_token_auto_reauth_when_expired(tmp_path) -> None:
    tm = TokenManager(tokens_file=tmp_path / "missing-tokens.json", token_label="Graph")
    tm._token = "expired-token"
    tm._expiry = time.time() - 1

    def run_reauth(*, headless: bool = True) -> bool:
        tm._token = "fresh-token"
        tm._expiry = time.time() + 3600
        return True

    tm.run_reauth = run_reauth  # type: ignore[method-assign]

    assert tm.get_token(auto_reauth=True) == "fresh-token"


def test_get_token_raises_clean_error_when_reauth_fails(tmp_path) -> None:
    tm = TokenManager(tokens_file=tmp_path / "missing-tokens.json", token_label="Graph")
    tm._token = "expired-token"
    tm._expiry = time.time() - 1
    tm.run_reauth = lambda *, headless=True: False  # type: ignore[method-assign]

    with pytest.raises(TokenExpiredError):
        tm.get_token(auto_reauth=True)


def test_get_token_raises_missing_when_no_token_and_reauth_fails(tmp_path) -> None:
    tm = TokenManager(tokens_file=tmp_path / "missing-tokens.json", token_label="Graph")
    tm.run_reauth = lambda *, headless=True: False  # type: ignore[method-assign]

    with pytest.raises(TokenNotFoundError):
        tm.get_token(auto_reauth=True)
