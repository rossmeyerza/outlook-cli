from __future__ import annotations

import base64
import json
import logging
import time
from pathlib import Path

import jwt

from . import config
from .errors import TokenExpiredError, TokenNotFoundError

log = logging.getLogger(__name__)


def _decode_claims(token: str) -> dict:
    """Decode JWT claims without signature verification."""
    try:
        return jwt.decode(token, options={"verify_signature": False}, algorithms=["RS256"])
    except Exception:
        payload = token.split(".")[1]
        payload += "=" * (4 - len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload))


class TokenManager:
    """Manages API tokens from outlook-draft-cli's local tokens.json."""

    def __init__(
        self,
        tokens_file: Path = config.TOKENS_FILE,
        *,
        token_domain: str = config.OUTLOOK_TOKEN_DOMAIN,
        token_label: str = "Outlook API",
    ):
        self._tokens_file = tokens_file
        self._token_domain = token_domain
        self._token_label = token_label
        self._token: str | None = None
        self._expiry: float = 0.0
        self._file_mtime: float = 0.0

    @property
    def token(self) -> str:
        """Get the current valid API token, reloading if file changed."""
        self._maybe_reload()
        if self._token is None:
            raise TokenNotFoundError(f"No {self._token_label} token available")
        if time.time() >= self._expiry:
            raise TokenExpiredError(f"{self._token_label} token has expired")
        return self._token

    @property
    def is_expired(self) -> bool:
        self._maybe_reload()
        if self._token is None:
            return True
        return time.time() >= self._expiry

    @property
    def expires_in(self) -> float:
        """Seconds until token expiry. Negative if expired."""
        if self._token is None:
            return -1
        return self._expiry - time.time()

    def _maybe_reload(self) -> None:
        if not self._tokens_file.exists():
            return
        mtime = self._tokens_file.stat().st_mtime
        if mtime != self._file_mtime:
            self._load_from_file()

    def _load_from_file(self) -> None:
        try:
            data = json.loads(self._tokens_file.read_text())
            tokens = data.get("tokens", {})
            token = tokens.get(self._token_domain)

            if not token:
                log.warning("No %s token in %s", self._token_domain, self._tokens_file)
                return

            claims = _decode_claims(token)
            exp = claims.get("exp", 0)

            self._token = token
            self._expiry = float(exp)
            self._file_mtime = self._tokens_file.stat().st_mtime

            remaining = self._expiry - time.time()
            log.info(
                "Loaded %s token: [REDACTED] in %.0f minutes (user: %s)",
                self._token_domain,
                remaining / 60,
                claims.get("upn", "unknown"),
            )

        except (json.JSONDecodeError, KeyError, IndexError) as e:
            log.error("Failed to parse tokens file: %s", e)

    def run_reauth(self, *, headless: bool = True) -> bool:
        """Run the local auth flow to capture fresh tokens."""
        from .auth import capture_tokens_via_browser

        log.info("Running local re-authentication")
        try:
            captured = capture_tokens_via_browser(headless=headless)
        except Exception:
            log.exception("Local re-authentication failed")
            return False
        if not captured:
            return False
        self.force_reload()
        return self._token is not None and not self.is_expired

    def force_reload(self) -> None:
        self._file_mtime = 0.0
        self._maybe_reload()
