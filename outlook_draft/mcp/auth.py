"""Auth status tool. Read-only: never triggers re-authentication."""

from __future__ import annotations

import logging

import json
import time

from .. import config
from ..token_manager import _decode_claims
from .shaping import JsonDict

TOKENS = (
    ("Outlook API", config.OUTLOOK_TOKEN_DOMAIN),
    ("Microsoft Graph", config.GRAPH_TOKEN_DOMAIN),
)


log = logging.getLogger(__name__)


def _token_state(tokens: JsonDict, label: str, domain: str) -> JsonDict:
    state: JsonDict = {"label": label, "domain": domain, "present": False, "valid": False}
    token = tokens.get(domain)
    if not token:
        return state
    try:
        claims = _decode_claims(token)
    except Exception:
        log.debug("Best-effort operation failed", exc_info=True)
        state["present"] = True
        state["error"] = "token could not be decoded"
        return state
    seconds_left = float(claims.get("exp", 0)) - time.time()
    state.update(
        present=True,
        valid=seconds_left > 0,
        expired=seconds_left <= 0,
        minutesLeft=round(seconds_left / 60, 1),
        user=claims.get("upn") or claims.get("unique_name") or "",
    )
    return state


def auth_status() -> JsonDict:
    """Report whether the Outlook and Graph tokens are present and valid, and minutes left. Read-only. If expired, run 'outlook-cli auth' in a terminal."""
    tokens: JsonDict = {}
    if config.TOKENS_FILE.exists():
        try:
            tokens = json.loads(config.TOKENS_FILE.read_text()).get("tokens", {})
        except (OSError, json.JSONDecodeError):
            tokens = {}
    return {
        "account": config.ACTIVE_ACCOUNT or "default",
        "tokens": [_token_state(tokens, label, domain) for label, domain in TOKENS],
    }


TOOLS = [auth_status]
