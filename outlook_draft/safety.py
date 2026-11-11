"""Request policies and command prefix controls."""
from __future__ import annotations

import os
import re
from contextlib import contextmanager
from contextvars import ContextVar
from dataclasses import dataclass
from urllib.parse import unquote, urlsplit

from .errors import OutlookDraftError


class SafetyError(OutlookDraftError):
    """An operation conflicts with the active safety policy."""


@dataclass(frozen=True)
class SafetyPolicy:
    readonly: bool = False
    no_send: bool = False


_DEFAULT_POLICY = SafetyPolicy()
_policy: ContextVar[SafetyPolicy] = ContextVar("outlook_cli_safety", default=_DEFAULT_POLICY)


def _env_flag(name: str) -> bool:
    return os.environ.get(name, "").strip().lower() not in {"", "0", "false", "no", "off"}


def current_policy() -> SafetyPolicy:
    """Return the active restrictions, including process environment controls."""
    active = _policy.get()
    return SafetyPolicy(active.readonly or _env_flag("OUTLOOK_CLI_READONLY"),
                        active.no_send or _env_flag("OUTLOOK_CLI_NO_SEND"))


@contextmanager
def policy(*, readonly: bool = False, no_send: bool = False):
    """Apply restrictions within this context without weakening an outer policy."""
    previous = _policy.get()
    token = _policy.set(SafetyPolicy(previous.readonly or readonly, previous.no_send or no_send))
    try:
        yield
    finally:
        _policy.reset(token)


def check_command(name: str) -> None:
    """Block operations that can bypass the request policy in another process."""
    active = _policy.get()
    restricted = (active.readonly or active.no_send or _env_flag("OUTLOOK_CLI_READONLY")
                  or _env_flag("OUTLOOK_CLI_NO_SEND"))
    if restricted and name in {"copilot.ask", "copilot.app.ask"}:
        raise SafetyError("Safety policy blocks Copilot prompts because Copilot can change remote data.")
    if restricted and name == "gateway.start":
        raise SafetyError("Safety policy blocks gateway start because child processes can change remote data.")


def _can_send(method: str, path: str) -> bool:
    if method in {"GET", "HEAD", "OPTIONS"}:
        return False
    segments = path.split("/")
    actions = {"send", "sendmail", "reply", "replyall", "forward", "accept", "decline", "tentativelyaccept", "cancel"}
    if segments[-1] in actions:
        return True
    if "chats" in segments or "channels" in segments:
        return method == "POST" and segments[-1] in {"messages", "replies"}
    # Event changes can send invitations or updates without a separate send request.
    return "events" in segments and method in {"POST", "PATCH", "PUT", "DELETE"}


def check_request(method: str, url: str) -> None:
    """Reject forbidden requests before authentication or network access."""
    active = _policy.get()
    method = method.upper()
    path = unquote(urlsplit(url).path).rstrip("/").lower()
    readonly = active.readonly or _env_flag("OUTLOOK_CLI_READONLY")
    read_post = method == "POST" and path.endswith(("/getschedule", "/findmeetingtimes"))
    if readonly and method not in {"GET", "HEAD", "OPTIONS"} and not read_post:
        raise SafetyError(f"Read-only policy blocks {method} requests.")
    if (active.no_send or _env_flag("OUTLOOK_CLI_NO_SEND")) and _can_send(method, path):
        raise SafetyError(f"No-send policy blocks {method} requests that can send messages.")


def _prefixes(values: str | list[str] | tuple[str, ...] | None) -> list[str]:
    if values is None:
        return []
    values = [values] if isinstance(values, str) else values
    prefixes = []
    for value in values:
        for prefix in value.split(","):
            parts = re.split(r"[.\s]+", prefix.strip())
            if not prefix.strip() or not all(re.fullmatch(r"[\w-]+", part) for part in parts):
                raise SafetyError("Command prefixes must contain names separated by dots or spaces.")
            prefixes.append(".".join(parts))
    return prefixes


def command_allowed(name: str, enabled=None, disabled=None) -> bool:
    """Match whole command segments; disabled prefixes take precedence."""
    command = ".".join(name.split())
    allow = _prefixes(enabled)
    deny = _prefixes(disabled)

    def matches(prefix: str) -> bool:
        return command == prefix or command.startswith(prefix + ".")

    return (not allow or any(map(matches, allow))) and not any(map(matches, deny))
