"""Client factories for MCP tools.

Tools build a fresh client per call so tokens refreshed by `outlook-cli auth`
take effect without restarting the server. Re-authentication is never run from
here: headless reauth prints an MFA number that a stdio server cannot show.
"""

from __future__ import annotations

from collections.abc import Iterator
from contextlib import contextmanager

from .. import config
from ..copilot import CopilotClient
from ..errors import OutlookDraftError, TokenExpiredError, TokenNotFoundError
from ..outlook_client import OutlookClient
from ..token_manager import CopilotTokenManager, TokenManager


class McpAuthError(OutlookDraftError):
    """Token is missing or expired and the user must re-authenticate."""


class NoReauthTokenManager(TokenManager):
    """Token manager that never launches browser re-authentication."""

    def get_token(self, *, auto_reauth: bool = False, headless: bool = True) -> str:
        return super().get_token(auto_reauth=False)

    def run_reauth(self, *, headless: bool = True) -> bool:
        return False


def _token_manager(domain: str, label: str) -> NoReauthTokenManager:
    manager = NoReauthTokenManager(token_domain=domain, token_label=label)
    try:
        manager.get_token()
    except (TokenNotFoundError, TokenExpiredError) as exc:
        raise McpAuthError(
            f"{label} token expired or missing. Run 'outlook-cli auth' in a terminal, then retry."
        ) from exc
    return manager


@contextmanager
def outlook_client() -> Iterator[OutlookClient]:
    """Outlook REST client (mail, drafts, calendar, tasks, contacts)."""
    client = OutlookClient(_token_manager(config.OUTLOOK_TOKEN_DOMAIN, "Outlook API"))
    try:
        yield client
    finally:
        client.close()


@contextmanager
def graph_client() -> Iterator[OutlookClient]:
    """Microsoft Graph client (Teams, share links)."""
    manager = _token_manager(config.GRAPH_TOKEN_DOMAIN, "Microsoft Graph")
    client = OutlookClient(manager, base_url=config.GRAPH_BASE_URL)
    try:
        yield client
    finally:
        client.close()


@contextmanager
def files_client() -> Iterator[OutlookClient]:
    """Graph client for OneDrive and SharePoint files."""
    client = OutlookClient(
        _token_manager(config.GRAPH_TOKEN_DOMAIN, "Microsoft Graph"),
        base_url=config.GRAPH_BASE_URL,
        timeout=60.0,
        json_content_type=False,
        log_retries=False,
    )
    try:
        yield client
    finally:
        client.close()


@contextmanager
def copilot_client() -> Iterator[CopilotClient]:
    """Microsoft 365 Copilot chat client.

    The Copilot token lasts about an hour. It is the one token that this server
    refreshes: the refresh uses the saved browser session and has no password
    or MFA step. If the saved session has expired, the user must run auth.
    """
    manager = CopilotTokenManager(full_login_fallback=False)
    try:
        manager.get_token(auto_reauth=True)
    except (TokenNotFoundError, TokenExpiredError) as exc:
        raise McpAuthError(
            "Microsoft 365 Copilot token expired or missing. Run 'outlook-cli auth' in a terminal, then retry."
        ) from exc
    yield CopilotClient(manager)


def clamp(value: int, *, low: int = 1, high: int = 100) -> int:
    """Cap a caller-supplied count to a sane range."""
    return max(low, min(int(value), high))
