"""Shared CLI console and authenticated client factories."""
from __future__ import annotations

from .errors import CliError

import json
import sys
from typing import Any

from rich.console import Console

from . import config
from .errors import TokenExpiredError, TokenNotFoundError
from .outlook_client import OutlookClient
from .token_manager import TokenManager

console = Console()

def get_client(
    *,
    base_url: str = config.OUTLOOK_BASE_URL,
    token_domain: str = config.OUTLOOK_TOKEN_DOMAIN,
    token_label: str = "Outlook API",
) -> OutlookClient:
    """Build a client, handling auth errors."""
    tm = TokenManager(token_domain=token_domain, token_label=token_label)
    try:
        _ = tm.get_token(auto_reauth=True)
    except (TokenNotFoundError, TokenExpiredError):
        raise CliError(f"[red]{token_label} authentication failed. Run 'outlook-cli auth' to retry.[/]")
    return OutlookClient(tm, base_url=base_url)

def get_graph_client() -> OutlookClient:
    """Build a Microsoft Graph client."""
    return get_client(
        base_url=config.GRAPH_BASE_URL,
        token_domain=config.GRAPH_TOKEN_DOMAIN,
        token_label="Microsoft Graph",
    )

def print_json(value: Any) -> None:
    print(json.dumps(value, indent=2, ensure_ascii=False))
