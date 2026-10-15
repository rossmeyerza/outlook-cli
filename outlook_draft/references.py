"""Resolve CLI references using the active account caches."""
from __future__ import annotations

from .errors import CliError

import sys

from . import cache
from .outlook_client import OutlookClient
from .runtime import console

def _resolve_message_id(client: OutlookClient, ref: str) -> str:
    """Resolve a message reference to a full ID.

    Accepts a numeric index, cached ID suffix, or full ID.
    """
    cached = cache.load_cache(cache.MAIL_CACHE)

    if ref.isdigit():
        idx = int(ref)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print("[red]No search results cached. Run 'outlook-cli mail unread' or 'outlook-cli mail search <query>' first.[/]")
        else:
            console.print(f"[red]Index {idx} out of range. Only {len(cached)} results cached.[/]")
        raise CliError("")

    if len(ref) < 40 and cached:
        matches = [item for item in cached if item["Id"].endswith(ref)]
        if len(matches) == 1:
            return matches[0]["Id"]
        if len(matches) > 1:
            raise CliError(f"[red]Ambiguous message ID suffix '{ref}'. Use a longer ID.[/]")

    return ref



def _resolve_contact_id(ref: str) -> str:
    if ref.isdigit():
        idx = int(ref)
        cached = cache.load_cache(cache.CONTACT_CACHE)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print("[red]No contacts cached. Run 'outlook-cli contact search <query>' first.[/]")
        else:
            console.print(f"[red]Index {idx} out of range.[/]")
        raise CliError("")
    return ref

def _resolve_task_id(client: OutlookClient, ref: str) -> str:
    if ref.isdigit():
        idx = int(ref)
        cached = cache.load_cache(cache.TASK_CACHE)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print("[red]No tasks cached. Run 'outlook-cli task list' first.[/]")
        else:
            console.print(f"[red]Index {idx} out of range.[/]")
        raise CliError("")
    return ref

def _resolve_cal_id(client: OutlookClient, ref: str) -> str:
    if ref.isdigit():
        idx = int(ref)
        cached = cache.load_cache(cache.CAL_CACHE)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print("[red]No agenda cached. Run 'outlook-cli cal agenda' first.[/]")
        else:
            console.print(f"[red]Index {idx} out of range.[/]")
        raise CliError("")
    return ref

