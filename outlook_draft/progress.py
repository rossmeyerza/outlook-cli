from __future__ import annotations

import os
from contextlib import contextmanager
from typing import Any, Iterator

from rich.console import Console

_status_console = Console(stderr=True)


def _spinner_disabled(args: Any | None = None) -> bool:
    value = os.environ.get("OUTLOOK_CLI_NO_SPINNER", "")
    if value.casefold() in {"1", "true", "yes", "on"}:
        return True
    return bool(getattr(args, "no_spinner", False))


@contextmanager
def spinner(args: Any | None, message: str) -> Iterator[None]:
    """Show an indeterminate status spinner for interactive human runs."""
    if _spinner_disabled(args) or not _status_console.is_terminal:
        yield
        return

    with _status_console.status(f"[dim]{message}[/]", spinner="dots"):
        yield
