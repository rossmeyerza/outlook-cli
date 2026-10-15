"""Small helpers that turn Outlook/Graph payloads into lean tool results."""

from __future__ import annotations

from pathlib import Path
from typing import Any

from ..formatting import html_to_text

JsonDict = dict[str, Any]
PREVIEW_CHARS = 300


def address(obj: JsonDict | None) -> JsonDict:
    """Shape a `{"EmailAddress": {...}}` wrapper as name and address."""
    email = (obj or {}).get("EmailAddress") or {}
    return {"name": email.get("Name", ""), "address": email.get("Address", "")}


def addresses(items: list[JsonDict] | None) -> list[JsonDict]:
    return [address(item) for item in items or []]


def truncate(text: str, limit: int) -> tuple[str, bool]:
    """Cut text to limit characters, reporting whether anything was removed."""
    if limit > 0 and len(text) > limit:
        return text[:limit], True
    return text, False


def body_text(body: JsonDict | None) -> str:
    """Return a message or event body as plain text."""
    body = body or {}
    content = body.get("Content") or body.get("content") or ""
    if (body.get("ContentType") or body.get("contentType") or "").lower() == "html":
        return html_to_text(content)
    return content.strip()


def require_absolute_dir(path: str, name: str = "dest_dir") -> Path:
    """Validate that a destination path is absolute and return it as a Path."""
    resolved = Path(path).expanduser()
    if not resolved.is_absolute():
        raise ValueError(f"{name} must be an absolute path, got {path!r}")
    return resolved
