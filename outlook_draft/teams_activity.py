"""Shared Teams activity fields for CLI and MCP output."""
from __future__ import annotations

from typing import Any

from .formatting import html_to_text


def is_system_event(message: dict[str, Any]) -> bool:
    content = (message.get("body") or {}).get("content") or ""
    return (
        bool(message.get("eventDetail"))
        or message.get("messageType") == "systemEventMessage"
        or "<systemEventMessage" in content
    )


def activity_time(chat: dict[str, Any]) -> str:
    preview = chat.get("lastMessagePreview") or {}
    return preview.get("createdDateTime") or ""


def activity_fields(chat: dict[str, Any]) -> dict[str, str]:
    preview = chat.get("lastMessagePreview") or {}
    sender = preview.get("from") or {}
    label = next(
        (entity.get("displayName") or entity.get("id")
         for kind in ("user", "application", "device")
         if (entity := sender.get(kind))), "",
    )
    body = preview.get("body") or {}
    text = body.get("content") or ""
    if (body.get("contentType") or "").lower() == "html":
        text = html_to_text(text)
    if preview.get("isDeleted") or preview.get("deletedDateTime"):
        text = "[deleted message]"
    elif is_system_event(preview):
        text = "[system event]"
        label = label or "System"
    elif preview and not text.strip():
        has_content = preview.get("attachments") or "<img" in (body.get("content") or "") or "<attachment" in (body.get("content") or "")
        text = "[attachment or rich content]" if has_content else "[empty message]"
    return {
        "lastMessageTime": preview.get("createdDateTime") or "",
        "lastSender": label or "",
        "preview": " ".join(text.split()),
    }
