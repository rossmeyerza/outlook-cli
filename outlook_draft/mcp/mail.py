"""Mail tools."""

from __future__ import annotations

import base64
from typing import Any

from ..commands.mail import _collect_message_links, _safe_filename, _unique_path
from .clients import clamp, outlook_client
from .shaping import (
    PREVIEW_CHARS,
    JsonDict,
    address,
    addresses,
    body_text,
    require_absolute_dir,
    truncate,
)

WELL_KNOWN_FOLDERS = {"inbox", "archive", "drafts", "sentitems", "deleteditems", "junkemail"}


def message_summary(msg: JsonDict) -> JsonDict:
    summary = {
        "id": msg.get("Id", ""),
        "subject": msg.get("Subject", ""),
        "from": address(msg.get("From")),
        "received": msg.get("ReceivedDateTime", ""),
        "isRead": msg.get("IsRead", True),
        "hasAttachments": msg.get("HasAttachments", False),
        "preview": (msg.get("BodyPreview") or "")[:PREVIEW_CHARS],
    }
    if msg.get("WebLink"):
        summary["webLink"] = msg["WebLink"]
    return summary


def mail_search(query: str, count: int = 20) -> list[JsonDict]:
    """Search mail by keyword. Read-only. Returns message summaries with full IDs."""
    with outlook_client() as client:
        messages = client.search_messages(query, top=clamp(count))
    return [message_summary(m) for m in messages]


def mail_unread(count: int = 20) -> list[JsonDict]:
    """List unread Inbox messages, newest first. Read-only. Returns message summaries."""
    with outlook_client() as client:
        messages = client.list_unread(top=clamp(count))
    return [message_summary(m) for m in messages]


def mail_read(message_id: str, include_html: bool = False, max_chars: int = 20000) -> JsonDict:
    """Read one message by full ID: headers, recipients, body text, attachments and links. Read-only."""
    with outlook_client() as client:
        msg = client.get_message(message_id)
        attachments = client.list_message_attachments(message_id) if msg.get("HasAttachments") else []
    text, truncated = truncate(body_text(msg.get("Body")), max_chars)
    result: JsonDict = {
        **message_summary(msg),
        "to": addresses(msg.get("ToRecipients")),
        "cc": addresses(msg.get("CcRecipients")),
        "importance": msg.get("Importance", ""),
        "body": text,
        "bodyTruncated": truncated,
        "attachments": [_attachment(a) for a in attachments],
        "links": _collect_message_links(msg, share_only=False)[:50],
    }
    if include_html:
        result["bodyHtml"] = (msg.get("Body") or {}).get("Content", "")
    return result


def _attachment(item: JsonDict) -> JsonDict:
    return {
        "id": item.get("Id", ""),
        "name": item.get("Name", ""),
        "contentType": item.get("ContentType", ""),
        "size": item.get("Size", 0),
        "isInline": item.get("IsInline", False),
    }


def mail_attachments(message_id: str) -> list[JsonDict]:
    """List attachments on a message (id, name, contentType, size, isInline). Read-only."""
    with outlook_client() as client:
        items = client.list_message_attachments(message_id)
    return [_attachment(a) for a in items]


def mail_download_attachments(
    message_id: str,
    dest_dir: str,
    attachment_ids: list[str] | None = None,
    overwrite: bool = False,
) -> list[str]:
    """Download file attachments to an absolute directory on this machine. Skips inline images unless named in attachment_ids. Writes files."""
    out_dir = require_absolute_dir(dest_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    written: list[str] = []
    with outlook_client() as client:
        for item in client.list_message_attachments(message_id):
            selected = item.get("Id", "") in attachment_ids if attachment_ids else not item.get("IsInline")
            if not selected:
                continue
            full = client.get_message_attachment(message_id, item.get("Id", ""))
            content = full.get("ContentBytes")
            if not content:
                continue
            path = _unique_path(out_dir, _safe_filename(full.get("Name") or "attachment"), overwrite=overwrite)
            path.write_bytes(base64.b64decode(content))
            written.append(str(path))
    return written


def mail_links(message_id: str, share_only: bool = True) -> list[JsonDict]:
    """List links in a message body. By default only SharePoint/OneDrive share links. Read-only."""
    with outlook_client() as client:
        msg = client.get_message(message_id)
    return _collect_message_links(msg, share_only=share_only)


def mail_folders() -> list[JsonDict]:
    """List top-level mail folders with unread and total counts. Read-only."""
    with outlook_client() as client:
        folders = client.list_mail_folders(top=250)
    return [
        {
            "id": f.get("Id", ""),
            "displayName": f.get("DisplayName", ""),
            "unreadItemCount": f.get("UnreadItemCount", 0),
            "totalItemCount": f.get("TotalItemCount", 0),
        }
        for f in folders
    ]


def mail_mark_read(message_id: str, is_read: bool = True) -> JsonDict:
    """Mark a message read or unread. Changes mailbox state."""
    with outlook_client() as client:
        client.update_message_read_state(message_id, is_read=is_read)
    return {"id": message_id, "isRead": is_read}


def _resolve_folder(client: Any, folder: str) -> str:
    """Resolve a well-known name, display name or folder ID to a folder ID."""
    if folder.lower() in WELL_KNOWN_FOLDERS:
        return folder.lower()
    folders = client.list_mail_folders(top=250)
    by_id = [f for f in folders if f.get("Id") == folder]
    if by_id:
        return folder
    by_name = [f for f in folders if f.get("DisplayName", "").lower() == folder.lower()]
    if len(by_name) == 1:
        return by_name[0]["Id"]
    if by_name:
        raise ValueError(f"Ambiguous folder name {folder!r}. Use the folder ID from mail_folders.")
    raise ValueError(f"No folder matching {folder!r}. Use mail_folders to list folders.")


def mail_move(message_id: str, folder: str) -> JsonDict:
    """Move a message to a folder given by name (e.g. 'inbox', 'archive', a display name) or ID. Changes mailbox state."""
    with outlook_client() as client:
        moved = client.move_message(message_id, _resolve_folder(client, folder))
    return {"id": moved.get("Id", message_id), "subject": moved.get("Subject", "")}


def mail_archive(message_id: str) -> JsonDict:
    """Move a message to the Archive folder. Changes mailbox state."""
    with outlook_client() as client:
        moved = client.archive_message(message_id)
    return {"id": moved.get("Id", message_id), "subject": moved.get("Subject", "")}


TOOLS = [
    mail_search,
    mail_unread,
    mail_read,
    mail_attachments,
    mail_download_attachments,
    mail_links,
    mail_folders,
    mail_mark_read,
    mail_move,
    mail_archive,
]
