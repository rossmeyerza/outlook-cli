"""Read-only Teams tools. Send tools use the separate outbound module."""

from __future__ import annotations

import logging
from typing import Any

from ..commands.files import _write_download
from ..commands.mail import _safe_filename, _unique_path
from ..commands.teams import (
    SELF_CHAT_ID,
    _chat_match_reasons,
    _chat_title,
    _collect_message_links_and_attachments,
    _latest_received_message_time,
    _resolve_chat_attachment_index,
    _search_members,
    _teams_sender,
    find_self_chat,
)
from ..concurrency import ordered_map
from ..errors import OutlookAPIError
from ..formatting import html_to_text
from ..teams_activity import activity_fields, activity_time, is_system_event
from .clients import clamp, graph_client
from .shaping import JsonDict, require_absolute_dir

MEMBER_LIMIT = 10


def _members(client: Any, chat: JsonDict, limit: int) -> list[JsonDict]:
    """Fetch chat members when the chat has no topic. Failures yield no members."""
    if "members" in chat and not chat.get("members@odata.nextLink"):
        return chat["members"]
    if chat.get("topic"):
        return []
    try:
        return client.list_teams_chat_members(chat["id"], top=limit)
    except OutlookAPIError:
        return []


def _chat_summary(chat: JsonDict, members: list[JsonDict]) -> JsonDict:
    return {
        "id": chat.get("id", ""),
        "title": _chat_title(chat, members),
        "chatType": chat.get("chatType", ""),
        "lastUpdated": chat.get("lastUpdatedDateTime", ""),
        **activity_fields(chat),
        "webUrl": chat.get("webUrl", ""),
        "created": chat.get("createdDateTime", ""),
        "lastReceived": chat.get("_lastReceivedMessageDateTime", ""),
        "members": [{"id": member.get("id", ""), "userId": member.get("userId", ""),
                     "displayName": member.get("displayName", ""), "email": member.get("email", ""),
                     "roles": member.get("roles", [])} for member in members],
    }


def teams_list(count: int = 20, sort_received: bool = False) -> list[JsonDict]:
    """List recent Teams chats. Set sort_received=true to sort the fetched chats by latest received message, not all historical chats. Read-only."""
    with graph_client() as client:
        chats = client.list_teams_chats(top=clamp(count))
        if sort_received:
            try:
                user_id = client.get_current_user().get("id", "")
            except OutlookAPIError as exc:
                logging.getLogger(__name__).warning("Cannot identify self for the received sort: %s", exc)
                user_id = ""
            with ordered_map(lambda chat: _latest_received_message_time(client, chat, user_id), chats) as times:
                for chat, timestamp in zip(chats, times, strict=False):
                    chat["_lastReceivedMessageDateTime"] = timestamp
            chats.sort(key=lambda chat: chat.get("_lastReceivedMessageDateTime", ""), reverse=True)
        else:
            chats.sort(key=activity_time, reverse=True)
        with ordered_map(lambda chat: _members(client, chat, MEMBER_LIMIT), chats) as memberships:
            return [_chat_summary(chat, members) for chat, members in zip(chats, memberships, strict=False)]


def teams_search(
    query: str, scan: int = 200, count: int = 10,
    one_on_one: bool = False, include_messages: bool = False, member_count: int = 100,
) -> list[JsonDict]:
    """Search recent chats by topic or participant. Use one_on_one to find a person's chat. Use include_messages to read recent messages. Read-only. Increase scan/count if searchIncomplete is true."""
    if not query.strip():
        raise ValueError("query must not be empty")
    results: list[JsonDict] = []
    member_errors = 0
    with graph_client() as client:
        chats = client.list_teams_chats(top=clamp(scan))
        chats.sort(key=activity_time, reverse=True)
        scan_limited = getattr(chats, "has_more", False)
        if one_on_one:
            chats = [chat for chat in chats if chat.get("chatType") == "oneOnOne"]
        with ordered_map(
            lambda chat: _search_members(client, chat, query, clamp(member_count)), chats,
        ) as memberships:
            for chat, (members, failed) in zip(chats, memberships, strict=False):
                member_errors += int(failed)
                reasons = _chat_match_reasons(chat, members, query)
                if reasons:
                    results.append({**_chat_summary(chat, members), "match": "; ".join(reasons)})
                    if len(results) > clamp(count):
                        break
    incomplete = scan_limited or len(results) > clamp(count) or member_errors > 0
    if member_errors:
        logging.getLogger(__name__).warning("Teams search could not read members for %d chats.", member_errors)
    if scan_limited:
        logging.getLogger(__name__).warning("Teams search may be truncated: Graph has more chats beyond scan=%d.", clamp(scan))
    if len(results) > clamp(count):
        logging.getLogger(__name__).warning("Teams matching results truncated at count=%d.", clamp(count))
    results = results[:clamp(count)]
    for result in results:
        result["searchIncomplete"] = incomplete
        result["memberLookupErrors"] = member_errors
        if include_messages:
            result["messages"] = teams_messages(result["id"])["messages"]
    return results


def _message_text(message: JsonDict) -> str:
    if message.get("deletedDateTime"):
        return "[deleted message]"
    if is_system_event(message):
        return "[system event]"
    body = message.get("body") or {}
    content = body.get("content") or ""
    if (body.get("contentType") or "").lower() == "html":
        content = html_to_text(content)
    content = content.strip()
    if content:
        return content
    if message.get("attachments") or message.get("hostedContents") or "<img" in (body.get("content") or "") or "<attachment" in (body.get("content") or ""):
        return "[attachment or rich content]"
    return "[empty message]"


def teams_messages(chat_id: str, count: int = 20) -> JsonDict:
    """Read recent messages in a Teams chat by full chat ID, newest first. Read-only."""
    with graph_client() as client:
        if chat_id == SELF_CHAT_ID:
            chat: JsonDict = {"id": SELF_CHAT_ID, "topic": "Self chat", "chatType": "oneOnOne"}
            members: list[JsonDict] = []
        else:
            chat = client.get_teams_chat(chat_id)
            members = client.list_teams_chat_members(chat_id, top=50)
        messages = client.list_teams_messages(chat_id, top=clamp(count))
    messages.sort(key=lambda m: m.get("createdDateTime", ""), reverse=True)
    return {
        "chat": _chat_summary(chat, members),
        "messages": [
            {
                "id": m.get("id", ""),
                "created": m.get("createdDateTime", ""),
                "sender": _teams_sender(m),
                "text": _message_text(m),
                "hasAttachments": bool(m.get("attachments")),
            }
            for m in messages
        ],
    }


def teams_attachments(chat_id: str, count: int = 50) -> list[JsonDict]:
    """List file attachments and SharePoint/OneDrive links in the latest messages of a chat. Count controls the message scan. Read-only."""
    with graph_client() as client:
        messages = client.list_teams_messages(chat_id, top=clamp(count))
    return _collect_message_links_and_attachments(messages)


def teams_download_attachments(
    chat_id: str,
    dest_dir: str,
    attachment: str | None = None,
    count: int = 50,
    overwrite: bool = False,
) -> list[str]:
    """Download Teams attachments or share links from a chat to an absolute directory. `attachment` is a 1-based index from teams_attachments, a URL, or a name fragment; omit for all. Writes files."""
    out_dir = require_absolute_dir(dest_dir)
    written: list[str] = []
    with graph_client() as client:
        messages = client.list_teams_messages(chat_id, top=clamp(count))
        selected = _resolve_chat_attachment_index(
            _collect_message_links_and_attachments(messages), attachment
        )
        if not selected:
            return written
        out_dir.mkdir(parents=True, exist_ok=True)
        for item in selected:
            metadata = client.get_share_drive_item(item["url"])
            content = client.download_share_url(item["url"])
            name = _safe_filename(metadata.get("name") or item["name"] or "download")
            path = _unique_path(out_dir, name, overwrite=overwrite)
            _write_download(path, content, overwrite=overwrite)
            written.append(str(path))
    return written


def teams_get(chat_id: str) -> JsonDict:
    """Read chat details, created time and members without a message fetch. Use a full chat ID or 48:notes for self-chat discovery. Read-only."""
    with graph_client() as client:
        if chat_id == SELF_CHAT_ID:
            found = find_self_chat(client)
            if not found:
                raise ValueError("Teams self-chat was not found.")
            chat, members = found
        else:
            chat = client.get_teams_chat(chat_id)
            members = client.list_teams_chat_members(chat_id, top=50)
    return _chat_summary(chat, members)


def teams_self() -> JsonDict:
    """Discover and check access to your Teams self-chat. Read-only."""
    with graph_client() as client:
        found = find_self_chat(client)
    if not found:
        return {"found": False}
    chat, members = found
    return {"found": True, **_chat_summary(chat, members)}


TOOLS = [teams_list, teams_search, teams_get, teams_self, teams_messages, teams_attachments, teams_download_attachments]
