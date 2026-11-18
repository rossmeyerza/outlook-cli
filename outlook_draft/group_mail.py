"""Read Microsoft 365 group inboxes through Microsoft Graph."""
from __future__ import annotations

from collections.abc import Iterator
from datetime import UTC, datetime
from urllib.parse import quote, urlsplit
from uuid import UUID

from .config import GRAPH_BASE_URL
from .errors import OutlookAPIError, OutlookDraftError

THREAD_FIELDS = "id,topic,lastDeliveredDateTime"
POST_FIELDS = "id,receivedDateTime,createdDateTime,from,sender"
OLDEST = datetime.min.replace(tzinfo=UTC)


def parse_timestamp(value: str) -> datetime:
    parsed = datetime.fromisoformat(value)
    return parsed.astimezone(UTC) if parsed.tzinfo else parsed.replace(tzinfo=UTC)


def post_time(post: dict) -> datetime:
    value = post.get("receivedDateTime") or post.get("createdDateTime")
    return parse_timestamp(value) if value else OLDEST


def thread_time(thread: dict) -> datetime:
    value = thread.get("lastDeliveredDateTime")
    return parse_timestamp(value) if value else OLDEST


def group_id(value: str) -> str:
    try:
        return str(UUID(value))
    except ValueError as error:
        raise ValueError("Use the full group ID from 'outlook-cli groups list'.") from error


def sender(post: dict) -> str:
    address = (post.get("from") or post.get("sender") or {}).get("emailAddress") or {}
    return address.get("address") or address.get("name") or ""


def _items(client, path: str, fields: str) -> Iterator[dict]:
    params = {"$select": fields}
    base = urlsplit(GRAPH_BASE_URL)
    visited = set()
    while path:
        if path in visited:
            raise OutlookDraftError("Graph repeated a group inbox page.")
        visited.add(path)
        try:
            page = client._request("GET", path, params=params).json()
        except OutlookAPIError as error:
            if error.status == 403:
                raise OutlookAPIError(
                    403, "Check group access and the Microsoft Graph Group-Conversation.Read.All permission."
                ) from error
            raise
        yield from page["value"]
        path, params = page.get("@odata.nextLink"), None
        if path:
            link = urlsplit(path)
            if (link.scheme, link.netloc) != (base.scheme, base.netloc) or not link.path.startswith(base.path + "/"):
                raise OutlookDraftError("Graph returned an invalid group inbox page link.")


def threads(client, group: str) -> Iterator[dict]:
    return _items(client, f"/groups/{group_id(group)}/threads", THREAD_FIELDS)


def posts(client, group: str, thread: str, *, include_body: bool = True) -> Iterator[dict]:
    if thread in ("", ".", ".."):
        raise ValueError("Use the full thread ID from 'outlook-cli groups threads'.")
    fields = POST_FIELDS + ",body" if include_body else POST_FIELDS
    return _items(client, f"/groups/{group_id(group)}/threads/{quote(thread, safe='')}/posts", fields)


def recent_posts(client, group: str, after: datetime) -> list[dict]:
    pending = {}
    for thread in threads(client, group):
        if thread.get("lastDeliveredDateTime") and thread_time(thread) < after:
            continue
        thread_id = thread["id"]
        try:
            thread_posts = list(posts(client, group, thread_id, include_body=False))
        except OutlookAPIError as error:
            if error.status == 404:
                continue  # The thread can be removed between requests.
            raise
        for post in thread_posts:
            when = post_time(post)
            if when >= after:
                pending[(thread_id, post["id"])] = {
                    "groupId": group,
                    "threadId": thread_id,
                    "postId": post["id"],
                    "receivedDateTime": when.isoformat(),
                    "subject": thread.get("topic") or "",
                    "sender": sender(post),
                }
    return sorted(pending.values(), key=lambda post: post["receivedDateTime"])
