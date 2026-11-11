"""Mail queries, complete conversations and explicit bulk operations."""
from __future__ import annotations

from datetime import UTC, datetime
from typing import Any
from urllib.parse import quote

from .errors import OutlookAPIError
from .mail_urls import validate_mail_url

MAIL_FIELDS = "Id,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,IsRead,HasAttachments,BodyPreview,ConversationId"
FILTER_FIELDS = {"sender", "recipient", "after", "before", "has_attachments", "unread"}
BULK_ACTIONS = {"archive", "trash", "move", "mark-read", "mark-unread"}


def _date(value: str) -> datetime:
    parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
    return parsed.replace(tzinfo=UTC) if parsed.tzinfo is None else parsed


def validate_filters(filters: dict | None) -> dict:
    result = dict(filters or {})
    if result.keys() - FILTER_FIELDS:
        raise ValueError("Unknown mail filters: " + ", ".join(sorted(result.keys() - FILTER_FIELDS)))
    for key in ("sender", "recipient"):
        if key in result and (not isinstance(result[key], str) or not result[key].strip()):
            raise ValueError(f"{key} must be a non-empty email address")
    for key in ("has_attachments", "unread"):
        if key in result and not isinstance(result[key], bool):
            raise ValueError(f"{key} must be a boolean")
    for key in ("after", "before"):
        if key in result:
            if not isinstance(result[key], str):
                raise ValueError(f"{key} must be an ISO date or timestamp")
            _date(result[key])
    if "after" in result and "before" in result and _date(result["after"]) >= _date(result["before"]):
        raise ValueError("after must be earlier than before")
    return result


def _matches(message: dict, filters: dict) -> bool:
    sender = (message.get("From") or {}).get("EmailAddress", {}).get("Address", "")
    if "sender" in filters and sender.casefold() != filters["sender"].strip().casefold():
        return False
    recipients = message.get("ToRecipients", []) + message.get("CcRecipients", [])
    if "recipient" in filters and not any(
        r.get("EmailAddress", {}).get("Address", "").casefold() == filters["recipient"].strip().casefold()
        for r in recipients
    ):
        return False
    if "has_attachments" in filters and message.get("HasAttachments", False) != filters["has_attachments"]:
        return False
    if "unread" in filters and (not message.get("IsRead", True)) != filters["unread"]:
        return False
    received = message.get("ReceivedDateTime")
    if "after" in filters and (not received or _date(received) < _date(filters["after"])):
        return False
    return not ("before" in filters and (not received or _date(received) >= _date(filters["before"])))


class MailQueryMixin:
    def _mail_page(self, path: str, params: dict | None = None) -> dict:
        validate_mail_url(path, self._base_url)
        return self._request("GET", path, params=params).json()

    def query_mail(
        self, query: str = "", *, filters: dict | None = None, folder: str | None = None,
        page: str | None = None, all_pages: bool = False, page_size: int = 100,
    ) -> dict[str, Any]:
        filters = validate_filters(filters)
        if folder is not None and not folder.strip():
            raise ValueError("folder must be a non-empty folder ID")
        if not 1 <= page_size <= 1000:
            raise ValueError("page_size must be between 1 and 1000")
        path = f"/me/mailfolders/{quote(folder, safe='')}/messages" if folder else "/me/messages"
        params = {"$top": str(page_size), "$select": MAIL_FIELDS}
        if query:
            params["$search"] = '"' + query.replace('\\', '\\\\').replace('"', '\\"') + '"'
        else:
            params["$orderby"] = "ReceivedDateTime desc"
        items, seen, scanned = [], set(), 0
        url, next_url = page or path, None
        while url:
            if url in seen:
                raise ValueError("The mail API repeated a continuation URL")
            seen.add(url)
            data = self._mail_page(url, params if not page and url == path else None)
            messages = data.get("value", [])
            scanned += len(messages)
            items.extend(m for m in messages if _matches(m, filters))
            next_url = data.get("@odata.nextLink") or data.get("odata.nextLink")
            if not all_pages:
                break
            url = next_url
        complete = not next_url and not query and not page
        return {
            "messages": items, "nextPage": next_url, "matchedCount": len(items),
            "scannedCount": scanned, "complete": complete, "countExact": complete,
            "warning": "Keyword search can have a server-side result limit; the count covers returned results only."
            if query else ("Filters apply to fetched pages only." if not complete else None),
        }

    def mail_conversation(self, message_id: str) -> dict[str, Any]:
        message = self._mail_page(f"/me/messages/{quote(message_id, safe='')}", {"$select": "Id,ConversationId"})
        conversation = message.get("ConversationId")
        if not conversation:
            raise ValueError("The message has no conversation ID")
        params = {"$filter": "ConversationId eq '" + conversation.replace("'", "''") + "'", "$top": "100"}
        url, items, seen = "/me/messages", [], set()
        while url:
            if url in seen:
                raise ValueError("The mail API repeated a conversation page")
            seen.add(url)
            data = self._mail_page(url, params if url == "/me/messages" else None)
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink") or data.get("odata.nextLink")
        items.sort(key=lambda m: _date(m.get("ReceivedDateTime") or m.get("CreatedDateTime") or "0001-01-01"))
        for item in items:
            item["Attachments"] = self._conversation_attachments(item["Id"])
        return {"conversationId": conversation, "messages": items, "complete": True, "count": len(items)}

    def _conversation_attachments(self, message_id: str) -> list[dict]:
        url = f"/me/messages/{quote(message_id, safe='')}/attachments"
        items, seen = [], set()
        params = {"$select": "Id,Name,ContentType,Size,IsInline"}
        while url:
            if url in seen:
                raise ValueError("The mail API repeated an attachment page")
            seen.add(url)
            data = self._mail_page(url, params)
            params = None
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink") or data.get("odata.nextLink")
        return items

    def bulk_mail(self, message_ids: list[str], action: str, *, folder: str | None = None, confirm: bool = False) -> dict:
        if type(confirm) is not bool:
            raise ValueError("confirm must be a boolean")
        if action not in BULK_ACTIONS:
            raise ValueError("Unsupported bulk mail action")
        if not message_ids or any(not isinstance(mid, str) or not mid.strip() for mid in message_ids):
            raise ValueError("Provide at least one non-empty message ID")
        ids = list(dict.fromkeys(message_ids))
        if action == "move" and not folder:
            raise ValueError("move requires a destination folder ID")
        result = {"action": action, "messageIds": ids, "folder": folder, "confirmed": confirm, "results": []}
        if not confirm:
            return result
        for index, mid in enumerate(ids):
            try:
                target = {"archive": "archive", "trash": "deleteditems"}.get(action, folder)
                if action in {"mark-read", "mark-unread"}:
                    value = self._request("PATCH", f"/me/messages/{quote(mid, safe='')}",
                                          json_body={"IsRead": action == "mark-read"}, max_retries=0)
                    output = {"id": mid, "status": "applied"}
                else:
                    value = self._request("POST", f"/me/messages/{quote(mid, safe='')}/move",
                                          json_body={"DestinationId": target}, max_retries=0)
                    output = {"id": mid, "newId": value.json().get("Id", ""), "status": "applied"}
                result["results"].append(output)
            except OutlookAPIError as exc:
                result["results"].append({"id": mid, "status": "uncertain" if exc.status == 0 or exc.status >= 500 else "failed", "error": str(exc)})
                result["results"].extend({"id": remaining, "status": "not_attempted"} for remaining in ids[index + 1:])
                break
        return result
