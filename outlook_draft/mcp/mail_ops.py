"""Complete mail queries, conversations and guarded bulk actions."""
from __future__ import annotations

from ..mail_query import MailQueryMixin
from .clients import outlook_client
from .mail import message_summary
from .shaping import body_text


def mail_query(
    query: str = "", sender: str | None = None, recipient: str | None = None,
    after: str | None = None, before: str | None = None, has_attachments: bool | None = None,
    unread: bool | None = None, folder_id: str | None = None, page: str | None = None,
    all_pages: bool = False, page_size: int = 100,
) -> dict:
    """Search mail with exact address and date filters. Filters apply to fetched pages. Use all_pages for exhaustive structured queries. Keyword searches can have a server-side limit. Keep all filters unchanged for a continuation page. Read-only."""
    filters = {key: value for key, value in {
        "sender": sender, "recipient": recipient, "after": after, "before": before,
        "has_attachments": has_attachments, "unread": unread,
    }.items() if value is not None}
    with outlook_client() as client:
        result = client.query_mail(query, filters=filters, folder=folder_id, page=page,
                                   all_pages=all_pages, page_size=page_size)
    result["messages"] = [message_summary(message) for message in result["messages"]]
    return result


def mail_conversation(message_id: str) -> dict:
    """Read every accessible message in a conversation, oldest first, including bodies and attachment metadata. Read-only. Does not download attachment bytes."""
    with outlook_client() as client:
        result = client.mail_conversation(message_id)
    result["messages"] = [{**message_summary(message), "body": body_text(message.get("Body")),
                           "attachments": message.get("Attachments", [])} for message in result["messages"]]
    return result


def mail_bulk(message_ids: list[str], action: str, folder_id: str | None = None, confirm: bool = False) -> dict:
    """Preview archive, trash, move, mark-read or mark-unread for explicit message IDs. Apply only with confirm=true after user approval. Stops after a failed or uncertain write. Never retry uncertain results automatically. No permanent deletion."""
    preview = MailQueryMixin.bulk_mail(None, message_ids, action, folder=folder_id)
    if not confirm:
        return preview
    with outlook_client() as client:
        return client.bulk_mail(message_ids, action, folder=folder_id, confirm=True)


def mail_trash(message_ids: list[str], confirm: bool = False) -> dict:
    """Preview a move to Deleted Items. Set confirm=true only after user approval. No permanent deletion."""
    return mail_bulk(message_ids, "trash", confirm=confirm)


TOOLS = [mail_query, mail_conversation, mail_bulk, mail_trash]
