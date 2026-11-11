"""Outbound tools. Preview unless the user explicitly approves sending."""

from __future__ import annotations

import base64

from .. import config
from .. import teams_new as teams_new_lib
from ..formatting import compose_email_html, parse_recipients
from ..mail_compose import encode_attachments
from ..signatures import load_signature
from ..teams_media import (
    DEFAULT_UPLOAD_FOLDER,
    media_plan,
    prepare_media,
    resolve_upload_folder,
    send_media,
    validate_content,
)
from .clients import graph_client, outlook_client
from .drafts import _check_importance, draft_summary
from .shaping import JsonDict, addresses, body_text


def draft_send(draft_id: str, confirm: bool = False) -> JsonDict:
    """Preview an existing draft. Set confirm=true only when the user explicitly asks to send it. Sends email immediately; do not retry after an uncertain result."""
    if not draft_id.strip():
        raise ValueError("draft_id must not be empty")
    with outlook_client() as client:
        draft = client.get_draft(draft_id)
        if not draft.get("IsDraft", False):
            raise ValueError("Refusing to send: this message is not a draft.")
        if not any(draft.get(key) for key in ("ToRecipients", "CcRecipients", "BccRecipients")):
            raise ValueError("Draft must have at least one recipient")
        result = {
            **draft_summary(draft), "cc": addresses(draft.get("CcRecipients")),
            "bcc": addresses(draft.get("BccRecipients")), "body": body_text(draft.get("Body")),
            "sent": False,
        }
        if confirm is True:
            client.send_message(draft_id)
            result["sent"] = True
    return result


def mail_send(
    to: list[str], subject: str, body: str,
    cc: list[str] | None = None, bcc: list[str] | None = None,
    html: bool = False, importance: str = "Normal", confirm: bool = False,
    attachments: list[str] | None = None,
) -> JsonDict:
    """Preview a new email with Aptos and the new-message signature. Set confirm=true only when the user explicitly asks to send. Saves to Sent Items. Do not retry after an uncertain result."""
    encoded = encode_attachments(attachments)
    recipients = parse_recipients(to)
    if not recipients:
        raise ValueError("At least one recipient is required in 'to'")
    if not body.strip():
        raise ValueError("body must not be empty")
    importance = _check_importance(importance)
    html_body = compose_email_html(
        body, is_html=html, signature_html=load_signature(config.SIGNATURE_NEW_FILE),
    )
    cc_list, bcc_list = parse_recipients(cc), parse_recipients(bcc)
    result = {
        "to": recipients, "cc": cc_list, "bcc": bcc_list, "subject": subject,
        "bodyHtml": html_body, "importance": importance, "sent": False,
    }
    if attachments:
        result["attachments"] = [{"name": item["Name"], "size": len(base64.b64decode(item["ContentBytes"]))} for item in encoded]
    if confirm is True:
        with outlook_client() as client:
            client.send_mail(
                subject=subject, body=html_body, to=recipients, cc=cc_list or None,
                bcc=bcc_list or None, content_type="HTML", importance=importance,
                save_to_sent_items=True,
                **({"attachments": encoded} if attachments else {}),
            )
        result["sent"] = True
    return result


def teams_send(
    chat_id: str, content: str = "", html: bool = False, confirm: bool = False,
    attachments: list[str] | None = None, images: list[str] | None = None,
    upload_folder: str = DEFAULT_UPLOAD_FOLDER,
) -> JsonDict:
    """Preview a Teams message, local file attachments or inline images. Files upload to an existing private OneDrive folder and receive read access for current internal chat members only. Images use hosted content. Set confirm=true only after the user approves the message, files and access. Never retry uncertain results. Do not use from the Teams gateway."""
    if not chat_id.strip():
        raise ValueError("chat_id must not be empty")
    media = prepare_media(attachments, images)
    validate_content(content, media, is_html=html)
    if media:
        if confirm is True or media.files:
            with graph_client() as client:
                return (send_media(client, chat_id, content, media, is_html=html, upload_folder=upload_folder)
                        if confirm is True else media_plan(client, chat_id, content, media, upload_folder, is_html=html))
        return media_plan(None, chat_id, content, media, upload_folder, is_html=html)
    result = {"chat_id": chat_id, "content": content, "contentType": "html" if html else "text", "sent": False}
    if confirm is True:
        with graph_client() as client:
            message = client.send_teams_message(chat_id, content, content_type=result["contentType"])
        result.update(sent=True, message=message)
    return result


def teams_new(
    to: list[str], content: str = "", topic: str | None = None,
    html: bool = False, confirm: bool = False,
    attachments: list[str] | None = None, images: list[str] | None = None,
    upload_folder: str = DEFAULT_UPLOAD_FOLDER,
) -> JsonDict:
    """Preview a new Teams chat and its first message, with optional local files and inline images.

    `to` accepts emails, UPNs, user IDs or unique names.
    One person reuses an existing 1:1 chat when possible.
    Two or more people create a new group chat each time.
    Confirmation also approves file uploads and read access for the named internal members.
    Existing owner access remains unchanged.
    Set confirm=true only after user approval. Never retry an uncertain result automatically.
    """
    media = prepare_media(attachments, images)
    validate_content(content, media, is_html=html)
    refs = teams_new_lib.split_refs(to)
    with graph_client() as client:
        me, users = teams_new_lib.resolve_recipients(client, refs)
        teams_new_lib.validate_new_chat(users, content, topic, has_media=bool(media))
        result: JsonDict = {
            "chatType": teams_new_lib.chat_type_for(users),
            "members": [
                {"id": user["id"], "displayName": user.get("displayName"),
                 "mail": user.get("mail") or user.get("userPrincipalName")}
                for user in users
            ],
            "topic": topic, "content": content,
            "contentType": "html" if html else "text", "sent": False,
        }
        if media:
            result.update(media.summary())
        if media.files:
            folder = resolve_upload_folder(client, upload_folder)
            result.update(uploadFolder={"id": folder["id"], "path": upload_folder}, accessUserIds=[user["id"] for user in users])
        if confirm is True:
            chat, message = teams_new_lib.start_chat(
                client, me, users, content, content_type=result["contentType"], topic=topic,
                **({"media": media, "upload_folder": upload_folder} if media else {}),
            )
            result.update(sent=True, chat_id=chat.get("id"), webUrl=chat.get("webUrl"), message=message)
    return result


TOOLS = [mail_send, draft_send, teams_send, teams_new]
