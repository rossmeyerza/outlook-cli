"""Outbound tools. Preview unless the user explicitly approves sending."""

from __future__ import annotations

from .. import config
from ..formatting import compose_email_html, parse_recipients
from ..signatures import load_signature
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
) -> JsonDict:
    """Preview a new email with Aptos and the new-message signature. Set confirm=true only when the user explicitly asks to send. Saves to Sent Items. Do not retry after an uncertain result."""
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
    if confirm is True:
        with outlook_client() as client:
            client.send_mail(
                subject=subject, body=html_body, to=recipients, cc=cc_list or None,
                bcc=bcc_list or None, content_type="HTML", importance=importance,
                save_to_sent_items=True,
            )
        result["sent"] = True
    return result


def teams_send(chat_id: str, content: str, html: bool = False, confirm: bool = False) -> JsonDict:
    """Preview a Teams message to a full chat ID. Set confirm=true only when the user explicitly asks to send. Do not use from the Teams gateway or retry after an uncertain result."""
    if not chat_id.strip():
        raise ValueError("chat_id must not be empty")
    if not content.strip():
        raise ValueError("content must not be empty")
    result = {"chat_id": chat_id, "content": content, "contentType": "html" if html else "text", "sent": False}
    if confirm is True:
        with graph_client() as client:
            message = client.send_teams_message(chat_id, content, content_type=result["contentType"])
        result.update(sent=True, message=message)
    return result


TOOLS = [mail_send, draft_send, teams_send]
