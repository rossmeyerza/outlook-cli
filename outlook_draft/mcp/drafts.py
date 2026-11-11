"""Draft tools. Drafts are saved for review in Outlook and are never sent."""

from __future__ import annotations

from .. import config
from ..email_markdown import EmailBodyFormat, is_html_format
from ..formatting import compose_email_html, parse_recipients
from ..mail_compose import attach_to_draft, draft_update_fields, encode_attachments, finish_quoted_draft
from ..signatures import load_signature
from .clients import clamp, outlook_client
from .shaping import PREVIEW_CHARS, JsonDict, addresses, body_text

IMPORTANCE_VALUES = {"Low", "Normal", "High"}


def draft_summary(draft: JsonDict) -> JsonDict:
    summary = {
        "id": draft.get("Id", ""),
        "subject": draft.get("Subject", ""),
        "to": addresses(draft.get("ToRecipients")),
        "importance": draft.get("Importance", ""),
        "created": draft.get("CreatedDateTime", ""),
        "preview": (draft.get("BodyPreview") or "")[:PREVIEW_CHARS],
    }
    if draft.get("WebLink"):
        summary["webLink"] = draft["WebLink"]
    return summary


def _check_importance(importance: str) -> str:
    if importance not in IMPORTANCE_VALUES:
        raise ValueError(f"importance must be one of {sorted(IMPORTANCE_VALUES)}, got {importance!r}")
    return importance


def draft_create(
    to: list[str],
    subject: str,
    body: str,
    cc: list[str] | None = None,
    bcc: list[str] | None = None,
    body_format: EmailBodyFormat = "markdown",
    importance: str = "Normal",
    attachments: list[str] | None = None,
) -> JsonDict:
    """Create a draft from Markdown (default) or HTML, with the Aptos style and saved signature. Markdown supports real bullet lists, numbered lists, emphasis and links. Does not send."""
    html = is_html_format(body_format)
    encoded = encode_attachments(attachments)
    recipients = parse_recipients(to)
    if not recipients:
        raise ValueError("At least one recipient is required in 'to'")
    if not body.strip():
        raise ValueError("body must not be empty")
    html_body = compose_email_html(
        body,
        is_html=html,
        signature_html=load_signature(config.SIGNATURE_NEW_FILE),
    )
    with outlook_client() as client:
        draft = client.create_draft(
            subject=subject,
            body=html_body,
            to=recipients,
            cc=parse_recipients(cc) or None,
            bcc=parse_recipients(bcc) or None,
            content_type="HTML",
            importance=_check_importance(importance),
        )
        attach_to_draft(client, draft, encoded)
    return {**draft_summary(draft), "bodyFormat": body_format}


def draft_reply(
    message_id: str, body: str, reply_all: bool = False, body_format: EmailBodyFormat = "markdown",
    attachments: list[str] | None = None, importance: str | None = None,
) -> JsonDict:
    """Create a reply draft from Markdown (default) or HTML. Preserves the reply signature and quoted original. Markdown dash lists become native bullets. Does not send."""
    html = is_html_format(body_format)
    if not body.strip():
        raise ValueError("body must not be empty")
    encoded = encode_attachments(attachments)
    signature = load_signature(config.SIGNATURE_REPLY_FILE)
    fields = {"importance": _check_importance(importance)} if importance is not None else {}
    with outlook_client() as client:
        draft = client.create_reply_draft(message_id, reply_all=reply_all)
        draft = finish_quoted_draft(client, draft, body, html, signature, **fields)
        attach_to_draft(client, draft, encoded)
    return {**draft_summary(draft), "bodyFormat": body_format}


def draft_list(count: int = 20) -> list[JsonDict]:
    """List drafts in the Drafts folder, newest first. Read-only."""
    with outlook_client() as client:
        drafts = client.list_drafts(top=clamp(count))
    return [draft_summary(d) for d in drafts]


def draft_get(draft_id: str, include_html: bool = False) -> JsonDict:
    """Get one draft by full ID, including recipients and body text. Read-only."""
    with outlook_client() as client:
        draft = client.get_draft(draft_id)
    result = {
        **draft_summary(draft),
        "cc": addresses(draft.get("CcRecipients")),
        "bcc": addresses(draft.get("BccRecipients")),
        "body": body_text(draft.get("Body")),
        "isDraft": draft.get("IsDraft", False),
    }
    if include_html:
        result["bodyHtml"] = (draft.get("Body") or {}).get("Content", "")
    return result


def draft_delete(draft_id: str) -> JsonDict:
    """Delete a draft by full ID. Refuses to delete anything that is not a draft."""
    with outlook_client() as client:
        draft = client.get_draft(draft_id)
        if not draft.get("IsDraft", False):
            raise ValueError("Refusing to delete: this message is not a draft.")
        client.delete_draft(draft_id)
    return {"id": draft_id, "deleted": True}


def draft_update(
    draft_id: str, subject: str | None = None, body: str | None = None,
    to: list[str] | None = None, cc: list[str] | None = None,
    bcc: list[str] | None = None, body_format: EmailBodyFormat = "markdown",
    importance: str | None = None, signature: str = "none",
    attachments: list[str] | None = None,
) -> JsonDict:
    """Edit a draft with a replacement body in Markdown (default) or HTML. Omitted fields stay unchanged; empty fields clear. Does not send."""
    html = is_html_format(body_format)
    encoded = encode_attachments(attachments)
    fields = draft_update_fields(subject, body, to, cc, bcc, html, importance, signature)
    with outlook_client() as client:
        draft = client.update_mail_draft(draft_id, **fields)
        attach_to_draft(client, draft, encoded)
    return {**draft_summary(draft), **({"bodyFormat": body_format} if body is not None else {})}


def draft_forward(
    message_id: str, to: list[str], body: str = "", cc: list[str] | None = None,
    bcc: list[str] | None = None, body_format: EmailBodyFormat = "markdown", attachments: list[str] | None = None,
    importance: str | None = None,
) -> JsonDict:
    """Create a forward draft from Markdown (default) or HTML. Preserves the reply signature and quoted original. Does not send."""
    html = is_html_format(body_format)
    encoded = encode_attachments(attachments)
    recipients = parse_recipients(to)
    if not recipients:
        raise ValueError("At least one recipient is required in 'to'")
    signature = load_signature(config.SIGNATURE_REPLY_FILE)
    fields = {"importance": _check_importance(importance)} if importance is not None else {}
    with outlook_client() as client:
        draft = client.create_forward_draft(message_id)
        draft = finish_quoted_draft(
            client, draft, body, html, signature, to=recipients,
            cc=parse_recipients(cc) or None, bcc=parse_recipients(bcc) or None, **fields,
        )
        attach_to_draft(client, draft, encoded)
    return {**draft_summary(draft), "bodyFormat": body_format}


TOOLS = [draft_create, draft_reply, draft_update, draft_forward, draft_list, draft_get, draft_delete]
