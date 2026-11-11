"""Shared draft edits and local attachment validation."""
from __future__ import annotations

import base64
import mimetypes
import stat
from pathlib import Path

from . import config
from .errors import OutlookAPIError
from .formatting import compose_email_html, parse_recipients, text_to_html
from .signatures import load_signature

MAX_ATTACHMENT_BYTES = 3 * 1024 * 1024


def encode_attachment(path: str | Path) -> dict:
    """Read one regular local file of at most 3 MiB."""
    if not str(path).strip():
        raise ValueError("Attachment path must not be empty.")
    source = Path(path).expanduser()
    try:
        if not stat.S_ISREG(source.stat().st_mode):
            raise ValueError(f"Attachment must be a regular file: {source}")
        with source.open("rb") as stream:
            data = stream.read(MAX_ATTACHMENT_BYTES + 1)
    except OSError as exc:
        raise ValueError(f"Cannot read attachment {source}: {exc}") from exc
    if len(data) > MAX_ATTACHMENT_BYTES:
        raise ValueError(f"Attachment exceeds 3 MiB: {source}")
    return {
        "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
        "Name": source.name,
        "ContentType": mimetypes.guess_type(source.name)[0] or "application/octet-stream",
        "ContentBytes": base64.b64encode(data).decode("ascii"),
    }


def encode_attachments(paths: list[str] | None) -> list[dict]:
    """Validate all files before any remote write."""
    return [encode_attachment(path) for path in paths or []]


def attach_to_draft(client, draft: dict, attachments: list[dict]) -> dict:
    """Add prevalidated files without retries or draft removal."""
    if not attachments:
        return draft
    draft_id = draft.get("Id", "")
    if not draft_id:
        raise OutlookAPIError(0, "Cannot attach files: the draft response has no ID.")
    for index, attachment in enumerate(attachments):
        try:
            client._add_encoded_draft_attachment(draft_id, attachment)
        except Exception as exc:
            raise OutlookAPIError(
                0, f"Draft {draft_id}: attachment {index + 1} failed after {index} completed uploads. "
                "The last upload outcome may be uncertain. Check the draft before any retry. "
                f"The draft remains saved. {exc}",
            ) from exc
    return draft


def compose_quoted_body(draft: dict, body: str, html: bool, signature: str) -> str:
    """Put new text before the signature and original message."""
    quoted = draft.get("Body", {}).get("Content", "")
    if draft.get("Body", {}).get("ContentType") == "Text":
        quoted = text_to_html(quoted)
    return compose_email_html(body, is_html=html, signature_html=signature, quoted_html=quoted)


def finish_quoted_draft(client, draft: dict, body: str, html: bool, signature: str, **fields) -> dict:
    """Complete a reply or forward while preserving the original message."""
    draft_id = draft.get("Id", "")
    if not draft_id:
        raise OutlookAPIError(0, "The draft response has no ID.")
    try:
        return client.update_draft(
            draft_id, body=compose_quoted_body(draft, body, html, signature),
            content_type="HTML", **fields,
        )
    except Exception as exc:
        raise OutlookAPIError(0, f"Draft {draft_id} remains saved; its update failed: {exc}") from exc


def draft_update_fields(subject, body, to, cc, bcc, html, importance, signature) -> dict:
    """Prepare explicit edits without appending a signature to an unchanged body."""
    if signature not in {"none", "new", "reply"}:
        raise ValueError("signature must be none, new or reply")
    fields = {"subject": subject, "importance": importance}
    if importance is not None and importance not in {"Low", "Normal", "High"}:
        raise ValueError("importance must be Low, Normal or High")
    for name, values in (("to", to), ("cc", cc), ("bcc", bcc)):
        if values is not None:
            fields[name] = parse_recipients(values)
    if body is not None:
        if signature != "none":
            path = config.SIGNATURE_REPLY_FILE if signature == "reply" else config.SIGNATURE_NEW_FILE
            saved_signature = load_signature(path)
            signature_html = "" if html and saved_signature in body else saved_signature
            body = compose_email_html(body, is_html=html, signature_html=signature_html)
            html = True
        fields.update(body=body, content_type="HTML" if html else "Text")
    return {key: value for key, value in fields.items() if value is not None}


class MailComposeMixin:
    """Outlook REST operations for draft edits, forwards and attachments."""

    def _add_encoded_draft_attachment(self, message_id: str, attachment: dict) -> dict:
        draft = self.get_draft(message_id)
        if not draft.get("IsDraft", False):
            raise ValueError("Refusing to attach: this message is not a draft.")
        response = self._request(
            "POST", f"/me/messages/{message_id}/attachments",
            json_body=attachment, max_retries=0,
        )
        return response.json()

    def add_draft_attachment(self, message_id: str, path: str | Path) -> dict:
        """Attach one local file to an existing draft. Does not send."""
        return self._add_encoded_draft_attachment(message_id, encode_attachment(path))

    def create_forward_draft(self, message_id: str) -> dict:
        """Create a forward draft without an automatic retry. Does not send."""
        response = self._request(
            "POST", f"/me/messages/{message_id}/createForward", json_body={}, max_retries=0,
        )
        return response.json()

    def update_mail_draft(
        self, message_id: str, *, subject: str | None = None, body: str | None = None,
        content_type: str | None = None, to: list[str] | None = None,
        cc: list[str] | None = None, bcc: list[str] | None = None,
        importance: str | None = None,
    ) -> dict:
        """Update only supplied fields after a draft check. Empty values clear fields."""
        draft = self.get_draft(message_id)
        if not draft.get("IsDraft", False):
            raise ValueError("Refusing to update: this message is not a draft.")
        fields = {key: value for key, value in {
            "subject": subject, "body": body, "content_type": content_type,
            "to": to, "cc": cc, "bcc": bcc, "importance": importance,
        }.items() if value is not None}
        if importance is not None and importance not in {"Low", "Normal", "High"}:
            raise ValueError("importance must be Low, Normal or High")
        if content_type is not None and content_type not in {"Text", "HTML"}:
            raise ValueError("content_type must be Text or HTML")
        return self.update_draft(message_id, **fields) if fields else draft
