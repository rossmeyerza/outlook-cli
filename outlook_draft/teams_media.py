"""Validated Teams files and inline images with explicit send boundaries."""
from __future__ import annotations

import base64
import html
import json
import mimetypes
import os
import stat
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from urllib.parse import quote, urlsplit

from . import config
from .errors import OutlookDraftError
from .safety import check_request

DEFAULT_UPLOAD_FOLDER = "Microsoft Teams Chat Files"
MAX_FILE_BYTES = 25 * 1024 * 1024
MAX_IMAGE_BYTES = 1024 * 1024
MAX_TOTAL_BYTES = 50 * 1024 * 1024
MAX_IMAGE_TOTAL = 2 * 1024 * 1024
MAX_BODY_BYTES = 28000
SELF_CHAT_ID = "48:notes"


@dataclass(frozen=True)
class MediaFile:
    name: str
    content_type: str
    data: bytes = field(repr=False)

    def summary(self) -> dict:
        return {"name": self.name, "contentType": self.content_type, "size": len(self.data)}


@dataclass(frozen=True)
class PreparedMedia:
    files: tuple[MediaFile, ...] = ()
    images: tuple[MediaFile, ...] = ()

    def __bool__(self) -> bool:
        return bool(self.files or self.images)

    def summary(self) -> dict:
        return {"attachments": [item.summary() for item in self.files],
                "images": [item.summary() for item in self.images]}


class MediaSendError(OutlookDraftError):
    """A media operation failed after a remote write could have occurred."""

    def __init__(self, chat_id: str, phase: str, uploads: list[dict], error: Exception):
        self.chat_id, self.phase, self.uploads = chat_id, phase, uploads
        super().__init__(f"Teams media operation stopped during {phase} for chat {chat_id}: {error}. "
                         f"Remote files: {json.dumps(uploads)}. The last write may have an uncertain result. "
                         "No automatic retry or cleanup occurred. Check the chat and files before any retry.")


def _read_file(value: str, limit: int) -> tuple[Path, bytes]:
    if not isinstance(value, str) or not value.strip():
        raise ValueError("A media path must be a non-empty local path")
    path = Path(value).expanduser()
    if any(char in path.name for char in '\\/:*?"<>|') or any(ord(char) < 32 for char in path.name):
        raise ValueError(f"Unsupported media filename: {path.name!r}")
    try:
        descriptor = os.open(path, os.O_RDONLY | os.O_NONBLOCK | os.O_NOFOLLOW)
        with os.fdopen(descriptor, "rb") as stream:
            if not stat.S_ISREG(os.fstat(stream.fileno()).st_mode):
                raise ValueError(f"Media must be a regular file, not a device or pipe: {path}")
            data = stream.read(limit + 1)
    except OSError as exc:
        raise ValueError(f"Cannot read media file {path}: {exc}") from exc
    if len(data) > limit:
        raise ValueError(f"Media file exceeds {limit // (1024 * 1024)} MiB: {path}")
    return path, data


def _image(value: str) -> MediaFile:
    path, data = _read_file(value, MAX_IMAGE_BYTES)
    signatures = [(b"\x89PNG\r\n\x1a\n", "image/png", {".png"}),
                  (b"\xff\xd8\xff", "image/jpeg", {".jpg", ".jpeg"}),
                  (b"GIF87a", "image/gif", {".gif"}), (b"GIF89a", "image/gif", {".gif"})]
    for signature, mime, suffixes in signatures:
        if data.startswith(signature) and path.suffix.lower() in suffixes:
            return MediaFile(path.name, mime, data)
    raise ValueError(f"Inline images must be PNG, JPEG or GIF files with a matching extension: {path}")


def prepare_media(attachments: list[str] | None = None, images: list[str] | None = None) -> PreparedMedia:
    if len(attachments or []) + len(images or []) > 10:
        raise ValueError("At most 10 files and inline images are supported per message")
    files = []
    for value in attachments or []:
        path, data = _read_file(value, MAX_FILE_BYTES)
        files.append(MediaFile(path.name, mimetypes.guess_type(path.name)[0] or "application/octet-stream", data))
        if sum(len(item.data) for item in files) > MAX_TOTAL_BYTES:
            raise ValueError("The total size of local media exceeds 50 MiB")
    media = PreparedMedia(tuple(files), tuple(_image(value) for value in images or []))
    if sum(len(item.data) for item in media.files + media.images) > MAX_TOTAL_BYTES:
        raise ValueError("The total size of local media exceeds 50 MiB")
    if sum(len(item.data) for item in media.images) > MAX_IMAGE_TOTAL:
        raise ValueError("The total size of inline images exceeds 2 MiB")
    return media


def validate_content(content: str, media: PreparedMedia, *, is_html: bool = False) -> None:
    if not content.strip() and not media:
        raise ValueError("The message must not be empty unless a file or inline image is supplied")
    if len(content.encode("utf-8")) > 20000:
        raise ValueError("The Teams message text exceeds the 20,000-byte client limit")
    if media and len(_body(content, media, ["0" * 36] * len(media.files), is_html).encode("utf-8")) > MAX_BODY_BYTES:
        raise ValueError("The final Teams HTML body exceeds the 28,000-byte client limit")


def _folder_path(value: str) -> str:
    parts = value.split("/")
    if not value.strip() or any(not part.strip() or part in {".", ".."} for part in parts):
        raise ValueError("Use an existing relative OneDrive folder path without dot segments")
    return "/".join(quote(part, safe="") for part in parts)


def resolve_upload_folder(client, path: str) -> dict:
    encoded = _folder_path(path)
    folder = client._request("GET", f"/me/drive/root:/{encoded}",
                             params={"$select": "id,name,folder,webUrl,shared"}).json()
    if not folder.get("id") or "folder" not in folder:
        raise ValueError("The upload destination is not an existing OneDrive folder")
    if folder.get("shared"):
        raise ValueError("Use a private OneDrive folder; this folder already has shared access")
    return folder


def _member_ids(me: dict, members: list[dict], tenant_id: str | None) -> list[str]:
    if len(members) >= 1001:
        raise ValueError("Cannot verify all chat members: the client limit is 1000")
    own_tenant = next((member.get("tenantId") for member in members if member.get("userId") == me.get("id")), None)
    tenant_id = tenant_id or own_tenant
    if not tenant_id or (own_tenant and own_tenant != tenant_id):
        raise ValueError("Cannot verify the tenant of every chat member")
    ids = set()
    for member in members:
        kind = member.get("@odata.type", "").lower()
        if kind and not kind.endswith(".aaduserconversationmember"):
            raise ValueError("Files can only be shared with identifiable internal user members")
        if "guest" in member.get("roles", []) or (tenant_id and member.get("tenantId") != tenant_id):
            raise ValueError("File attachments for guests or external chat members are not supported")
        value = member.get("userId", "")
        try:
            uuid.UUID(value)
        except (ValueError, TypeError, AttributeError) as exc:
            raise ValueError("Cannot identify a chat member safely for file access") from exc
        ids.add(value)
    if me.get("id") not in ids:
        raise ValueError("The current user is not a verified member of this chat")
    return sorted(ids - {me["id"]})


def file_audience(client, chat_id: str) -> list[str]:
    me = client.get_current_user()
    if chat_id == SELF_CHAT_ID:
        client.list_teams_message_metadata(SELF_CHAT_ID, top=1)
        return []
    chat = client.get_teams_chat(chat_id)
    if chat.get("chatType") not in {"oneOnOne", "group"}:
        raise ValueError("File attachments currently support only 1:1 chats and group chats")
    members = client.list_teams_chat_members(chat_id, top=1001)
    return _member_ids(me, members, chat.get("tenantId"))


def media_plan(client, chat_id: str, content: str, media: PreparedMedia,
               upload_folder: str = DEFAULT_UPLOAD_FOLDER, *, is_html: bool = False) -> dict:
    validate_content(content, media, is_html=is_html)
    result = {"chat_id": chat_id, "content": content, "sent": False, **media.summary()}
    if media.files:
        folder = resolve_upload_folder(client, upload_folder)
        result.update(uploadFolder={"id": folder["id"], "path": upload_folder},
                      accessUserIds=file_audience(client, chat_id))
    return result


def _upload(client, folder_id: str, file: MediaFile, uploads: list[dict]) -> dict:
    remote_name = f"{uuid.uuid4().hex}-{file.name}"
    artifact = {"name": file.name, "remoteName": remote_name, "folderId": folder_id, "status": "uncertain"}
    uploads.append(artifact)
    path = f"/me/drive/items/{quote(folder_id, safe='')}:/{quote(remote_name, safe='')}:/content"
    item = client._request("PUT", path, data=file.data, extra_headers={"Content-Type": file.content_type}, max_retries=0).json()
    if item.get("id"):
        artifact.update(id=item["id"], status="uploaded")
    if not item.get("id") or not item.get("webUrl"):
        raise ValueError("The upload response has no file ID or URL")
    url = urlsplit(item["webUrl"])
    if url.scheme != "https" or url.username or url.password or not (url.hostname or "").endswith(".sharepoint.com"):
        raise ValueError("The upload response has no trusted SharePoint file URL")
    artifact.update(id=item["id"], webUrl=item["webUrl"], status="uploaded")
    metadata = client._request("GET", f"/me/drive/items/{quote(item['id'], safe='')}",
                               params={"$select": "id,eTag,webDavUrl,sharepointIds"}).json()
    reference_url = metadata.get("webDavUrl", "")
    reference = urlsplit(reference_url)
    if reference.scheme != "https" or reference.netloc != url.netloc or reference.username or reference.password:
        raise ValueError("The uploaded file has no trusted webDavUrl for a Teams reference")
    try:
        attachment_id = str(uuid.UUID(metadata.get("eTag", "").split(",")[0].strip('\"{}')))
    except ValueError as exc:
        raise ValueError("The uploaded file has no SharePoint GUID in its ETag") from exc
    return {**item, "referenceUrl": reference_url, "attachmentId": attachment_id}


def _grant_access(client, item_id: str, audience: list[str]) -> None:
    if not audience:
        return
    payload = {"recipients": [{"objectId": user_id} for user_id in audience],
               "requireSignIn": True, "sendInvitation": False, "roles": ["read"], "retainInheritedPermissions": True}
    response = client._request("POST", f"/me/drive/items/{quote(item_id, safe='')}/invite",
                               json_body=payload, max_retries=0)
    if response.status_code != 200:
        raise ValueError("The file access request did not confirm every recipient; the message was not sent")
    permissions = response.json().get("value", [])
    granted = set()
    for permission in permissions:
        identities = permission.get("grantedToIdentitiesV2", []) + permission.get("grantedToIdentities", [])
        identities += [permission.get("grantedToV2", {}), permission.get("grantedTo", {})]
        granted.update(identity.get("user", {}).get("id") for identity in identities)
    if any(permission.get("error") for permission in permissions) or not set(audience).issubset(granted):
        raise ValueError("The file access response has missing or failed permissions; the message was not sent")


def _body(content: str, media: PreparedMedia, attachment_ids: list[str], is_html: bool) -> str:
    body = content if is_html else html.escape(content).replace("\n", "<br>")
    for attachment_id in attachment_ids:
        body += f'<attachment id="{attachment_id}"></attachment>'
    for index, file in enumerate(media.images, 1):
        body += f'<p><img src="../hostedContents/{index}/$value" alt="{html.escape(file.name, quote=True)}"></p>'
    return body


def _payload(content: str, media: PreparedMedia, items: list[dict], is_html: bool) -> dict:
    body = _body(content, media, [item["attachmentId"] for item in items], is_html)
    attachments, hosted = [], []
    for file, item in zip(media.files, items, strict=True):
        attachment_id = item["attachmentId"]
        attachments.append({"id": attachment_id, "contentType": "reference", "contentUrl": item["referenceUrl"], "name": file.name})
    for index, file in enumerate(media.images, 1):
        hosted.append({"@microsoft.graph.temporaryId": str(index), "contentType": file.content_type,
                       "contentBytes": base64.b64encode(file.data).decode("ascii")})
    payload = {"body": {"contentType": "html", "content": body}}
    if attachments:
        payload["attachments"] = attachments
    if hosted:
        payload["hostedContents"] = hosted
    return payload


def send_media(client, chat_id: str, content: str, media: PreparedMedia, *, is_html: bool = False,
               upload_folder: str = DEFAULT_UPLOAD_FOLDER) -> dict:
    if not chat_id.strip():
        raise ValueError("chat_id must not be empty")
    check_request("POST", f"{config.GRAPH_BASE_URL}/chats/{quote(chat_id, safe='')}/messages")
    plan = media_plan(client, chat_id, content, media, upload_folder, is_html=is_html)
    uploads, items, phase = [], [], "upload"
    try:
        for file in media.files:
            phase = "upload"
            check_request("POST", f"{config.GRAPH_BASE_URL}/chats/{quote(chat_id, safe='')}/messages")
            item = _upload(client, plan["uploadFolder"]["id"], file, uploads)
            items.append(item)
            phase = "file access"
            check_request("POST", f"{config.GRAPH_BASE_URL}/chats/{quote(chat_id, safe='')}/messages")
            _grant_access(client, item["id"], plan["accessUserIds"])
        phase = "message send"
        payload = _payload(content, media, items, is_html)
        response = client._request("POST", f"/chats/{quote(chat_id, safe='')}/messages", json_body=payload, max_retries=0).json()
        if not response.get("id"):
            raise ValueError("The message response has no ID")
    except (OutlookDraftError, ValueError) as exc:
        raise MediaSendError(chat_id, phase, uploads, exc) from exc
    return {**plan, "sent": True, "message_id": response["id"], "message": response, "uploads": uploads}
