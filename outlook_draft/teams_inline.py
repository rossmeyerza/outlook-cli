"""Identify images that Teams hosts for a chat message."""

from html.parser import HTMLParser
from urllib.parse import quote, unquote, urlsplit

from .errors import CliError


def _id_segment(value: str) -> str:
    if not value or value in {".", ".."} or any(char.isspace() or ord(char) < 32 for char in value):
        raise CliError("Invalid Teams ID.")
    return quote(value, safe="")


def message_path(chat_id: str, message_id: str) -> str:
    return f"/chats/{_id_segment(chat_id)}/messages/{_id_segment(message_id)}"


def image_path(chat_id: str, message_id: str, image_id: str) -> str:
    return f"{message_path(chat_id, message_id)}/hostedContents/{_id_segment(image_id)}/$value"


class _ImageReferences(HTMLParser):
    def __init__(self, chat_id: str, message_id: str):
        super().__init__()
        self.chat_id = chat_id
        self.message_id = message_id
        self.ids: dict[str, None] = {}

    def handle_starttag(self, tag, attrs):
        src = dict(attrs).get("src") if tag == "img" else None
        if not src:
            return
        try:
            url = urlsplit(src)
        except ValueError:
            return
        if (url.scheme or url.netloc) and (url.scheme, url.netloc.lower()) != ("https", "graph.microsoft.com"):
            return
        parts = [unquote(part) for part in url.path.lstrip("./").split("/")]
        if parts[:1] in (["v1.0"], ["beta"]):
            parts = parts[1:]
        if parts[:4] == ["chats", self.chat_id, "messages", self.message_id]:
            parts = parts[4:]
        elif url.netloc:
            return
        if len(parts) == 3 and parts[0] == "hostedContents" and parts[1] and parts[2] == "$value":
            self.ids[parts[1]] = None


def inline_image_ids(message: dict, chat_id: str) -> list[str]:
    body = message.get("body") or {}
    if (body.get("contentType") or "").lower() != "html" or message.get("deletedDateTime"):
        return []
    parser = _ImageReferences(chat_id, message.get("id", ""))
    parser.feed(body.get("content") or "")
    parser.close()
    return list(parser.ids)


def is_supported_image(data: bytes) -> bool:
    return (
        data.startswith((b"\x89PNG\r\n\x1a\n", b"\xff\xd8\xff", b"GIF87a", b"GIF89a", b"BM"))
        or (data.startswith(b"RIFF") and data[8:12] == b"WEBP")
    )
