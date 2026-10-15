"""Pure text and email formatting helpers shared by the CLI and the MCP server."""

from __future__ import annotations

import logging
import re
from datetime import datetime
from html import escape

from . import config

log = logging.getLogger(__name__)


def html_to_text(html: str) -> str:
    """Simple HTML to plain text conversion."""
    # Remove style and script blocks
    text = re.sub(r"<style[^>]*>.*?</style>", "", html, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r"<script[^>]*>.*?</script>", "", text, flags=re.DOTALL | re.IGNORECASE)
    # Convert <br> and block elements to newlines
    text = re.sub(r"<br\s*/?>\s*", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"</(p|div|tr|li|h[1-6])>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<(p|div|tr|li|h[1-6])[^>]*>", "", text, flags=re.IGNORECASE)
    # Strip remaining tags
    text = re.sub(r"<[^>]+>", "", text)
    # Decode common HTML entities
    text = text.replace("&nbsp;", " ").replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">").replace("&quot;", '"')
    # Collapse whitespace
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def format_datetime(value: str) -> str:
    """Format an ISO timestamp for terminal output."""
    if not value:
        return ""
    try:
        parsed = datetime.fromisoformat(value.replace("Z", "+00:00"))
        return parsed.strftime("%b %d %H:%M")
    except Exception:
        log.debug("Best-effort operation failed", exc_info=True)
        return value[:16]


def text_to_html(text: str) -> str:
    """Convert plain text to simple HTML paragraphs."""
    text = text.replace("\r\n", "\n").strip()
    if not text:
        return "<p></p>"

    parts = []
    for block in text.split("\n\n"):
        block = block.strip()
        if not block:
            continue
        block_html = escape(block).replace("\n", "<br>\n")
        parts.append(f"<p>{block_html}</p>")
    return "".join(parts) or "<p></p>"


def compose_email_html(
    body: str,
    *,
    is_html: bool,
    signature_html: str,
    quoted_html: str = "",
) -> str:
    """Build the final HTML body with Aptos styling and signature."""
    content_html = body if is_html else text_to_html(body)
    style = (
        f"font-family: {config.EMAIL_FONT_FAMILY}; "
        f"font-size: {config.EMAIL_FONT_SIZE}; "
        "color: rgb(0, 0, 0);"
    )
    return f'<div style="{style}">{content_html}</div>{signature_html}{quoted_html}'


def parse_recipients(values: list[str] | None) -> list[str]:
    """Split repeatable and comma-separated address values into a flat list."""
    if not values:
        return []
    return [email.strip() for value in values for email in value.split(",") if email.strip()]
