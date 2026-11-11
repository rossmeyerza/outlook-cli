"""Render Markdown email bodies with explicit, compact block styles."""
from __future__ import annotations

from typing import Literal

from markdown_it import MarkdownIt

EmailBodyFormat = Literal["markdown", "html"]

BLOCK_STYLES = {
    "p": "margin: 0 0 8px 0; line-height: 1.2;",
    "ul": "margin: 0 0 8px 0; padding-left: 24px; line-height: 1.2;",
    "ol": "margin: 0 0 8px 0; padding-left: 24px; line-height: 1.2;",
    "li": "margin: 0; padding: 0; line-height: 1.2;",
    "blockquote": "margin: 0 0 8px 12px; padding-left: 12px; border-left: 2px solid #cccccc;",
    "pre": "margin: 0 0 8px 0; white-space: pre-wrap;",
    "hr": "margin: 8px 0; border: 0; border-top: 1px solid #cccccc;",
}


def is_html_format(body_format: str) -> bool:
    if body_format not in {"markdown", "html"}:
        raise ValueError("body_format must be markdown or html")
    return body_format == "html"


def markdown_to_email_html(body: str) -> str:
    """Render real lists and inline styles without execution of embedded HTML."""
    renderer = MarkdownIt("commonmark", {"html": False, "breaks": True})
    environment: dict = {}
    tokens = renderer.parse(body.replace("\r\n", "\n").replace("\r", "\n"), environment)
    parents: list[str] = []
    for token in tokens:
        if token.nesting == -1:
            if parents:
                parents.pop()
            continue
        style = BLOCK_STYLES.get(token.tag)
        if token.tag == "p" and "li" in parents:
            style = "margin: 0; line-height: 1.2;"
        elif token.tag in {"h1", "h2", "h3", "h4", "h5", "h6"}:
            style = "margin: 0 0 8px 0; line-height: 1.2;"
        if style:
            token.attrSet("style", style)
        if token.nesting == 1:
            parents.append(token.tag)
    return renderer.renderer.render(tokens, renderer.options, environment)
