from __future__ import annotations

import pytest

from outlook_draft.commands.notes import (
    _fallback_markdown_to_html,
    _plain_text_to_html,
    _replace_change_from_html,
)


def test_plain_text_to_html_uses_paragraphs_and_breaks() -> None:
    assert _plain_text_to_html("Hello\nthere\n\nNext") == "<p>Hello<br>there</p><p>Next</p>"


def test_fallback_markdown_supports_headings_and_lists() -> None:
    html = _fallback_markdown_to_html("# Title\n\n- one\n- two")

    assert "<h1>Title</h1>" in html
    assert "<li>one</li>" in html
    assert "<li>two</li>" in html


def test_replace_change_requires_exactly_one_match() -> None:
    page_html = """
    <html><head><title>Page</title></head><body>
      <p id="p:{abc}{1}">Status: Draft</p>
      <p id="p:{abc}{2}">Owner: Ross</p>
    </body></html>
    """

    change = _replace_change_from_html(page_html, "Status: Draft", "Status: Final")

    assert change == {
        "target": "p:{abc}{1}",
        "action": "replace",
        "content": "<p>Status: Final</p>",
    }


def test_replace_change_throws_when_missing() -> None:
    with pytest.raises(ValueError, match="did not match"):
        _replace_change_from_html("<p id='p:1'>Hello</p>", "Missing", "New")


def test_replace_change_throws_when_ambiguous() -> None:
    page_html = "<p id='p:1'>Status: Draft</p><p id='p:2'>Status: Draft</p>"

    with pytest.raises(ValueError, match="matched 2 places"):
        _replace_change_from_html(page_html, "Status: Draft", "Status: Final")
