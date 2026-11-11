"""Email composition accepts Markdown or HTML, with explicit block styles."""
from __future__ import annotations

import asyncio
from html.parser import HTMLParser
from unittest.mock import Mock

import pytest
from mcp_fakes import FakeClient, use_client

from outlook_draft import cli
from outlook_draft.commands import drafts as commands
from outlook_draft.email_markdown import is_html_format, markdown_to_email_html
from outlook_draft.formatting import compose_email_html, text_to_html
from outlook_draft.mail_compose import draft_update_fields
from outlook_draft.mcp import drafts, sending
from outlook_draft.mcp.server import build_server


class Elements(HTMLParser):
    def __init__(self, content):
        super().__init__()
        self.elements = []
        self.feed(content)

    def handle_starttag(self, tag, attrs):
        self.elements.append((tag, dict(attrs)))


def tags(content, name):
    return [attrs for tag, attrs in Elements(content).elements if tag == name]


def test_real_bullets_and_inline_markdown():
    content = markdown_to_email_html("Hi HP,\n\n- **Content:** the proposal\n- CCM: [signal report](https://example.com/report)")
    assert len(tags(content, "ul")) == 1
    assert len(tags(content, "li")) == 2
    assert "<strong>Content:</strong>" in content
    assert tags(content, "a")[0]["href"] == "https://example.com/report"
    assert all("margin: 0;" in item["style"] for item in tags(content, "li"))


def test_loose_list_paragraphs_have_no_extra_margins():
    content = markdown_to_email_html("- First bullet\n\n- Second bullet\n\n  Second paragraph\n\n  - Nested bullet")
    assert len(tags(content, "ul")) == 2
    assert len(tags(content, "li")) == 3
    assert tags(content, "p")
    assert all(item["style"] == "margin: 0; line-height: 1.2;" for item in tags(content, "p"))


def test_numbered_lists_breaks_and_unicode():
    content = markdown_to_email_html("Hello,\r\nNext line\r\n\r\n1. Café\r\n2. Done")
    assert len(tags(content, "br")) == 1
    assert len(tags(content, "ol")) == 1
    assert len(tags(content, "li")) == 2
    assert "Café" in content


def test_raw_html_and_unsafe_links_do_not_execute_in_markdown():
    content = markdown_to_email_html('<script>alert(1)</script>\n\n[bad](javascript:alert(1))')
    assert not tags(content, "script")
    assert not any(item.get("href", "").startswith("javascript:") for item in tags(content, "a"))
    assert "&lt;script&gt;" in content


def test_paragraph_spacing_is_explicit_and_not_one_em():
    content = compose_email_html("First paragraph.\n\nSecond paragraph.", is_html=False, signature_html="")
    assert all(item["style"] == "margin: 0 0 8px 0; line-height: 1.2;" for item in tags(content, "p"))
    assert "line-height: 1.2" in tags(content, "div")[0]["style"]
    assert "1em" not in content


def test_html_signature_and_quoted_original_are_preserved():
    body = '<p style="margin:18px;line-height:2">Custom HTML</p>'
    signature = '<div style="margin:7px">Signature</div>'
    quoted = '<blockquote style="line-height:2.5"><p>Original</p></blockquote>'
    result = compose_email_html(body, is_html=True, signature_html=signature, quoted_html=quoted)
    assert body in result
    assert result.endswith(signature + quoted)
    assert text_to_html("*literal quote*") == "<p>*literal quote*</p>"


def test_draft_update_renders_markdown_but_does_not_duplicate_signature():
    result = draft_update_fields(None, "- One\n- Two", None, None, None, False, None, "none")
    assert result["content_type"] == "HTML"
    assert len(tags(result["body"], "li")) == 2
    assert draft_update_fields("Subject", None, None, None, None, False, None, "none") == {"subject": "Subject"}
    assert draft_update_fields(None, "", None, None, None, False, None, "none")["body"] == ""


@pytest.mark.parametrize("value", ["text", "Markdown", "", None])
def test_invalid_format_fails(value):
    with pytest.raises(ValueError, match="markdown or html"):
        is_html_format(value)


@pytest.fixture
def fake(monkeypatch):
    client = FakeClient(create_draft={"Id": "d1"}, create_reply_draft={"Id": "d1", "Body": {"ContentType": "HTML", "Content": "<p>Original</p>"}},
                        create_forward_draft={"Id": "d1", "Body": {"ContentType": "HTML", "Content": "<p>Original</p>"}},
                        update_draft={"Id": "d1"}, update_mail_draft={"Id": "d1"})
    use_client(monkeypatch, drafts, "outlook_client", client)
    use_client(monkeypatch, sending, "outlook_client", client)
    monkeypatch.setattr(drafts, "load_signature", lambda path: "<div>Signature</div>")
    monkeypatch.setattr(sending, "load_signature", lambda path: "<div>Signature</div>")
    return client


@pytest.mark.parametrize("operation, method", [("create", "create_draft"), ("reply", "update_draft"),
                                                 ("forward", "update_draft"), ("update", "update_mail_draft")])
def test_all_draft_entry_points_use_markdown_by_default(fake, operation, method):
    body = "- One\n- Two"
    if operation == "create":
        result = drafts.draft_create(["a@example.test"], "Subject", body)
    elif operation == "reply":
        result = drafts.draft_reply("m1", body)
    elif operation == "forward":
        result = drafts.draft_forward("m1", ["a@example.test"], body)
    else:
        result = drafts.draft_update("d1", body=body)
    assert result["bodyFormat"] == "markdown"
    assert len(tags(fake.call(method)[1]["body"], "li")) == 2
    assert fake.call(method)[1]["content_type"] == "HTML"
    assert "send_message" not in fake.names()


@pytest.mark.parametrize("confirm", [False, True])
def test_send_uses_same_markdown_output_in_preview_and_confirm(fake, confirm):
    result = sending.mail_send(["a@example.test"], "Subject", "- One\n- Two", confirm=confirm)
    assert result["bodyFormat"] == "markdown"
    assert len(tags(result["bodyHtml"], "li")) == 2
    if confirm:
        assert fake.call("send_mail")[1]["body"] == result["bodyHtml"]
        assert fake.call("send_mail")[1]["content_type"] == "HTML"
    else:
        assert "send_mail" not in fake.names()


@pytest.mark.parametrize("command", [
    ["draft", "create", "--to", "a@example.test", "-s", "S"],
    ["draft", "reply", "m1"], ["draft", "forward", "m1", "--to", "a@example.test"],
    ["draft", "update", "d1"], ["mail", "send", "--to", "a@example.test", "-s", "S"],
])
def test_cli_advertises_two_formats_and_keeps_old_html_alias(command):
    parser = cli.build_parser()
    assert parser.parse_args(command).body_format == "markdown"
    assert parser.parse_args([*command, "--format", "html"]).body_format == "html"
    assert parser.parse_args([*command, "--html"]).body_format == "html"
    with pytest.raises(SystemExit):
        parser.parse_args([*command, "--format", "text"])


def test_mcp_agent_schemas_advertise_markdown_and_html_only():
    tools = {tool.name: tool for tool in asyncio.run(build_server().list_tools())}
    for name in ("draft_create", "draft_reply", "draft_update", "draft_forward", "mail_send"):
        properties = tools[name].inputSchema["properties"]
        assert properties["body_format"]["enum"] == ["markdown", "html"]
        assert properties["body_format"]["default"] == "markdown"
        assert "html" not in properties
        assert "Markdown" in tools[name].description
    assert "html" in tools["teams_send"].inputSchema["properties"]


def test_invalid_format_does_not_authenticate(monkeypatch):
    monkeypatch.setattr(drafts, "outlook_client", Mock(side_effect=AssertionError("No client")))
    with pytest.raises(ValueError):
        drafts.draft_reply("m1", "Body", body_format="text")


def test_cli_send_without_signature_still_renders_markdown(monkeypatch):
    client = Mock()
    monkeypatch.setattr(commands.runtime, "get_client", lambda: client)
    args = cli.build_parser().parse_args(["mail", "send", "--to", "a@example.test", "-s", "S", "-b", "- One\n- Two", "--yes"])
    args.func(args)
    kwargs = client.send_mail.call_args.kwargs
    assert len(tags(kwargs["body"], "li")) == 2
    assert kwargs["content_type"] == "HTML"
