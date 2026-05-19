from __future__ import annotations

from outlook_draft.links import (
    encode_share_id,
    extract_links_from_html,
    filter_share_links,
    looks_like_share_url,
)


def test_extract_links_from_html_collects_anchors_and_images() -> None:
    html = (
        '<p>See <a href="https://example.com/x">link</a></p>'
        '<img src="https://example.com/img.png">'
        '<a href="https://example.com/x">duplicate</a>'
    )

    links = extract_links_from_html(html)

    assert links == [
        {"url": "https://example.com/x", "label": "link", "kind": "link"},
        {"url": "https://example.com/img.png", "label": "https://example.com/img.png", "kind": "image"},
    ]


def test_filter_share_links_keeps_share_hosts_only() -> None:
    links = extract_links_from_html(
        '<a href="https://example.com/x">x</a>'
        '<a href="https://contoso.sharepoint.com/sites/file.pdf">file</a>'
    )

    filtered = filter_share_links(links)

    assert len(filtered) == 1
    assert filtered[0]["url"].endswith("file.pdf")


def test_looks_like_share_url() -> None:
    assert looks_like_share_url("https://contoso.sharepoint.com/foo")
    assert not looks_like_share_url("https://example.com/foo")


def test_encode_share_id_uses_url_safe_base64_with_u_prefix() -> None:
    share_id = encode_share_id("https://example.com/file.pdf")

    assert share_id.startswith("u!")
    assert "=" not in share_id
