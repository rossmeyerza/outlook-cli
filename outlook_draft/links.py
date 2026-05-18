from __future__ import annotations

import base64
import html
import re
from typing import Iterable


SHAREPOINT_HOSTS = (
    "sharepoint.com",
    "sharepoint-df.com",
    "my.sharepoint.com",
    "1drv.ms",
)
ONEDRIVE_HOSTS = (
    "onedrive.live.com",
    "my.sharepoint.com",
    "1drv.ms",
)


def encode_share_id(url: str) -> str:
    """Encode a sharing URL as a Graph share id (`u!base64url`)."""
    payload = base64.urlsafe_b64encode(url.encode("utf-8")).decode("ascii").rstrip("=")
    return f"u!{payload}"


def extract_links_from_html(content: str) -> list[dict[str, str]]:
    """Pull <a href> links and <img src> URLs out of an HTML body.

    Returns a list of {url, label, kind} dicts in document order, deduplicated.
    """
    if not content:
        return []

    results: list[dict[str, str]] = []
    seen: set[tuple[str, str]] = set()

    anchor_pattern = re.compile(
        r'<a\b[^>]*href="([^"]+)"[^>]*>(.*?)</a>',
        re.IGNORECASE | re.DOTALL,
    )
    for match in anchor_pattern.finditer(content):
        url = html.unescape(match.group(1)).strip()
        if not url or url.startswith("#"):
            continue
        label = re.sub(r"<[^>]+>", "", match.group(2) or "")
        label = html.unescape(label).strip() or url
        key = ("link", url)
        if key in seen:
            continue
        seen.add(key)
        results.append({"url": url, "label": label, "kind": "link"})

    img_pattern = re.compile(r'<img\b[^>]*src="([^"]+)"', re.IGNORECASE)
    for match in img_pattern.finditer(content):
        url = html.unescape(match.group(1)).strip()
        if not url:
            continue
        key = ("image", url)
        if key in seen:
            continue
        seen.add(key)
        results.append({"url": url, "label": url, "kind": "image"})

    return results


def looks_like_share_url(url: str) -> bool:
    return any(host in url for host in SHAREPOINT_HOSTS + ONEDRIVE_HOSTS)


def filter_share_links(links: Iterable[dict[str, str]]) -> list[dict[str, str]]:
    return [link for link in links if looks_like_share_url(link["url"])]
