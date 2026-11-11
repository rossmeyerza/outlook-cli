"""Validate mail API paths before an authenticated request."""
from __future__ import annotations

from urllib.parse import unquote, urlsplit


def validate_mail_url(url: str, base_url: str) -> None:
    """Reject host changes, API escapes and encoded dot segments."""
    if any(ord(char) < 32 or ord(char) == 127 for char in url):
        raise ValueError("Control characters are not permitted in mail API URLs")
    base, target = urlsplit(base_url), urlsplit(url)
    absolute = bool(target.scheme or target.netloc)
    if absolute and (target.scheme != "https" or target.netloc != base.netloc or target.username or target.password):
        raise ValueError("Refusing a continuation URL outside the mail API")
    if target.fragment:
        raise ValueError("Fragments are not permitted in mail API URLs")
    prefix = base.path.rstrip("/") + "/me/" if absolute else "/me/"
    if not target.path.startswith(prefix):
        raise ValueError("Invalid mail API path")
    path = target.path
    for _ in range(6):
        if "\\" in path or any(part in {".", ".."} for part in path.split("/")):
            raise ValueError("Dot segments and backslashes are not permitted in mail API paths")
        decoded = unquote(path)
        if decoded == path:
            return
        path = decoded
    raise ValueError("Excessive encoding in mail API path")
