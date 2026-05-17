from __future__ import annotations

import re
from pathlib import Path


def sanitize_signature_html(html: str) -> str:
    """Remove Outlook editor noise and unsupported blob images from signature HTML."""
    html = html.strip()
    html = re.sub(r"\s+(?:id|class)=\"[^\"]*\"", "", html)
    html = re.sub(r"<img\b[^>]*\bsrc=\"blob:[^\"]*\"[^>]*>", "", html, flags=re.IGNORECASE)
    html = html.replace(
        "var(--darkColor_rgb_0__0__0_, rgb(0, 0, 0))",
        "rgb(0, 0, 0)",
    )
    html = html.replace(
        "var(--darkColor_rgb_255__255__255_, rgb(255, 255, 255))",
        "rgb(255, 255, 255)",
    )
    html = html.replace("var(--darkColor_black, black)", "black")
    html = re.sub(r"<span>\s*</span>", "", html)
    html = re.sub(r">\s+<", "><", html)
    return html


def load_signature(path: Path) -> str:
    """Load and sanitize a saved Outlook signature HTML file."""
    if not path.exists():
        raise FileNotFoundError(f"Signature file not found: {path}")
    return sanitize_signature_html(path.read_text(encoding="utf-8"))
