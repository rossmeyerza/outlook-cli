"""Shared validation for MCP changes."""
from __future__ import annotations

import re


def require_id(value: str, name: str) -> str:
    """Require a full ID rather than a cached CLI index."""
    if not value or value in {".", ".."} or value != value.strip() or value.isdecimal() or any(ord(char) < 32 or ord(char) == 127 for char in value):
        raise ValueError(f"{name} must be a non-empty full ID without control characters, not a CLI index.")
    return value


def require_email(value: str) -> str:
    """Reject empty addresses and control characters."""
    if not re.fullmatch(r"[^@\s<>]+@[^@\s<>]+", value):
        raise ValueError("email must contain a valid address")
    return value


def require_name(value: str, name: str = "name") -> str:
    """Require a single file or folder name."""
    if not value.strip() or value in {".", ".."} or any(char in value for char in '/\\:*?"<>|') or any(ord(char) < 32 for char in value):
        raise ValueError(f"{name} must be a non-empty file name without separators or reserved characters.")
    return value
