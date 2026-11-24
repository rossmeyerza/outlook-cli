"""Compare the installed CLI version with the latest stable GitHub release."""
from __future__ import annotations

import json
import os
import re
import sys
import time
from contextlib import suppress
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from tempfile import NamedTemporaryFile

import httpx

RELEASES_URL = "https://github.com/rossmeyerza/outlook-cli/releases"
LATEST_URL = "https://api.github.com/repos/rossmeyerza/outlook-cli/releases/latest"
SUCCESS_TTL = 24 * 60 * 60
FAILURE_TTL = 60 * 60


def installed_version() -> str:
    try:
        return version("outlook-draft-cli")
    except PackageNotFoundError:
        return "unknown"


def _version_key(value: object) -> tuple[int, ...] | None:
    if not isinstance(value, str) or not re.fullmatch(r"v?[0-9]{1,9}\.[0-9]{1,9}\.[0-9]{1,9}", value):
        return None
    return tuple(int(part) for part in value.removeprefix("v").split("."))


def _fetch_latest() -> str | None:
    try:
        response = httpx.get(LATEST_URL, headers={"Accept": "application/vnd.github+json"}, timeout=1.0)
        response.raise_for_status()
        payload = response.json()
    except (httpx.HTTPError, httpx.InvalidURL, ImportError, ValueError, OSError):
        return None
    if not isinstance(payload, dict) or payload.get("draft") is not False or payload.get("prerelease") is not False:
        return None
    tag = payload.get("tag_name")
    return tag if _version_key(tag) is not None else None


def _read_cache(path: Path, now: float) -> tuple[bool, str | None]:
    try:
        record = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return False, None
    if not isinstance(record, dict) or "tag" not in record:
        return False, None
    tag, checked_at = record["tag"], record.get("checked_at")
    if (tag is not None and _version_key(tag) is None) or not isinstance(checked_at, (int, float)):
        return False, None
    ttl = FAILURE_TTL if tag is None else SUCCESS_TTL
    return now - ttl < checked_at <= now, tag


def _write_cache(path: Path, tag: str | None, now: float) -> None:
    temp_path = None
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(dir=path.parent, mode="w", encoding="utf-8", delete=False) as file:
            temp_path = Path(file.name)
            json.dump({"tag": tag, "checked_at": now}, file)
        temp_path.replace(path)
    except OSError:
        if temp_path is not None:
            with suppress(OSError):
                temp_path.unlink(missing_ok=True)


def notify_if_update_available() -> None:
    if os.environ.get("OUTLOOK_CLI_NO_UPDATE_CHECK") == "1":
        return
    current = installed_version()
    current_key = _version_key(current)
    if current_key is None:
        return
    cache_home = Path(os.environ.get("XDG_CACHE_HOME") or Path.home() / ".cache")
    path = cache_home / "outlook-cli" / "update-check.json"
    fresh, tag = _read_cache(path, time.time())
    if not fresh:
        tag = _fetch_latest()
        _write_cache(path, tag, time.time())
    latest_key = _version_key(tag)
    if latest_key is not None and latest_key > current_key:
        with suppress(OSError):
            print(f"outlook-cli: Update available ({current} to {tag}). See {RELEASES_URL}/latest", file=sys.stderr)
