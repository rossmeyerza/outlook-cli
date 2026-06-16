from __future__ import annotations

import json
from pathlib import Path

from . import config


def _cache_path(name: str) -> Path:
    return config.CACHE_DIR / name


def refresh_paths() -> None:
    global MAIL_CACHE, MAIL_FOLDER_CACHE, CONTACT_CACHE, CAL_CACHE, TASK_CACHE, TEAMS_CACHE
    MAIL_CACHE = _cache_path("mail-cache.json")
    MAIL_FOLDER_CACHE = _cache_path("mail-folder-cache.json")
    CONTACT_CACHE = _cache_path("contact-cache.json")
    CAL_CACHE = _cache_path("cal-cache.json")
    TASK_CACHE = _cache_path("task-cache.json")
    TEAMS_CACHE = _cache_path("teams-cache.json")


refresh_paths()


def save_cache(path: Path, items: list[dict], *, id_key: str = "Id") -> None:
    """Persist items to disk so index-based refs work in a separate invocation."""
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps([{"Id": item[id_key]} for item in items if item.get(id_key)]))


def load_cache(path: Path) -> list[dict]:
    """Load cached items."""
    if not path.exists():
        return []
    try:
        return json.loads(path.read_text())
    except Exception:
        return []
