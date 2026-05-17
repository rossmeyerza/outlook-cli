from __future__ import annotations

import json
import tempfile
from pathlib import Path

MAIL_CACHE = Path(tempfile.gettempdir()) / "outlook-cli-mail-cache.json"
CAL_CACHE = Path(tempfile.gettempdir()) / "outlook-cli-cal-cache.json"
TASK_CACHE = Path(tempfile.gettempdir()) / "outlook-cli-task-cache.json"
TEAMS_CACHE = Path(tempfile.gettempdir()) / "outlook-cli-teams-cache.json"


def save_cache(path: Path, items: list[dict], *, id_key: str = "Id") -> None:
    """Persist items to disk so index-based refs work in a separate invocation."""
    path.write_text(json.dumps([{"Id": item[id_key]} for item in items if item.get(id_key)]))


def load_cache(path: Path) -> list[dict]:
    """Load cached items."""
    if not path.exists():
        return []
    try:
        return json.loads(path.read_text())
    except Exception:
        return []
