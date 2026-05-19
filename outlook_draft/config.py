from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

_project_root = Path(__file__).resolve().parent.parent
load_dotenv(_project_root / ".env")

MS_EMAIL: str = os.environ.get("MS_EMAIL", "")
MS_PASSWORD: str = os.environ.get("MS_PASSWORD", "")
LOCAL_TIMEZONE: str = os.environ.get("LOCAL_TIMEZONE", "Europe/London")
OUTLOOK_TIMEZONE: str = os.environ.get("OUTLOOK_TIMEZONE", "GMT Standard Time")
SESSION_DIR: Path = _project_root / "session_state"
TOKENS_FILE: Path = SESSION_DIR / "tokens.json"


def resolve_config_path(env_name: str, default: Path) -> Path:
    """Resolve an optional path from .env, relative to the project root if needed."""
    raw = os.environ.get(env_name, "").strip()
    if not raw:
        return default
    path = Path(raw).expanduser()
    if not path.is_absolute():
        path = _project_root / path
    return path


OUTLOOK_BASE_URL: str = "https://outlook.office.com/api/v2.0"
GRAPH_BASE_URL: str = "https://graph.microsoft.com/v1.0"
OUTLOOK_TOKEN_DOMAIN: str = "outlook.office.com"
GRAPH_TOKEN_DOMAIN: str = "graph.microsoft.com"
SIGNATURE_NEW_FILE: Path = resolve_config_path(
    "SIGNATURE_NEW_FILE",
    _project_root / "signature-new.html",
)
SIGNATURE_REPLY_FILE: Path = resolve_config_path(
    "SIGNATURE_REPLY_FILE",
    _project_root / "signature-reply.html",
)
EMAIL_FONT_FAMILY: str = (
    "Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif"
)
EMAIL_FONT_SIZE: str = "12pt"

# Token refresh: warn this many seconds before expiry
TOKEN_WARN_BEFORE_EXPIRY: int = 300  # 5 minutes

# Gateway
GATEWAY_PID_FILE: Path = SESSION_DIR / "gateway.pid"
GATEWAY_LOG_FILE: Path = SESSION_DIR / "gateway.log"
GATEWAY_CHAT_ID_FILE: Path = SESSION_DIR / "gateway_chat_id.txt"
GATEWAY_CHAT_ID: str = os.environ.get("GATEWAY_CHAT_ID", "")
GATEWAY_TRIGGER: str = os.environ.get("GATEWAY_TRIGGER", "@Marlow")
GATEWAY_POLL_INTERVAL: int = int(os.environ.get("GATEWAY_POLL_INTERVAL", "30"))
