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

OUTLOOK_BASE_URL: str = "https://outlook.office.com/api/v2.0"
GRAPH_BASE_URL: str = "https://graph.microsoft.com/v1.0"
OUTLOOK_TOKEN_DOMAIN: str = "outlook.office.com"
GRAPH_TOKEN_DOMAIN: str = "graph.microsoft.com"
SIGNATURE_NEW_FILE: Path = _project_root / "signature-new.html"
SIGNATURE_REPLY_FILE: Path = _project_root / "signature-reply.html"
EMAIL_FONT_FAMILY: str = (
    "Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif"
)
EMAIL_FONT_SIZE: str = "12pt"

# Token refresh: warn this many seconds before expiry
TOKEN_WARN_BEFORE_EXPIRY: int = 300  # 5 minutes
