from __future__ import annotations

import os
from pathlib import Path

from dotenv import load_dotenv

_project_root = Path(__file__).resolve().parent.parent
load_dotenv(_project_root / ".env")

MS_EMAIL: str = os.environ.get("MS_EMAIL", "")
MS_GRAPH_EXPLORER_DIR: Path = Path(
    os.environ.get("MS_GRAPH_EXPLORER_DIR", str(_project_root.parent / "ms-graph-explorer"))
)
TOKENS_FILE: Path = MS_GRAPH_EXPLORER_DIR / "session_state" / "tokens.json"
AUTH_SCRIPT: Path = MS_GRAPH_EXPLORER_DIR / "auth.py"
AUTH_PYTHON: Path = MS_GRAPH_EXPLORER_DIR / ".venv" / "bin" / "python"

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
