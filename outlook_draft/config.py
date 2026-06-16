from __future__ import annotations

import os
import re
from pathlib import Path

from dotenv import dotenv_values, load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_DIR = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "outlook-cli"
DATA_DIR = Path(os.environ.get("XDG_DATA_HOME", Path.home() / ".local" / "share")) / "outlook-cli"
ACCOUNTS_CONFIG_DIR: Path = CONFIG_DIR / "accounts"
ACCOUNTS_DATA_DIR: Path = DATA_DIR / "accounts"
ACTIVE_ACCOUNT_FILE: Path = DATA_DIR / "active-account.txt"

CONFIG_FILE = CONFIG_DIR / ".env"
LEGACY_CONFIG_FILE = PROJECT_ROOT / ".env"

load_dotenv(LEGACY_CONFIG_FILE)
load_dotenv(CONFIG_FILE, override=True)

ACTIVE_ACCOUNT: str | None = None
ACCOUNT_DATA_DIR: Path = DATA_DIR
SESSION_DIR: Path = DATA_DIR / "session_state"
CACHE_DIR: Path = DATA_DIR / "cache"
TOKENS_FILE: Path = SESSION_DIR / "tokens.json"
_ACCOUNT_ENV: dict[str, str] = {}


def _valid_account_name(name: str) -> bool:
    return bool(re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]{0,63}", name))


def _read_active_account() -> str | None:
    env_account = os.environ.get("OUTLOOK_CLI_ACCOUNT", "").strip()
    if env_account:
        return env_account
    try:
        value = ACTIVE_ACCOUNT_FILE.read_text().strip()
    except FileNotFoundError:
        return None
    return value or None


def account_env_file(account: str) -> Path:
    if not _valid_account_name(account):
        raise ValueError(
            "Account names may contain letters, numbers, dots, underscores, and dashes."
        )
    return ACCOUNTS_CONFIG_DIR / f"{account}.env"


def list_accounts() -> list[str]:
    if not ACCOUNTS_CONFIG_DIR.exists():
        return []
    return sorted(path.stem for path in ACCOUNTS_CONFIG_DIR.glob("*.env"))


def account_exists(account: str) -> bool:
    return account_env_file(account).exists()


def _load_account_env(account: str | None) -> dict[str, str]:
    if not account:
        return {}
    path = account_env_file(account)
    if path.exists():
        return {str(k): str(v) for k, v in dotenv_values(path).items() if v is not None}
    return {}


def _env(name: str, default: str = "") -> str:
    if ACTIVE_ACCOUNT and name in _ACCOUNT_ENV:
        return _ACCOUNT_ENV[name]
    return os.environ.get(name, default)


def _account_data_dir(account: str | None) -> Path:
    if not account:
        return DATA_DIR
    return ACCOUNTS_DATA_DIR / account


def resolve_config_path(env_name: str, default: Path, *, base_dir: Path | None = None) -> Path:
    """Resolve an optional path from .env, relative to the data dir if needed."""
    raw = _env(env_name).strip()
    if not raw:
        return default
    path = Path(raw).expanduser()
    if not path.is_absolute():
        path = (base_dir or ACCOUNT_DATA_DIR) / path
    return path


def activate_account(account: str | None) -> None:
    """Activate an account profile for this process."""
    global ACTIVE_ACCOUNT, ACCOUNT_DATA_DIR, SESSION_DIR, CACHE_DIR, TOKENS_FILE, _ACCOUNT_ENV
    global MS_EMAIL, MS_PASSWORD, LOCAL_TIMEZONE, OUTLOOK_TIMEZONE
    global SIGNATURE_NEW_FILE, SIGNATURE_REPLY_FILE
    global GATEWAY_PID_FILE, GATEWAY_LOG_FILE, GATEWAY_CHAT_ID_FILE, GATEWAY_STATE_FILE
    global GATEWAY_WORKSPACE_DIR, GATEWAY_CHAT_ID, GATEWAY_TRIGGER, GATEWAY_POLL_INTERVAL

    if account and not _valid_account_name(account):
        raise ValueError(
            "Account names may contain letters, numbers, dots, underscores, and dashes."
        )

    ACTIVE_ACCOUNT = account
    ACCOUNT_DATA_DIR = _account_data_dir(account)
    SESSION_DIR = ACCOUNT_DATA_DIR / "session_state"
    CACHE_DIR = ACCOUNT_DATA_DIR / "cache"
    TOKENS_FILE = SESSION_DIR / "tokens.json"

    _ACCOUNT_ENV = _load_account_env(account)

    MS_EMAIL = _env("MS_EMAIL", "")
    MS_PASSWORD = _env("MS_PASSWORD", "")
    LOCAL_TIMEZONE = _env("LOCAL_TIMEZONE", "Europe/London")
    OUTLOOK_TIMEZONE = _env("OUTLOOK_TIMEZONE", "GMT Standard Time")

    SIGNATURE_NEW_FILE = resolve_config_path(
        "SIGNATURE_NEW_FILE",
        ACCOUNT_DATA_DIR / "signature-new.html",
        base_dir=ACCOUNT_DATA_DIR,
    )
    SIGNATURE_REPLY_FILE = resolve_config_path(
        "SIGNATURE_REPLY_FILE",
        ACCOUNT_DATA_DIR / "signature-reply.html",
        base_dir=ACCOUNT_DATA_DIR,
    )

    GATEWAY_PID_FILE = SESSION_DIR / "gateway.pid"
    GATEWAY_LOG_FILE = SESSION_DIR / "gateway.log"
    GATEWAY_CHAT_ID_FILE = SESSION_DIR / "gateway_chat_id.txt"
    GATEWAY_STATE_FILE = SESSION_DIR / "gateway_state.json"
    GATEWAY_WORKSPACE_DIR = ACCOUNT_DATA_DIR / "gateway_workspaces"
    GATEWAY_CHAT_ID = _env("GATEWAY_CHAT_ID", "")
    GATEWAY_TRIGGER = _env("GATEWAY_TRIGGER", "@Marlow")
    GATEWAY_POLL_INTERVAL = int(_env("GATEWAY_POLL_INTERVAL", "30"))


def set_active_account(account: str) -> None:
    if not account_exists(account):
        raise FileNotFoundError(f"No account profile named '{account}'.")
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    ACTIVE_ACCOUNT_FILE.write_text(account + "\n")
    ACTIVE_ACCOUNT_FILE.chmod(0o600)
    activate_account(account)


def clear_active_account() -> None:
    ACTIVE_ACCOUNT_FILE.unlink(missing_ok=True)
    activate_account(None)


def ensure_dirs() -> None:
    CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    ACCOUNTS_CONFIG_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    if ACTIVE_ACCOUNT:
        ACCOUNTS_DATA_DIR.mkdir(parents=True, exist_ok=True)
        ACCOUNT_DATA_DIR.mkdir(parents=True, exist_ok=True)
    SESSION_DIR.mkdir(parents=True, exist_ok=True)
    CACHE_DIR.mkdir(parents=True, exist_ok=True)


OUTLOOK_BASE_URL: str = "https://outlook.office.com/api/v2.0"
GRAPH_BASE_URL: str = "https://graph.microsoft.com/v1.0"
OUTLOOK_TOKEN_DOMAIN: str = "outlook.office.com"
GRAPH_TOKEN_DOMAIN: str = "graph.microsoft.com"
EMAIL_FONT_FAMILY: str = (
    "Aptos, Aptos_EmbeddedFont, Aptos_MSFontService, Calibri, Helvetica, sans-serif"
)
EMAIL_FONT_SIZE: str = "12pt"

# Token refresh: warn this many seconds before expiry
TOKEN_WARN_BEFORE_EXPIRY: int = 300  # 5 minutes

activate_account(_read_active_account())
