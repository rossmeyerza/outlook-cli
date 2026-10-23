"""Native host entry point. Never write diagnostic output to stdout."""
from __future__ import annotations

import logging
import sys

from .. import config
from ..errors import CopilotError
from .native_host import run, secure_directory


def main() -> int:
    try:
        log_dir = config.DATA_DIR / "session_state"
        secure_directory(log_dir)
        log_path = log_dir / "copilot-bridge.log"
        log_path.touch(mode=0o600, exist_ok=True)
        log_path.chmod(0o600)
        logging.basicConfig(filename=log_path, level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
        run()
    except (CopilotError, OSError) as exc:
        logging.error("Native host stopped: %s", exc)
        print(f"Copilot native host: {exc}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
