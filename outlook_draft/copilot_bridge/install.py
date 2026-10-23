"""Install the Chrome native messaging manifest and launcher, not the extension."""
from __future__ import annotations

import json
import os
import shlex
import shutil
import sys
from pathlib import Path

from .. import config
from ..errors import CopilotError
from .protocol import EXTENSION_ID, HOST_NAME


def extension_source() -> Path:
    source = Path(__file__).resolve().parents[2] / "extension"
    if not (source / "manifest.json").is_file():
        # Wheel installs carry a packaged copy of the extension.
        source = Path(sys.prefix) / "share" / "outlook-cli" / "extension"
    if not (source / "manifest.json").is_file():
        raise CopilotError("The Chrome extension files are missing from this install.")
    return source


def install(*, config_home: Path | None = None, data_dir: Path | None = None) -> dict:
    config_home = config_home or Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    data_dir = data_dir or config.DATA_DIR
    host_dir = config_home / "google-chrome" / "NativeMessagingHosts"
    host_dir.mkdir(parents=True, exist_ok=True)
    data_dir.mkdir(parents=True, exist_ok=True)
    target = data_dir / "copilot-extension"
    source = extension_source()
    if target.is_symlink():
        raise CopilotError("Refusing to install into a symlinked extension folder.")
    shutil.copytree(source, target, dirs_exist_ok=True, ignore=shutil.ignore_patterns("tests", "*.test.js"))
    launcher = data_dir / "copilot-native-host"
    python = shlex.quote(sys.executable)
    project = shlex.quote(str(Path(__file__).resolve().parents[2]))
    # Use this install's interpreter. Run as a module so no CLI output reaches stdout.
    script = f'#!/bin/sh\ncd {project} || exit 1\nexec {python} -m outlook_draft.copilot_bridge "$@"\n'
    launcher.write_text(script)
    launcher.chmod(0o700)
    manifest = host_dir / f"{HOST_NAME}.json"
    manifest.write_text(json.dumps({
        "name": HOST_NAME, "description": "outlook-cli Copilot bridge",
        "path": str(launcher), "type": "stdio",
        "allowed_origins": [f"chrome-extension://{EXTENSION_ID}/"],
    }, indent=2) + "\n")
    manifest.chmod(0o600)
    return {"manifest": str(manifest), "launcher": str(launcher), "extension": str(target), "extensionId": EXTENSION_ID}
