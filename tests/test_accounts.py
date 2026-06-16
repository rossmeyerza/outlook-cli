from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path


def run_cli(
    config_home: Path,
    data_home: Path,
    *args: str,
    check: bool = True,
) -> subprocess.CompletedProcess[str]:
    env = os.environ.copy()
    env["XDG_CONFIG_HOME"] = str(config_home)
    env["XDG_DATA_HOME"] = str(data_home)
    return subprocess.run(
        [sys.executable, "-m", "outlook_draft", *args],
        text=True,
        capture_output=True,
        env=env,
        check=check,
    )


def test_account_add_switch_and_current_json(tmp_path: Path) -> None:
    config_home = tmp_path / "config"
    data_home = tmp_path / "data"

    run_cli(
        config_home,
        data_home,
        "account",
        "add",
        "ikea",
        "--email",
        "ikea@example.com",
        "--password",
        "secret",
        "--switch",
    )

    current = run_cli(config_home, data_home, "account", "current", "--json")
    payload = json.loads(current.stdout)

    assert payload["activeAccount"] == "ikea"
    assert payload["email"] == "ikea@example.com"
    assert payload["sessionDir"].endswith("/outlook-cli/accounts/ikea/session_state")


def test_account_env_does_not_inherit_global_password(tmp_path: Path) -> None:
    config_home = tmp_path / "config"
    data_home = tmp_path / "data"
    config_dir = config_home / "outlook-cli"
    config_dir.mkdir(parents=True)
    (config_dir / ".env").write_text(
        "MS_EMAIL=global@example.com\nMS_PASSWORD=global-secret\n",
        encoding="utf-8",
    )

    run_cli(
        config_home,
        data_home,
        "account",
        "add",
        "ikea",
        "--email",
        "ikea@example.com",
        "--switch",
    )

    check = run_cli(config_home, data_home, "config", "check", "--json", check=False)
    items = {item["name"]: item for item in json.loads(check.stdout)}

    assert items["MS_EMAIL"]["detail"] == "ikea@example.com"
    assert items["MS_PASSWORD"]["ok"] is False


def test_account_override_can_appear_after_domain(tmp_path: Path) -> None:
    config_home = tmp_path / "config"
    data_home = tmp_path / "data"
    run_cli(
        config_home,
        data_home,
        "account",
        "add",
        "ikea",
        "--email",
        "ikea@example.com",
    )

    current = run_cli(
        config_home,
        data_home,
        "account",
        "--account",
        "ikea",
        "current",
        "--json",
    )
    payload = json.loads(current.stdout)

    assert payload["activeAccount"] == "ikea"
