from __future__ import annotations

import subprocess
import sys


def run_help(*args: str) -> str:
    result = subprocess.run(
        [sys.executable, "-m", "outlook_draft", *args, "--help"],
        check=True,
        text=True,
        capture_output=True,
    )
    return result.stdout


def test_json_before_domain_is_accepted() -> None:
    output = run_help("--json", "mail", "search", "test")

    assert "--json" in output
    assert "--table" in output


def test_json_between_domain_and_command_is_accepted() -> None:
    output = run_help("mail", "--json", "search", "test")

    assert "--json" in output
    assert "--table" in output


def test_json_after_leaf_command_is_accepted() -> None:
    output = run_help("mail", "search", "test", "--json")

    assert "--json" in output
    assert "--table" in output


def test_calendar_table_alias_is_accepted() -> None:
    output = run_help("cal", "agenda", "--table")

    assert "--json" in output
    assert "--table" in output


def test_no_spinner_global_flag_is_accepted() -> None:
    output = run_help("--no-spinner", "teams", "search", "ross")

    assert "--scan" in output
