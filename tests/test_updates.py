import json
import sys
from unittest.mock import Mock

import httpx
import pytest

from outlook_draft import cli, updates


@pytest.fixture
def check(monkeypatch, tmp_path):
    monkeypatch.delenv("OUTLOOK_CLI_NO_UPDATE_CHECK")
    monkeypatch.setenv("XDG_CACHE_HOME", str(tmp_path))
    monkeypatch.setattr(updates, "installed_version", lambda: "0.9.0")
    clock = Mock(return_value=100_000)
    monkeypatch.setattr(updates, "time", Mock(time=clock))
    request = httpx.Request("GET", updates.LATEST_URL)
    get = Mock(return_value=httpx.Response(200, request=request, json={
        "draft": False, "prerelease": False, "tag_name": "v0.10.0",
    }))
    monkeypatch.setattr(updates.httpx, "get", get)
    return get, clock, tmp_path / "outlook-cli" / "update-check.json"


@pytest.mark.parametrize("command", [["mail", "unread", "--json"], ["mcp"]])
def test_notice_stays_off_stdout(check, command, monkeypatch, capsys):
    monkeypatch.setattr(sys, "argv", ["outlook-cli", *command])
    monkeypatch.setattr(cli.account, "refresh_account_paths", lambda: None)
    handler = (cli.mail, "cmd_unread") if command[0] == "mail" else (cli.mcp, "cmd_mcp")
    monkeypatch.setattr(*handler, lambda _args: print('{"ok": true}'))
    cli.main()
    output = capsys.readouterr()
    assert json.loads(output.out) == {"ok": True}
    assert "Update available (0.9.0 to v0.10.0)" in output.err


def test_success_cache_and_version_comparison(check, monkeypatch, capsys):
    get, clock, _ = check
    updates.notify_if_update_available()
    assert "Update available" in capsys.readouterr().err
    get.assert_called_once_with(updates.LATEST_URL, headers={"Accept": "application/vnd.github+json"}, timeout=1.0)
    for current in ("0.10.0", "0.11.0"):
        monkeypatch.setattr(updates, "installed_version", lambda current=current: current)
        updates.notify_if_update_available()
        assert not capsys.readouterr().err
    assert get.call_count == 1
    clock.return_value += updates.SUCCESS_TTL
    updates.notify_if_update_available()
    assert get.call_count == 2


@pytest.mark.parametrize("failure", ["offline", 404, 429, "invalid", "prerelease", "draft", "invalid-tag"])
def test_failed_check_is_quiet_and_retries_after_backoff(check, failure, capsys):
    get, clock, _ = check
    request = httpx.Request("GET", updates.LATEST_URL)
    if failure == "offline":
        get.side_effect = httpx.ConnectTimeout("offline")
    elif isinstance(failure, int):
        get.return_value = httpx.Response(failure, request=request)
    elif failure == "invalid":
        get.return_value = httpx.Response(200, request=request, content=b"not JSON")
    else:
        get.return_value = httpx.Response(200, request=request, json={
            "draft": failure == "draft",
            "prerelease": failure == "prerelease",
            "tag_name": "not-a-version" if failure == "invalid-tag" else "v9.0.0",
        })
    updates.notify_if_update_available()
    updates.notify_if_update_available()
    assert get.call_count == 1
    clock.return_value += updates.FAILURE_TTL
    updates.notify_if_update_available()
    assert get.call_count == 2
    assert not capsys.readouterr().err


def test_corrupt_and_unwritable_cache(check, capsys):
    get, _, path = check
    path.parent.mkdir()
    path.mkdir()
    updates.notify_if_update_available()
    get.assert_called_once()
    assert "Update available" in capsys.readouterr().err
    assert list(path.parent.iterdir()) == [path]


def test_opt_out_and_dry_run_do_not_check(check, monkeypatch):
    get, _, path = check
    monkeypatch.setenv("OUTLOOK_CLI_NO_UPDATE_CHECK", "1")
    updates.notify_if_update_available()
    monkeypatch.delenv("OUTLOOK_CLI_NO_UPDATE_CHECK")
    monkeypatch.setattr(sys, "argv", ["outlook-cli", "mail", "unread", "--dry-run"])
    cli.main()
    get.assert_not_called()
    assert not path.exists()
