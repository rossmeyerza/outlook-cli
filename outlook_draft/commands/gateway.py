"""Teams-to-pi gateway daemon.

Polls a designated Teams chat for messages containing the configured trigger
(default: @Marlow), passes them to pi, and posts the response back to Teams.

Commands:
  outlook-cli gateway start [--chat-id ID] [--poll SECONDS]
  outlook-cli gateway stop
  outlook-cli gateway status
"""

from __future__ import annotations

import argparse
import html
import os
import re
import shutil
import signal
import subprocess
import sys
import time
from datetime import datetime, timezone

from rich.console import Console

from .. import config
from ..errors import OutlookAPIError
from ..outlook_client import OutlookClient
from ..token_manager import TokenManager


def _graph_client() -> OutlookClient:
    tm = TokenManager(
        token_domain=config.GRAPH_TOKEN_DOMAIN,
        token_label="Graph",
    )
    return OutlookClient(tm, base_url=config.GRAPH_BASE_URL)


def _log(message: str) -> None:
    config.SESSION_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with config.GATEWAY_LOG_FILE.open("a") as f:
        f.write(f"[{ts}] {message}\n")


def _strip_html(text: str) -> str:
    text = re.sub(r"<[^>]+>", "", text)
    return html.unescape(text).strip()


def _resolve_chat_id(args: argparse.Namespace, console: Console) -> str:
    chat_id = getattr(args, "chat_id", None) or config.GATEWAY_CHAT_ID
    if chat_id:
        return chat_id
    if config.GATEWAY_CHAT_ID_FILE.exists():
        stored = config.GATEWAY_CHAT_ID_FILE.read_text().strip()
        if stored:
            return stored

    console.print("[yellow]No gateway chat configured. Fetching your Teams chats...[/]")
    client = _graph_client()
    try:
        chats = client.list_teams_chats(top=20)
        chats.sort(key=lambda c: c.get("lastUpdatedDateTime") or "", reverse=True)
    finally:
        client.close()

    if not chats:
        console.print("[red]No Teams chats found.[/]")
        sys.exit(1)

    console.print("\nPick the chat to monitor for trigger messages:\n")
    for i, chat in enumerate(chats, 1):
        topic = chat.get("topic") or ""
        chat_type = chat.get("chatType", "")
        updated = (chat.get("lastUpdatedDateTime") or "")[:10]
        console.print(f"  [bold]{i:2}[/]  {updated}  [{chat_type}]  {topic or '(no topic)'}")

    console.print()
    raw = input("Enter number: ").strip()
    try:
        idx = int(raw) - 1
        chosen = chats[idx]
    except (ValueError, IndexError):
        console.print("[red]Invalid selection.[/]")
        sys.exit(1)

    selected_id: str = chosen["id"]
    config.GATEWAY_CHAT_ID_FILE.write_text(selected_id)
    console.print(f"[green]Chat saved:[/] {selected_id}")
    return selected_id


def _read_pid() -> int | None:
    if not config.GATEWAY_PID_FILE.exists():
        return None
    try:
        return int(config.GATEWAY_PID_FILE.read_text().strip())
    except (ValueError, OSError):
        return None


def _process_is_running(pid: int) -> bool:
    try:
        os.kill(pid, 0)
        return True
    except (ProcessLookupError, PermissionError):
        return False


def _call_pi(prompt: str) -> str:
    pi_bin = shutil.which("pi")
    if not pi_bin:
        return "[Error: pi not found in PATH]"
    try:
        result = subprocess.run(
            [pi_bin, "--print", prompt],
            capture_output=True,
            text=True,
            timeout=180,
        )
        response = result.stdout.strip()
        if not response and result.stderr.strip():
            response = result.stderr.strip()
        return response or "[No response from pi]"
    except subprocess.TimeoutExpired:
        return "[pi timed out after 3 minutes]"
    except Exception as exc:
        return f"[pi error: {exc}]"


def _poll_loop(chat_id: str, trigger: str, poll_interval: int) -> None:
    trigger_re = re.compile(re.escape(trigger), re.IGNORECASE)
    last_seen_time = datetime.now(timezone.utc).isoformat()
    seen_ids: set[str] = set()

    _log(f"Gateway started. chat={chat_id} trigger={trigger} poll={poll_interval}s")

    while True:
        try:
            client = _graph_client()
            try:
                messages = client.list_teams_messages(chat_id, top=10)
            finally:
                client.close()

            messages.sort(key=lambda m: m.get("createdDateTime") or "")
            for msg in messages:
                msg_id = msg.get("id", "")
                msg_time = msg.get("createdDateTime") or ""
                if msg_id in seen_ids:
                    continue
                seen_ids.add(msg_id)
                if msg_time <= last_seen_time:
                    continue
                if msg.get("deletedDateTime"):
                    continue
                if msg.get("messageType") == "systemEventMessage":
                    continue

                body = msg.get("body", {})
                content = body.get("content", "")
                if body.get("contentType") == "html":
                    content = _strip_html(content)
                content = content.strip()

                if not content or not trigger_re.search(content):
                    continue

                prompt = trigger_re.sub("", content).strip().lstrip(":,. ")
                if not prompt:
                    prompt = "(no message)"

                sender = (
                    ((msg.get("from") or {}).get("user") or {}).get("displayName")
                    or "unknown"
                )
                _log(f"Triggered by {sender}: {prompt[:120]}")

                response = _call_pi(prompt)
                _log(f"Response ({len(response)} chars): {response[:80]}")

                try:
                    post_client = _graph_client()
                    try:
                        post_client._send_teams_message_internal(
                            chat_id,
                            f"[Marlow] {response}",
                        )
                    finally:
                        post_client.close()
                    _log("Response posted to Teams.")
                except Exception as exc:
                    _log(f"Failed to post response: {exc}")

        except OutlookAPIError as exc:
            _log(f"Graph API error: {exc}")
        except Exception as exc:
            _log(f"Unexpected poll error: {exc}")

        time.sleep(poll_interval)


def cmd_gateway_start(args: argparse.Namespace) -> None:
    console = Console()

    existing_pid = _read_pid()
    if existing_pid and _process_is_running(existing_pid):
        console.print(f"[yellow]Gateway already running (PID {existing_pid}).[/]")
        return

    chat_id = _resolve_chat_id(args, console)
    trigger = getattr(args, "trigger", None) or config.GATEWAY_TRIGGER
    poll_interval = getattr(args, "poll", None) or config.GATEWAY_POLL_INTERVAL

    console.print("[green]Starting gateway...[/]")
    console.print(f"  Chat:    [bold]{chat_id}[/]")
    console.print(f"  Trigger: [bold]{trigger}[/]")
    console.print(f"  Poll:    every [bold]{poll_interval}s[/]")
    console.print(f"  Logs:    {config.GATEWAY_LOG_FILE}")

    pid = os.fork()
    if pid > 0:
        config.SESSION_DIR.mkdir(parents=True, exist_ok=True)
        config.GATEWAY_PID_FILE.write_text(str(pid))
        console.print(f"[green]Gateway running (PID {pid}).[/]")
        return

    os.setsid()
    devnull = os.open(os.devnull, os.O_RDWR)
    for fd in (0, 1, 2):
        os.dup2(devnull, fd)
    os.close(devnull)

    try:
        _poll_loop(chat_id, trigger, poll_interval)
    except Exception as exc:
        _log(f"Gateway crashed: {exc}")
    finally:
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)
    os._exit(0)


def cmd_gateway_stop(args: argparse.Namespace) -> None:
    console = Console()
    pid = _read_pid()
    if not pid:
        console.print("[yellow]Gateway is not running (no PID file).[/]")
        return
    if not _process_is_running(pid):
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)
        console.print("[yellow]Gateway process not found. PID file cleaned up.[/]")
        return
    try:
        os.kill(pid, signal.SIGTERM)
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)
        console.print(f"[green]Gateway stopped (PID {pid}).[/]")
    except ProcessLookupError:
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)
        console.print("[yellow]Process already gone.[/]")


def cmd_gateway_status(args: argparse.Namespace) -> None:
    console = Console()
    pid = _read_pid()
    running = bool(pid and _process_is_running(pid))

    chat_id = config.GATEWAY_CHAT_ID
    if not chat_id and config.GATEWAY_CHAT_ID_FILE.exists():
        chat_id = config.GATEWAY_CHAT_ID_FILE.read_text().strip()

    if running:
        console.print(f"[green]Gateway is running[/] (PID {pid})")
    else:
        console.print("[dim]Gateway is not running.[/]")

    if chat_id:
        console.print(f"  Chat:    {chat_id}")
    console.print(f"  Trigger: {config.GATEWAY_TRIGGER}")
    console.print(f"  Poll:    every {config.GATEWAY_POLL_INTERVAL}s")
    console.print(f"  Log:     {config.GATEWAY_LOG_FILE}")

    if config.GATEWAY_LOG_FILE.exists():
        lines = config.GATEWAY_LOG_FILE.read_text().splitlines()
        recent = lines[-10:]
        if recent:
            console.print("\n[bold]Recent log:[/]")
            for line in recent:
                console.print(f"  [dim]{line}[/]")
    else:
        console.print("\n[dim]No log yet.[/]")
