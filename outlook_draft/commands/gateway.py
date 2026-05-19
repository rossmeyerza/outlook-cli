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
from pathlib import Path

from rich.console import Console

from .. import config
from ..errors import OutlookAPIError
from ..outlook_client import OutlookClient
from ..token_manager import TokenManager
from .pi_session import PiSession, PiSessionError


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


def _chat_label(chat: dict, members: list[dict] | None) -> str:
    """Build a human-readable label: topic if present, else member names."""
    topic = chat.get("topic") or ""
    if topic:
        return topic
    if not members:
        return "(no topic)"
    names = [
        (m.get("displayName") or m.get("email") or "").strip()
        for m in members
    ]
    names = [n for n in names if n]
    if not names:
        return "(no topic)"
    if len(names) <= 3:
        return ", ".join(names)
    return f"{', '.join(names[:3])} +{len(names) - 3}"


def _resolve_chat_id(args: argparse.Namespace, console: Console) -> str:
    chat_id = getattr(args, "chat_id", None) or config.GATEWAY_CHAT_ID
    if chat_id:
        return chat_id
    if config.GATEWAY_CHAT_ID_FILE.exists():
        stored = config.GATEWAY_CHAT_ID_FILE.read_text().strip()
        if stored:
            return stored

    console.print("[yellow]No gateway chat configured. Fetching ALL your Teams chats (this can take a moment)...[/]")
    client = _graph_client()
    try:
        chats = client.list_teams_chats(top=500)
        chats.sort(key=lambda c: c.get("lastUpdatedDateTime") or "", reverse=True)
        # Resolve members for chats without a topic, in parallel-ish (sequential but quick per call)
        console.print(f"[dim]Resolving members for {len(chats)} chats...[/]")
        labels: list[tuple[str, dict]] = []
        for chat in chats:
            members = None
            if not chat.get("topic"):
                try:
                    members = client.list_teams_chat_members(chat["id"], top=20)
                except OutlookAPIError:
                    members = None
            label = _chat_label(chat, members)
            labels.append((label, chat))
    finally:
        client.close()

    if not labels:
        console.print("[red]No Teams chats found.[/]")
        sys.exit(1)

    chosen = _pick_chat_with_fzf(labels, console)
    if not chosen:
        console.print("[red]No selection made.[/]")
        sys.exit(1)

    selected_id: str = chosen["id"]
    config.GATEWAY_CHAT_ID_FILE.write_text(selected_id)
    console.print(f"[green]Chat saved:[/] {selected_id}")
    return selected_id


def _pick_chat_with_fzf(labels: list[tuple[str, dict]], console: Console) -> dict | None:
    """Use fzf if available for interactive type-to-filter, else fall back to numbered list."""
    fzf_bin = shutil.which("fzf")
    # Build display lines: index<TAB>updated<TAB>type<TAB>label
    lines = []
    for i, (label, chat) in enumerate(labels):
        updated = (chat.get("lastUpdatedDateTime") or "")[:10]
        chat_type = chat.get("chatType", "")
        lines.append(f"{i}\t{updated}\t{chat_type:<10}\t{label}")

    if fzf_bin:
        try:
            result = subprocess.run(
                [fzf_bin, "--prompt=Chat: ", "--with-nth=2..", "--delimiter=\t",
                 "--height=70%", "--reverse", "--no-mouse"],
                input="\n".join(lines),
                capture_output=True,
                text=True,
            )
            if result.returncode == 0 and result.stdout.strip():
                idx_str = result.stdout.split("\t", 1)[0]
                idx = int(idx_str)
                return labels[idx][1]
            return None
        except Exception as exc:
            console.print(f"[dim]fzf failed ({exc}), falling back to text picker.[/]")

    # Fallback: show first 50 numbered + accept search query
    console.print("\n[bold]First 50 chats (most recent):[/]")
    for i, (label, chat) in enumerate(labels[:50], 1):
        updated = (chat.get("lastUpdatedDateTime") or "")[:10]
        chat_type = chat.get("chatType", "")
        console.print(f"  [bold]{i:3}[/]  {updated}  [{chat_type}]  {label}")
    console.print(f"\n[dim]Total: {len(labels)} chats. Type a number, or type a search term to filter.[/]")
    raw = input("> ").strip()
    if not raw:
        return None
    if raw.isdigit():
        idx = int(raw) - 1
        if 0 <= idx < len(labels):
            return labels[idx][1]
        return None
    # Search
    matches = [(i, lbl, ch) for i, (lbl, ch) in enumerate(labels) if raw.lower() in lbl.lower()]
    if not matches:
        console.print("[red]No chats matched.[/]")
        return None
    if len(matches) == 1:
        return matches[0][2]
    console.print(f"\n[bold]{len(matches)} matches:[/]")
    for n, (i, lbl, ch) in enumerate(matches[:50], 1):
        updated = (ch.get("lastUpdatedDateTime") or "")[:10]
        console.print(f"  [bold]{n:3}[/]  {updated}  {lbl}")
    raw2 = input("Pick number: ").strip()
    if raw2.isdigit():
        idx = int(raw2) - 1
        if 0 <= idx < len(matches):
            return matches[idx][2]
    return None


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


# Persistent pi RPC sessions, keyed by chat_id
_pi_sessions: dict[str, PiSession] = {}

# Tracks the timestamp of the last @Marlow turn we processed per chat. Used to
# build a "catch-up" of any chat that happened between Marlow turns so Marlow
# always has the full context of what you've been discussing.
_last_marlow_turn_time: dict[str, str] = {}


def _get_pi_session(chat_id: str) -> PiSession:
    """Get or create the persistent pi session for this chat."""
    sess = _pi_sessions.get(chat_id)
    if sess and sess.is_alive():
        return sess
    if sess:
        try:
            sess.close()
        except Exception:
            pass
    session_dir = config.SESSION_DIR / "gateway_sessions" / f"{abs(hash(chat_id))}"
    sess = PiSession(session_dir, log_fn=_log)
    sess.start()
    _pi_sessions[chat_id] = sess
    return sess


def _build_preamble(
    chat_id: str,
    chat_topic: str,
    chat_type: str,
    members: list[dict],
    sender: str,
    sender_time: str,
    recent_messages: list[dict],
) -> str:
    """Build the context preamble injected as the first message of a session."""
    member_names = [
        (m.get("displayName") or m.get("email") or "").strip()
        for m in members
    ]
    member_names = [n for n in member_names if n]
    members_str = ", ".join(member_names) if member_names else "(unknown)"

    history_lines = []
    for m in recent_messages[-15:]:
        if m.get("messageType") == "systemEventMessage":
            continue
        if m.get("deletedDateTime"):
            continue
        body = m.get("body", {})
        content = body.get("content", "")
        if body.get("contentType") == "html":
            content = _strip_html(content)
        content = content.strip().replace("\n", " ")
        if not content:
            continue
        ts = (m.get("createdDateTime") or "")[11:16]
        sender_name = ((m.get("from") or {}).get("user") or {}).get("displayName") or "?"
        history_lines.append(f"[{ts}] {sender_name}: {content[:300]}")
    history_block = "\n".join(history_lines) if history_lines else "(no recent messages)"

    return f"""You are Marlow, responding inside a Microsoft Teams chat as Ross Meyer.
Ross has set up a gateway that pipes Teams messages to you when they contain @Marlow.

CHAT CONTEXT
- Type: {chat_type}
- Topic: {chat_topic or '(none)'}
- Members: {members_str}

RECENT MESSAGES (oldest first, just for context)
{history_block}

TOOLS YOU HAVE
You have full bash access. The `outlook-cli` command is installed. Useful subcommands:
  outlook-cli mail unread / mail search <q> / mail read <n>
  outlook-cli cal agenda / cal show <n>
  outlook-cli teams list / teams messages <chat>
  outlook-cli files list [path] [--site NAME]
  outlook-cli task list / contact search <q>
Use them only if you need information that isn't already in the chat above.

RESPONSE GUIDELINES
- This is a Teams chat, not a code review. Keep replies short and conversational.
- Plain text or simple HTML only. Teams strips most markdown.
- No code fences unless the user is explicitly asking about code.
- Don't sign off, don't introduce yourself again, just answer.
- Don't use em-dashes (—). Use commas, full stops, or restructure.

The rest of this conversation will be incoming Teams messages. Each one will name
the sender. Respond to the latest one.
"""


def _format_catchup(
    messages: list[dict], since_time: str, trigger_msg_id: str
) -> str:
    """Format messages between since_time and now (excluding the trigger) as catch-up.

    Skips system events, deleted messages, and the trigger message itself.
    Returns empty string if nothing to catch up on.
    """
    lines = []
    for m in messages:
        if m.get("id") == trigger_msg_id:
            continue
        if m.get("messageType") == "systemEventMessage":
            continue
        if m.get("deletedDateTime"):
            continue
        m_time = m.get("createdDateTime") or ""
        if m_time <= since_time:
            continue
        body = m.get("body", {})
        content = body.get("content", "")
        if body.get("contentType") == "html":
            content = _strip_html(content)
        content = content.strip().replace("\n", " ")
        if not content:
            continue
        ts = m_time[11:16]
        sender_name = ((m.get("from") or {}).get("user") or {}).get("displayName") or "?"
        lines.append(f"[{ts}] {sender_name}: {content[:300]}")
    return "\n".join(lines)


def _call_pi(
    chat_id: str,
    prompt: str,
    *,
    chat_topic: str = "",
    chat_type: str = "",
    members: list[dict] | None = None,
    sender: str = "",
    sender_time: str = "",
    recent_messages: list[dict] | None = None,
    trigger_msg_id: str = "",
) -> str:
    """Send a prompt to the per-chat persistent pi session."""
    try:
        sess = _get_pi_session(chat_id)
        is_first = (
            not sess.has_existing_session()
            and not getattr(sess, "_preamble_sent", False)
        )

        if is_first:
            preamble = _build_preamble(
                chat_id=chat_id,
                chat_topic=chat_topic,
                chat_type=chat_type,
                members=members or [],
                sender=sender,
                sender_time=sender_time,
                recent_messages=recent_messages or [],
            )
            full_prompt = (
                preamble
                + f"\n\nLATEST MESSAGE\n[{sender_time}] {sender}: {prompt}"
            )
            sess._preamble_sent = True  # type: ignore[attr-defined]
        else:
            # Build catch-up block of any chat that happened since our last turn
            since = _last_marlow_turn_time.get(chat_id, "")
            catchup = ""
            if since and recent_messages:
                catchup = _format_catchup(recent_messages, since, trigger_msg_id)

            if catchup:
                full_prompt = (
                    "CATCH-UP (chat that happened while you were away)\n"
                    f"{catchup}\n\n"
                    f"LATEST MESSAGE\n[{sender_time}] {sender}: {prompt}"
                )
            else:
                full_prompt = f"[{sender_time}] {sender}: {prompt}"

        response = sess.prompt(full_prompt)
        # Mark this trigger time as the last Marlow turn for catch-up tracking
        if sender_time:
            _last_marlow_turn_time[chat_id] = sender_time
        return response
    except PiSessionError as exc:
        _log(f"PiSession error: {exc}")
        # Reset the session so the next message gets a clean process
        try:
            old = _pi_sessions.pop(chat_id, None)
            if old:
                old.close()
        except Exception:
            pass
        return f"[pi error: {exc}]"
    except Exception as exc:
        _log(f"Unexpected pi error: {exc}")
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
                # Fetch a wider window so we can build the catch-up between Marlow turns
                messages = client.list_teams_messages(chat_id, top=50)
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

                # Fetch chat metadata + members for context (cheap, cached effectively
                # because pi keeps the session alive across messages — the preamble is
                # only built on the first message of a fresh session).
                chat_topic = ""
                chat_type = ""
                members: list[dict] = []
                try:
                    meta_client = _graph_client()
                    try:
                        meta = meta_client.get_teams_chat(chat_id)
                        chat_topic = meta.get("topic") or ""
                        chat_type = meta.get("chatType") or ""
                        members = meta_client.list_teams_chat_members(chat_id, top=20)
                    finally:
                        meta_client.close()
                except Exception:
                    pass

                response = _call_pi(
                    chat_id,
                    prompt,
                    chat_topic=chat_topic,
                    chat_type=chat_type,
                    members=members,
                    sender=sender,
                    sender_time=msg_time[:16],
                    recent_messages=messages,
                    trigger_msg_id=msg_id,
                )
                _log(f"Response ({len(response)} chars): {response[:80]}")

                # Format as HTML with robot emoji prefix and a styled block
                html_body = (
                    "<div style='border-left:3px solid #6264a7;"
                    "padding:6px 12px;background:#f3f2f1;'>"
                    "<strong>\U0001f916 Marlow</strong><br>"
                    f"{html.escape(response).replace(chr(10), '<br>')}"
                    "</div>"
                )

                try:
                    post_client = _graph_client()
                    try:
                        post_client._send_teams_message_internal(
                            chat_id,
                            html_body,
                            content_type="html",
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
