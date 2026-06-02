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
import hashlib
import html
import json
import os
import re
import shlex
import shutil
import signal
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import quote

from rich.console import Console

from .. import config
from ..errors import OutlookAPIError, TokenExpiredError, TokenNotFoundError
from ..outlook_client import OutlookClient
from ..token_manager import TokenManager
from .pi_session import PiSession, PiSessionError
from .teams import SELF_CHAT_ID, _chat_title, find_self_chat


EXPORT_MANIFEST = ".marlow-export.json"
EXPORT_MAX_BYTES = 25 * 1024 * 1024
EXPORT_REMOTE_ROOT = "Outlook CLI/Gateway"
EXPORT_ALLOWED_SUFFIXES = {
    ".csv",
    ".docx",
    ".html",
    ".jpeg",
    ".jpg",
    ".json",
    ".md",
    ".pdf",
    ".png",
    ".pptx",
    ".txt",
    ".xlsx",
}


def _graph_client() -> OutlookClient:
    tm = TokenManager(
        token_domain=config.GRAPH_TOKEN_DOMAIN,
        token_label="Graph",
    )
    if tm.is_expired:
        _log("Graph token missing or expired; attempting headless re-authentication")
        if not tm.run_reauth(headless=True):
            raise TokenExpiredError("Graph token has expired and re-authentication failed")
        _log("Graph re-authentication succeeded")
    return OutlookClient(tm, base_url=config.GRAPH_BASE_URL)


def _log(message: str) -> None:
    config.SESSION_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with config.GATEWAY_LOG_FILE.open("a") as f:
        f.write(f"[{ts}] {message}\n")


def _read_gateway_state() -> dict:
    if not config.GATEWAY_STATE_FILE.exists():
        return {}
    try:
        return json.loads(config.GATEWAY_STATE_FILE.read_text())
    except (OSError, json.JSONDecodeError):
        return {}


def _write_gateway_state(**updates: object) -> None:
    config.SESSION_DIR.mkdir(parents=True, exist_ok=True)
    state = _read_gateway_state()
    state.update(updates)
    state["updated_at"] = datetime.now(timezone.utc).isoformat()
    config.GATEWAY_STATE_FILE.write_text(json.dumps(state, indent=2, sort_keys=True) + "\n")
    config.GATEWAY_STATE_FILE.chmod(0o600)


def _session_dir_for_chat(chat_id: str) -> Path:
    digest = hashlib.sha1(chat_id.encode("utf-8")).hexdigest()[:16]
    return config.SESSION_DIR / "gateway_sessions" / digest


def _workspace_dir_for_chat(chat_id: str) -> Path:
    digest = hashlib.sha1(chat_id.encode("utf-8")).hexdigest()[:16]
    return config.GATEWAY_WORKSPACE_DIR / digest


def _chat_digest(chat_id: str) -> str:
    return hashlib.sha1(chat_id.encode("utf-8")).hexdigest()[:16]


def _safe_export_name(name: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9._ -]+", "_", Path(name).name).strip(" .")
    return safe or "marlow-export"


def _safe_workspace_file(chat_id: str, relpath: str) -> Path:
    workspace = _workspace_dir_for_chat(chat_id).resolve()
    path = Path(relpath)
    if path.is_absolute():
        raise ValueError("absolute paths are not allowed")
    resolved = (workspace / path).resolve()
    if not resolved.is_relative_to(workspace):
        raise ValueError("path escapes the gateway workspace")
    if not resolved.exists():
        raise ValueError("file does not exist")
    if not resolved.is_file():
        raise ValueError("only regular files can be sent")
    if resolved.name.startswith("."):
        raise ValueError("hidden files cannot be sent")
    if resolved.suffix.lower() not in EXPORT_ALLOWED_SUFFIXES:
        allowed = ", ".join(sorted(EXPORT_ALLOWED_SUFFIXES))
        raise ValueError(f"unsupported file type {resolved.suffix!r}; allowed: {allowed}")
    size = resolved.stat().st_size
    if size > EXPORT_MAX_BYTES:
        raise ValueError(f"file is too large ({size} bytes, max {EXPORT_MAX_BYTES})")
    return resolved


def _workspace_relpath(chat_id: str, path: Path) -> str:
    return path.resolve().relative_to(_workspace_dir_for_chat(chat_id).resolve()).as_posix()


def _configure_pi_session(sess: PiSession, chat_id: str) -> None:
    state = _read_gateway_state()
    sess.workspace_dir = _workspace_dir_for_chat(chat_id)
    sess.provider = str(state.get("pi_provider") or "") or None
    sess.model = str(state.get("pi_model") or "") or None
    sess.thinking = str(state.get("pi_thinking") or "") or None
    sess.models = str(state.get("pi_models") or "") or None


def _format_pi_settings(chat_id: str) -> str:
    state = _read_gateway_state()
    values = {
        "Provider": state.get("pi_provider") or "(pi default)",
        "Model": state.get("pi_model") or "(pi default)",
        "Thinking": state.get("pi_thinking") or "(pi default)",
        "Model cycle": state.get("pi_models") or "(pi default)",
        "Workspace": str(_workspace_dir_for_chat(chat_id)),
        "Sessions": str(_session_dir_for_chat(chat_id)),
    }
    return "\n".join(f"{key}: {value}" for key, value in values.items())


def _display_time(value: str) -> str:
    if "T" in value and len(value) >= 16:
        return value[11:16]
    return value[:16] if value else ""


def _parse_graph_time(value: str) -> datetime:
    if not value:
        return datetime.min.replace(tzinfo=timezone.utc)
    normalized = value.replace("Z", "+00:00")
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError:
        return datetime.min.replace(tzinfo=timezone.utc)
    if parsed.tzinfo is None:
        return parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(timezone.utc)


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _format_age(value: str) -> str:
    parsed = _parse_graph_time(value)
    if parsed == datetime.min.replace(tzinfo=timezone.utc):
        return "unknown"
    seconds = max(0, int((datetime.now(timezone.utc) - parsed).total_seconds()))
    if seconds < 90:
        return f"{seconds}s ago"
    minutes = seconds // 60
    if minutes < 90:
        return f"{minutes}m ago"
    hours = minutes // 60
    if hours < 48:
        return f"{hours}h ago"
    days = hours // 24
    return f"{days}d ago"


def _heartbeat_is_stale(state: dict, running: bool) -> bool:
    if not running:
        return False
    last_poll = str(state.get("last_poll_at") or "")
    parsed = _parse_graph_time(last_poll)
    if parsed == datetime.min.replace(tzinfo=timezone.utc):
        return False
    poll_interval = int(state.get("poll_interval") or config.GATEWAY_POLL_INTERVAL)
    stale_after = max(120, poll_interval * 4)
    return (datetime.now(timezone.utc) - parsed).total_seconds() > stale_after


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
    if getattr(args, "self_chat", False):
        console.print("[dim]Finding Teams self-chat...[/]")
        client = _graph_client()
        try:
            found = find_self_chat(client)
        finally:
            client.close()
        if not found:
            console.print("[red]No Teams self-chat found in Microsoft Graph.[/]")
            sys.exit(1)
        chat, members = found
        chat_id = SELF_CHAT_ID
        _write_gateway_state(
            chat_id=chat_id,
            chat_label=_chat_title(chat, members),
            last_seen_time=datetime.now(timezone.utc).isoformat(),
        )
        config.GATEWAY_CHAT_ID_FILE.write_text(chat_id)
        config.GATEWAY_CHAT_ID_FILE.chmod(0o600)
        return chat_id

    chat_id = getattr(args, "chat_id", None) or config.GATEWAY_CHAT_ID
    if chat_id:
        _write_gateway_state(chat_id=chat_id)
        return chat_id
    state_chat_id = str(_read_gateway_state().get("chat_id") or "").strip()
    if state_chat_id:
        return state_chat_id
    if config.GATEWAY_CHAT_ID_FILE.exists():
        stored = config.GATEWAY_CHAT_ID_FILE.read_text().strip()
        if stored:
            _write_gateway_state(chat_id=stored)
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
    selected_label = next(
        (label for label, chat in labels if chat is chosen),
        _chat_label(chosen, None),
    )
    config.GATEWAY_CHAT_ID_FILE.write_text(selected_id)
    config.GATEWAY_CHAT_ID_FILE.chmod(0o600)
    _write_gateway_state(chat_id=selected_id, chat_label=selected_label)
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
    session_dir = _session_dir_for_chat(chat_id)
    sess = PiSession(session_dir, log_fn=_log)
    _configure_pi_session(sess, chat_id)
    sess.start()
    _pi_sessions[chat_id] = sess
    return sess


def _new_pi_session(chat_id: str, *, clear: bool = False) -> str:
    sess = _pi_sessions.pop(chat_id, None)
    if sess:
        try:
            sess.close()
        except Exception:
            pass

    session_dir = _session_dir_for_chat(chat_id)
    if clear and session_dir.exists():
        shutil.rmtree(session_dir)
    session_dir.mkdir(parents=True, exist_ok=True)

    sess = PiSession(session_dir, log_fn=_log, resume=False)
    _configure_pi_session(sess, chat_id)
    sess.start()
    sess._preamble_sent = False  # type: ignore[attr-defined]
    _pi_sessions[chat_id] = sess
    _last_marlow_turn_time.pop(chat_id, None)
    _write_gateway_state(
        active_session_dir=str(session_dir),
        active_workspace_dir=str(_workspace_dir_for_chat(chat_id)),
        last_command="reset" if clear else "new",
    )
    return str(session_dir)


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

WORKSPACE GUIDANCE
- Your shell working directory is a per-chat gateway workspace, not a product repo.
- Do not infer Ross's whole day from the current directory.
- If you create files that should be sent back to Teams, keep them under this
  workspace and write .marlow-export.json with:
  {{"files":["relative/path.ext"],"message":"Created the report."}}
  Use only relative paths. The gateway will upload and post links after your response.
- For broad questions about the day, use relevant records first: calendar, Teams,
  mail, Workbook timesheets, and only inspect git repos when the user asks about
  code work or names a project.

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
        if _parse_graph_time(m_time) <= _parse_graph_time(since_time):
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
    progress_fn=None,
) -> str:
    """Send a prompt to the per-chat persistent pi session."""
    try:
        sess = _get_pi_session(chat_id)
        is_first = (
            not getattr(sess, "resumed_existing", False)
            and not getattr(sess, "_preamble_sent", False)
        )

        display_sender_time = _display_time(sender_time)

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
                + f"\n\nLATEST MESSAGE\n[{display_sender_time}] {sender}: {prompt}"
            )
            sess._preamble_sent = True  # type: ignore[attr-defined]
        else:
            # Build catch-up block of any chat that happened since our last turn
            since = _last_marlow_turn_time.get(chat_id, "")
            catchup = ""
            if since and recent_messages:
                catchup = _format_catchup(recent_messages, since, trigger_msg_id)

            if catchup:
                catchup_lines = catchup.count("\n") + 1
                _log(f"Catch-up: {catchup_lines} message(s) since last Marlow turn")
                full_prompt = (
                    "CATCH-UP (chat that happened while you were away)\n"
                    f"{catchup}\n\n"
                    f"LATEST MESSAGE\n[{display_sender_time}] {sender}: {prompt}"
                )
            else:
                full_prompt = f"[{display_sender_time}] {sender}: {prompt}"

        response = sess.prompt(full_prompt, progress_fn=progress_fn)
        # Mark this trigger time as the last Marlow turn for catch-up tracking
        if sender_time:
            _last_marlow_turn_time[chat_id] = sender_time
            _write_gateway_state(last_marlow_turn_time=sender_time)
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


def _format_gateway_status(chat_id: str, trigger: str, poll_interval: int) -> str:
    pid = _read_pid()
    running = bool(pid and _process_is_running(pid))
    state = _read_gateway_state()
    lines = [
        f"Gateway: {'running' if running else 'not running'}" + (f" (PID {pid})" if pid else ""),
        f"Chat: {state.get('chat_label') or chat_id}",
        f"Trigger: {trigger}",
        f"Poll: every {poll_interval}s",
        f"Paused: {'yes' if state.get('paused') else 'no'}",
    ]
    if state.get("started_at"):
        lines.append(f"Started: {state['started_at']} ({_format_age(str(state['started_at']))})")
    if state.get("last_poll_at"):
        lines.append(f"Last poll: {state['last_poll_at']} ({_format_age(str(state['last_poll_at']))})")
    if _heartbeat_is_stale(state, running):
        lines.append("Warning: gateway heartbeat is stale")
    if state.get("last_seen_time"):
        lines.append(f"Last seen: {state['last_seen_time']}")
    if state.get("last_marlow_turn_time"):
        lines.append(f"Last Marlow turn: {state['last_marlow_turn_time']}")
    if state.get("auth_error"):
        lines.append(f"Auth error: {state['auth_error']}")
    if state.get("stopped_reason"):
        lines.append(f"Stopped reason: {state['stopped_reason']}")
    if state.get("stopped_at"):
        lines.append(f"Stopped at: {state['stopped_at']} ({_format_age(str(state['stopped_at']))})")
    if state.get("last_poll_error"):
        lines.append(f"Last poll error: {state['last_poll_error']}")
    lines.append(_format_pi_settings(chat_id))
    return "\n".join(lines)


def _recent_gateway_logs(limit: int = 10) -> str:
    if not config.GATEWAY_LOG_FILE.exists():
        return "No gateway log yet."
    lines = config.GATEWAY_LOG_FILE.read_text().splitlines()[-limit:]
    return "\n".join(lines) if lines else "No gateway log yet."


def _list_workspace_files(chat_id: str, *, limit: int = 20) -> str:
    workspace = _workspace_dir_for_chat(chat_id)
    if not workspace.exists():
        return "No gateway workspace exists for this chat yet."
    files = [
        path
        for path in workspace.rglob("*")
        if path.is_file()
        and not path.name.startswith(".")
        and path.suffix.lower() in EXPORT_ALLOWED_SUFFIXES
    ]
    files.sort(key=lambda path: path.stat().st_mtime, reverse=True)
    if not files:
        return "No sendable files found in this chat workspace."
    lines = ["Sendable files in this chat workspace:"]
    for path in files[:limit]:
        stat = path.stat()
        relpath = _workspace_relpath(chat_id, path)
        modified = datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d %H:%M")
        lines.append(f"{relpath} ({stat.st_size} bytes, modified {modified})")
    if len(files) > limit:
        lines.append(f"... {len(files) - limit} more")
    lines.append("Use !send <path> to publish one of these files.")
    return "\n".join(lines)


def _read_export_manifest(chat_id: str) -> tuple[list[str], str] | None:
    manifest_path = _workspace_dir_for_chat(chat_id) / EXPORT_MANIFEST
    if not manifest_path.exists():
        return None
    try:
        data = json.loads(manifest_path.read_text())
    except (OSError, json.JSONDecodeError) as exc:
        raise ValueError(f"could not read {EXPORT_MANIFEST}: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError(f"{EXPORT_MANIFEST} must be a JSON object")
    files = data.get("files")
    if not isinstance(files, list) or not files:
        raise ValueError(f"{EXPORT_MANIFEST} must include a non-empty files list")
    paths: list[str] = []
    for item in files:
        if not isinstance(item, str) or not item.strip():
            raise ValueError(f"{EXPORT_MANIFEST} files must be strings")
        paths.append(item.strip())
    message = data.get("message")
    if message is not None and not isinstance(message, str):
        raise ValueError(f"{EXPORT_MANIFEST} message must be a string")
    return paths, (message or "Generated file(s):").strip()


def _remote_export_path(chat_id: str, local_path: Path) -> str:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    filename = _safe_export_name(local_path.name)
    parts = [*EXPORT_REMOTE_ROOT.split("/"), _chat_digest(chat_id), f"{stamp}-{filename}"]
    return "/".join(quote(part.strip("/"), safe="") for part in parts)


def _ensure_onedrive_folder(client: OutlookClient, path: str) -> None:
    parent = client._onedrive_item_by_path_internal("")
    current = ""
    for raw_part in path.strip("/").split("/"):
        part = raw_part.strip()
        if not part:
            continue
        current = f"{current}/{part}".strip("/")
        try:
            parent = client._onedrive_item_by_path_internal(current)
            continue
        except OutlookAPIError as exc:
            if exc.status != 404:
                raise
        try:
            parent = client._onedrive_create_folder_internal(parent["id"], part)
        except OutlookAPIError as exc:
            if exc.status != 409:
                raise
            parent = client._onedrive_item_by_path_internal(current)


def _publish_workspace_files(chat_id: str, relpaths: list[str], message: str = "Generated file(s):") -> str:
    if not relpaths:
        return "No files were provided."
    paths = [_safe_workspace_file(chat_id, relpath) for relpath in relpaths]
    client = _graph_client()
    try:
        _ensure_onedrive_folder(client, f"{EXPORT_REMOTE_ROOT}/{_chat_digest(chat_id)}")
        lines = [message]
        for path in paths:
            remote_path = _remote_export_path(chat_id, path)
            item = client._onedrive_upload_file_internal(remote_path, path.read_bytes())
            link_url = item.get("webUrl") or ""
            try:
                permission = client._onedrive_create_share_link_internal(item["id"])
                link_url = ((permission.get("link") or {}).get("webUrl")) or link_url
            except OutlookAPIError as exc:
                _log(f"Could not create share link for {path.name}: {exc}")
            relpath = _workspace_relpath(chat_id, path)
            if link_url:
                lines.append(f"{relpath}: {link_url}")
            else:
                lines.append(f"{relpath}: uploaded, but Graph did not return a link")
        return "\n".join(lines)
    finally:
        client.close()


def _publish_export_manifest(chat_id: str) -> str | None:
    manifest_path = _workspace_dir_for_chat(chat_id) / EXPORT_MANIFEST
    manifest = _read_export_manifest(chat_id)
    if manifest is None:
        return None
    relpaths, message = manifest
    try:
        return _publish_workspace_files(chat_id, relpaths, message)
    finally:
        manifest_path.unlink(missing_ok=True)


def _handle_gateway_command(chat_id: str, command_text: str, trigger: str, poll_interval: int) -> str | None:
    stripped = command_text.strip()
    if not stripped.startswith("!"):
        return None

    parts = stripped[1:].split()
    command = parts[0].lower() if parts else "help"

    if command in {"help", "commands"}:
        topic = parts[1].lower() if len(parts) > 1 else ""
        return _gateway_command_help(topic)
    if command == "status":
        return _format_gateway_status(chat_id, trigger, poll_interval)
    if command == "new":
        try:
            session_dir = _new_pi_session(chat_id, clear=False)
            _log(f"Gateway command !new: fresh session in {session_dir}")
            return "Started a fresh Pi conversation for this chat."
        except PiSessionError as exc:
            _log(f"Gateway command !new failed: {exc}")
            return f"Could not start a fresh Pi conversation: {exc}"
    if command == "reset":
        try:
            session_dir = _new_pi_session(chat_id, clear=True)
            _log(f"Gateway command !reset: cleared session in {session_dir}")
            return "Reset this chat's Pi session and started fresh."
        except PiSessionError as exc:
            _log(f"Gateway command !reset failed: {exc}")
            return f"Could not reset the Pi conversation: {exc}"
    if command == "model":
        return _handle_model_command(chat_id, stripped)
    if command == "pause":
        _write_gateway_state(paused=True, last_command="pause")
        _log("Gateway command !pause: normal prompts paused")
        return "Paused normal Marlow prompts in this chat. Use @Marlow !resume to turn them back on."
    if command == "resume":
        _write_gateway_state(paused=False, last_command="resume")
        _log("Gateway command !resume: normal prompts resumed")
        return "Resumed normal Marlow prompts in this chat."
    if command == "tools":
        return (
            "Marlow can use the Pi agent tools available on this machine, plus installed CLIs "
            "such as outlook-cli and workbook-cli. Gateway commands use !command after @Marlow."
        )
    if command == "files":
        return _list_workspace_files(chat_id)
    if command == "send":
        try:
            send_tokens = shlex.split(stripped)
        except ValueError as exc:
            return f"Could not parse send command: {exc}"
        if len(send_tokens) < 2:
            return _gateway_command_help("send")
        try:
            return _publish_workspace_files(chat_id, [" ".join(send_tokens[1:])], "Published file:")
        except (OSError, OutlookAPIError, ValueError) as exc:
            _log(f"Gateway command !send failed: {exc}")
            return f"Could not send file: {exc}"
    if command == "logs":
        return _recent_gateway_logs(limit=8)
    if command in {"compact", "prune"}:
        return "Context pruning is not wired into the gateway yet. Use !new or !reset for now."

    return f"Unknown gateway command: !{command}. Try !help."


def _handle_model_command(chat_id: str, command_text: str) -> str:
    try:
        tokens = shlex.split(command_text)
    except ValueError as exc:
        return f"Could not parse model command: {exc}"

    args = tokens[1:]
    if not args:
        return "Current Pi settings:\n" + _format_pi_settings(chat_id)

    if args[0] in {"help", "-h", "--help"}:
        return _model_help()

    if args[0] == "list":
        search = " ".join(args[1:]).strip()
        return _list_pi_models(search or None)

    if args == ["reset"]:
        _write_gateway_state(
            pi_provider=None,
            pi_model=None,
            pi_thinking=None,
            pi_models=None,
            last_command="model reset",
        )
        old = _pi_sessions.pop(chat_id, None)
        if old:
            old.close()
        return "Reset Pi model settings to pi defaults. The next prompt will restart Pi with the default model."

    provider: str | None = None
    model: str | None = None
    thinking: str | None = None
    models: str | None = None

    i = 0
    while i < len(args):
        token = args[i]
        if token in {"--provider", "-p"} and i + 1 < len(args):
            provider = args[i + 1]
            i += 2
            continue
        if token in {"--model", "-m"} and i + 1 < len(args):
            model = args[i + 1]
            i += 2
            continue
        if token == "--thinking" and i + 1 < len(args):
            thinking = args[i + 1]
            i += 2
            continue
        if token == "--models" and i + 1 < len(args):
            models = args[i + 1]
            i += 2
            continue
        if model is None:
            model = token
            i += 1
            continue
        return f"Unknown model option: {token}"

    updates: dict[str, object] = {"last_command": "model"}
    if provider is not None:
        updates["pi_provider"] = provider
    if model is not None:
        updates["pi_model"] = model
    if thinking is not None:
        updates["pi_thinking"] = thinking
    if models is not None:
        updates["pi_models"] = models
    _write_gateway_state(**updates)

    old = _pi_sessions.pop(chat_id, None)
    if old:
        old.close()

    return "Updated Pi settings. The next prompt will restart Pi with:\n" + _format_pi_settings(chat_id)


def _gateway_command_help(topic: str = "") -> str:
    if topic in {"", "commands"}:
        return (
            "Marlow gateway commands:\n"
            "!help [command] - show command help\n"
            "!status - show gateway, model, session, and workspace status\n"
            "!new - start a fresh Pi conversation for this chat\n"
            "!reset - clear this chat's Pi session and start fresh\n"
            "!model - show or change the Pi model for this chat\n"
            "!pause - ignore normal prompts until resumed\n"
            "!resume - resume normal prompt handling\n"
            "!tools - show what Marlow can use\n"
            "!files - list sendable files in this chat workspace\n"
            "!send <path> - publish a workspace file back to Teams\n"
            "!logs - show recent gateway log lines"
        )
    if topic == "model":
        return _model_help()
    if topic == "status":
        return "!status\nShows gateway state, current model settings, workspace, session path, auth errors, and recent activity."
    if topic == "new":
        return "!new\nStarts a fresh Pi conversation for this chat without deleting old session files."
    if topic == "reset":
        return "!reset\nClears this chat's persisted Pi session and starts fresh."
    if topic == "pause":
        return "!pause\nStops normal prompts from going to Marlow. Gateway commands still work."
    if topic == "resume":
        return "!resume\nResumes normal Marlow prompt handling after !pause."
    if topic == "tools":
        return "!tools\nShows the local tools and CLIs Marlow can use."
    if topic == "files":
        return "!files\nLists recent sendable files in this chat's gateway workspace."
    if topic == "send":
        return "!send <path>\nUploads a file from this chat's gateway workspace to OneDrive and posts a link. Paths must be relative to the workspace."
    if topic == "logs":
        return "!logs\nShows recent gateway log lines."
    return f"No help available for !{topic}. Try !help."


def _model_help() -> str:
    return (
        "!model commands:\n"
        "!model - show current Pi model settings\n"
        "!model help - show this help\n"
        "!model list [search] - list available Pi models, optionally filtered\n"
        "!model <model> - set model by Pi fuzzy/pattern match\n"
        "!model <model>:<thinking> - set model with thinking shorthand, e.g. sonnet:high\n"
        "!model --provider <provider> --model <model> --thinking <level> - set explicit provider/model/thinking\n"
        "!model reset - return to Pi defaults\n\n"
        "Tip: short names like sonnet are convenient but can be ambiguous. Use !model list sonnet, then set an exact model ID when repeatability matters."
    )


def _list_pi_models(search: str | None = None) -> str:
    pi_bin = shutil.which("pi")
    if not pi_bin:
        return "pi not found in PATH."

    cmd = [pi_bin, "--list-models"]
    if search:
        cmd.append(search)
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=20,
        )
    except subprocess.TimeoutExpired:
        return "Timed out while listing Pi models."
    except Exception as exc:
        return f"Could not list Pi models: {exc}"

    output = (result.stdout or result.stderr or "").strip()
    if result.returncode != 0:
        return f"Could not list Pi models:\n{output[:2500] or f'exit {result.returncode}'}"
    if not output:
        return "No Pi models returned."
    if len(output) > 3000:
        output = output[:3000].rstrip() + "\n..."
    label = f"Pi models matching '{search}':" if search else "Pi models:"
    return f"{label}\n{output}"


def _chunk_response(response: str, *, max_chars: int = 3500) -> list[str]:
    if len(response) <= max_chars:
        return [response]

    chunks: list[str] = []
    remaining = response
    while remaining:
        if len(remaining) <= max_chars:
            chunks.append(remaining)
            break
        split_at = remaining.rfind("\n", 0, max_chars)
        if split_at < max_chars // 2:
            split_at = remaining.rfind(" ", 0, max_chars)
        if split_at < max_chars // 2:
            split_at = max_chars
        chunks.append(remaining[:split_at].strip())
        remaining = remaining[split_at:].strip()
    return [chunk for chunk in chunks if chunk]


def _post_gateway_response(chat_id: str, response: str) -> None:
    for chunk in _chunk_response(response):
        _post_gateway_response_chunk(chat_id, chunk)


def _post_gateway_response_chunk(chat_id: str, response: str) -> dict:
    html_body = (
        "<div style='border-left:3px solid #6264a7;"
        "padding:6px 12px;background:#f3f2f1;'>"
        "<strong>\U0001f916 Marlow</strong><br>"
        f"{html.escape(response).replace(chr(10), '<br>')}"
        "</div>"
    )
    post_client = _graph_client()
    try:
        return post_client._send_teams_message_internal(
            chat_id,
            html_body,
            content_type="html",
        )
    finally:
        post_client.close()


def _post_gateway_status_message(chat_id: str, response: str) -> dict | None:
    try:
        return _post_gateway_response_chunk(chat_id, response)
    except Exception as exc:
        _log(f"Failed to post status message: {exc}")
        return None


def _delete_gateway_message(chat_id: str, message: dict | None) -> None:
    message_id = (message or {}).get("id")
    if not message_id:
        return
    client = _graph_client()
    try:
        user_id = client.get_current_user().get("id", "")
        if not user_id:
            return
        client._soft_delete_teams_message_internal(chat_id, message_id, user_id)
        _log(f"Status message deleted: {message_id}")
    except Exception as exc:
        _log(f"Failed to delete status message {message_id}: {exc}")
    finally:
        client.close()


def _poll_loop(chat_id: str, trigger: str, poll_interval: int) -> None:
    trigger_re = re.compile(re.escape(trigger), re.IGNORECASE)
    state = _read_gateway_state()
    last_seen_time = state.get("last_seen_time") or datetime.now(timezone.utc).isoformat()
    seen_ids: set[str] = set()

    _log(f"Gateway started. chat={chat_id} trigger={trigger} poll={poll_interval}s")
    _write_gateway_state(
        chat_id=chat_id,
        trigger=trigger,
        poll_interval=poll_interval,
        last_seen_time=last_seen_time,
        process_status="running",
        started_at=_utc_now_iso(),
        stopped_at=None,
        stopped_reason=None,
        last_poll_at=None,
        last_poll_error=None,
    )

    while True:
        try:
            _write_gateway_state(last_poll_at=_utc_now_iso(), last_poll_error=None)
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
                if _parse_graph_time(msg_time) <= _parse_graph_time(last_seen_time):
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
                last_seen_time = msg_time
                _write_gateway_state(last_seen_time=last_seen_time)

                command_response = _handle_gateway_command(
                    chat_id,
                    prompt,
                    trigger,
                    poll_interval,
                )
                if command_response is not None:
                    try:
                        _post_gateway_response(chat_id, command_response)
                        _log(f"Gateway command response posted ({len(command_response)} chars).")
                    except Exception as exc:
                        _log(f"Failed to post gateway command response: {exc}")
                    continue

                if _read_gateway_state().get("paused"):
                    _log("Gateway is paused; normal prompt ignored")
                    continue

                receipt_message = _post_gateway_status_message(chat_id, "...")
                progress_seen: set[str] = set()

                def progress_update(line: str) -> None:
                    if line in progress_seen:
                        return
                    progress_seen.add(line)
                    _log(f"Pi progress: {line}")
                    _post_gateway_status_message(chat_id, f"... {line}")

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
                    sender_time=msg_time,
                    recent_messages=messages,
                    trigger_msg_id=msg_id,
                    progress_fn=progress_update,
                )
                _log(f"Response ({len(response)} chars): {response[:80]}")

                try:
                    _delete_gateway_message(chat_id, receipt_message)
                    _post_gateway_response(chat_id, response)
                    try:
                        publish_response = _publish_export_manifest(chat_id)
                        if publish_response:
                            _post_gateway_response(chat_id, publish_response)
                            _log("Export manifest published to Teams.")
                    except Exception as exc:
                        _log(f"Failed to publish export manifest: {exc}")
                        _post_gateway_response(chat_id, f"Could not publish generated files: {exc}")
                    _log("Response posted to Teams.")
                except Exception as exc:
                    _log(f"Failed to post response: {exc}")

        except OutlookAPIError as exc:
            _log(f"Graph API error: {exc}")
            _write_gateway_state(last_poll_error=f"Graph API error: {exc}")
        except (TokenExpiredError, TokenNotFoundError) as exc:
            _log(f"Graph auth error: {exc}")
            _write_gateway_state(
                auth_error=str(exc),
                process_status="stopped",
                stopped_at=_utc_now_iso(),
                stopped_reason="graph_auth_error",
                last_poll_error=f"Graph auth error: {exc}",
            )
            return
        except Exception as exc:
            _log(f"Unexpected poll error: {exc}")
            _write_gateway_state(last_poll_error=f"Unexpected poll error: {exc}")

        time.sleep(poll_interval)


def cmd_gateway_start(args: argparse.Namespace) -> None:
    console = Console()

    existing_pid = _read_pid()
    if existing_pid and _process_is_running(existing_pid):
        console.print(f"[yellow]Gateway already running (PID {existing_pid}).[/]")
        return
    if existing_pid:
        _write_gateway_state(
            process_status="stopped",
            stopped_at=_utc_now_iso(),
            stopped_reason=f"stale_pid:{existing_pid}",
        )
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)

    chat_id = _resolve_chat_id(args, console)
    trigger = getattr(args, "trigger", None) or config.GATEWAY_TRIGGER
    poll_interval = getattr(args, "poll", None) or config.GATEWAY_POLL_INTERVAL
    session_dir = _session_dir_for_chat(chat_id)
    workspace_dir = _workspace_dir_for_chat(chat_id)
    session_dir.mkdir(parents=True, exist_ok=True)
    workspace_dir.mkdir(parents=True, exist_ok=True)
    state_updates: dict[str, object] = dict(
        chat_id=chat_id,
        trigger=trigger,
        poll_interval=poll_interval,
        auth_error=None,
        stopped_reason=None,
        stopped_at=None,
        process_status="starting",
        active_session_dir=str(session_dir),
        active_workspace_dir=str(workspace_dir),
    )
    if getattr(args, "provider", None):
        state_updates["pi_provider"] = args.provider
    if getattr(args, "model", None):
        state_updates["pi_model"] = args.model
    if getattr(args, "thinking", None):
        state_updates["pi_thinking"] = args.thinking
    if getattr(args, "models", None):
        state_updates["pi_models"] = args.models
    _write_gateway_state(**state_updates)

    console.print("[green]Starting gateway...[/]")
    console.print(f"  Chat:    [bold]{chat_id}[/]")
    console.print(f"  Trigger: [bold]{trigger}[/]")
    console.print(f"  Poll:    every [bold]{poll_interval}s[/]")
    console.print(f"  Model:   [bold]{_read_gateway_state().get('pi_model') or '(pi default)'}[/]")
    console.print(f"  Workspace: {workspace_dir}")
    console.print(f"  Logs:    {config.GATEWAY_LOG_FILE}")

    pid = os.fork()
    if pid > 0:
        config.SESSION_DIR.mkdir(parents=True, exist_ok=True)
        config.GATEWAY_PID_FILE.write_text(str(pid))
        _write_gateway_state(pid=pid, process_status="running")
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
        _write_gateway_state(
            process_status="stopped",
            stopped_at=_utc_now_iso(),
            stopped_reason=f"crashed:{exc}",
        )
    finally:
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)
    os._exit(0)


def cmd_gateway_stop(args: argparse.Namespace) -> None:
    console = Console()
    pid = _read_pid()
    if not pid:
        _write_gateway_state(process_status="stopped", stopped_at=_utc_now_iso(), stopped_reason="manual_stop_no_pid")
        console.print("[yellow]Gateway is not running (no PID file).[/]")
        return
    if not _process_is_running(pid):
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)
        _write_gateway_state(process_status="stopped", stopped_at=_utc_now_iso(), stopped_reason=f"stale_pid:{pid}")
        console.print("[yellow]Gateway process not found. PID file cleaned up.[/]")
        return
    try:
        os.kill(pid, signal.SIGTERM)
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)
        _write_gateway_state(process_status="stopped", stopped_at=_utc_now_iso(), stopped_reason="manual_stop")
        console.print(f"[green]Gateway stopped (PID {pid}).[/]")
    except ProcessLookupError:
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)
        _write_gateway_state(process_status="stopped", stopped_at=_utc_now_iso(), stopped_reason=f"stale_pid:{pid}")
        console.print("[yellow]Process already gone.[/]")


def cmd_gateway_status(args: argparse.Namespace) -> None:
    console = Console()
    pid = _read_pid()
    running = bool(pid and _process_is_running(pid))
    stale_pid = bool(pid and not running)
    if stale_pid:
        config.GATEWAY_PID_FILE.unlink(missing_ok=True)
        _write_gateway_state(
            process_status="stopped",
            stopped_at=_utc_now_iso(),
            stopped_reason=f"stale_pid:{pid}",
        )

    chat_id = config.GATEWAY_CHAT_ID
    state = _read_gateway_state()
    if not chat_id:
        chat_id = str(state.get("chat_id") or "")
    if not chat_id and config.GATEWAY_CHAT_ID_FILE.exists():
        chat_id = config.GATEWAY_CHAT_ID_FILE.read_text().strip()

    if running:
        console.print(f"[green]Gateway is running[/] (PID {pid})")
    elif stale_pid:
        console.print(f"[yellow]Gateway is not running.[/] Stale PID file cleaned up (PID {pid}).")
    else:
        console.print("[dim]Gateway is not running.[/]")

    if chat_id:
        console.print(f"  Chat:    {state.get('chat_label') or chat_id}")
        if state.get("chat_label"):
            console.print(f"  Chat ID: {chat_id}")
    console.print(f"  Trigger: {state.get('trigger') or config.GATEWAY_TRIGGER}")
    console.print(f"  Poll:    every {state.get('poll_interval') or config.GATEWAY_POLL_INTERVAL}s")
    if state.get("last_seen_time"):
        console.print(f"  Last seen: {state['last_seen_time']}")
    if state.get("started_at"):
        console.print(f"  Started: {state['started_at']} ({_format_age(str(state['started_at']))})")
    if state.get("last_poll_at"):
        console.print(f"  Last poll: {state['last_poll_at']} ({_format_age(str(state['last_poll_at']))})")
    if _heartbeat_is_stale(state, running):
        console.print("  [yellow]Warning: gateway heartbeat is stale.[/]")
    if state.get("last_marlow_turn_time"):
        console.print(f"  Last Marlow turn: {state['last_marlow_turn_time']}")
    if state.get("auth_error"):
        console.print(f"  Auth error: {state['auth_error']}")
    if state.get("stopped_reason"):
        console.print(f"  Stopped reason: {state['stopped_reason']}")
    if state.get("stopped_at"):
        console.print(f"  Stopped at: {state['stopped_at']} ({_format_age(str(state['stopped_at']))})")
    if state.get("last_poll_error"):
        console.print(f"  Last poll error: {state['last_poll_error']}")
    if state.get("paused"):
        console.print("  Paused:  yes")
    console.print(f"  Model:   {state.get('pi_model') or '(pi default)'}")
    if state.get("pi_provider"):
        console.print(f"  Provider: {state['pi_provider']}")
    if state.get("pi_thinking"):
        console.print(f"  Thinking: {state['pi_thinking']}")
    if state.get("pi_models"):
        console.print(f"  Model cycle: {state['pi_models']}")
    if chat_id:
        console.print(f"  Workspace: {_workspace_dir_for_chat(chat_id)}")
        console.print(f"  Sessions:  {_session_dir_for_chat(chat_id)}")
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
