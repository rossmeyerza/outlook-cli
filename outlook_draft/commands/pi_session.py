"""Persistent pi --mode rpc session client.

Keeps one long-running `pi --mode rpc` subprocess per Teams chat. Sessions
persist to disk under --session-dir so they survive gateway restarts.
Responses are accumulated from message_update text_delta events and returned
when agent_end fires.
"""

from __future__ import annotations

import json
import os
import shutil
import queue
import subprocess
import threading
import time
from pathlib import Path
from typing import Callable


class PiSessionError(Exception):
    pass


class PiSession:
    """Wraps a single long-running `pi --mode rpc` subprocess.

    Thread safety: only one outstanding prompt() at a time per instance.
    """

    def __init__(self, session_dir: Path, log_fn=None, *, resume: bool = True):
        self.session_dir = session_dir
        self.session_dir.mkdir(parents=True, exist_ok=True)
        self._log = log_fn or (lambda msg: None)
        self._resume = resume
        self.resumed_existing = False
        self._proc: subprocess.Popen | None = None
        self._reader_thread: threading.Thread | None = None
        self._events: queue.Queue = queue.Queue()
        self._closed = False

    # ------------------------------------------------------------------
    # Lifecycle
    # ------------------------------------------------------------------

    def start(self) -> None:
        if self._proc is not None:
            return
        pi_bin = shutil.which("pi")
        if not pi_bin:
            raise PiSessionError("pi not found in PATH")

        cmd = [pi_bin, "--mode", "rpc", "--session-dir", str(self.session_dir)]
        # If there's an existing session jsonl in the dir, resume the latest one.
        existing = sorted(self.session_dir.glob("*.jsonl"), key=os.path.getmtime)
        if self._resume and existing:
            cmd += ["--session", str(existing[-1])]
            self.resumed_existing = True

        self._log(f"Spawning: {' '.join(cmd)}")
        self._proc = subprocess.Popen(
            cmd,
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1,
        )
        self._reader_thread = threading.Thread(target=self._read_stdout, daemon=True)
        self._reader_thread.start()
        # Brief settle so the agent is ready to accept prompts
        time.sleep(0.5)

    def close(self) -> None:
        if self._closed:
            return
        self._closed = True
        if self._proc and self._proc.stdin:
            try:
                self._proc.stdin.close()
            except Exception:
                pass
        if self._proc:
            try:
                self._proc.wait(timeout=5)
            except subprocess.TimeoutExpired:
                self._proc.terminate()
                try:
                    self._proc.wait(timeout=2)
                except subprocess.TimeoutExpired:
                    self._proc.kill()

    def is_alive(self) -> bool:
        return self._proc is not None and self._proc.poll() is None

    def has_existing_session(self) -> bool:
        """True if a session jsonl already exists in the session dir."""
        return any(self.session_dir.glob("*.jsonl"))

    # ------------------------------------------------------------------
    # Internal: stdout reader thread
    # ------------------------------------------------------------------

    def _read_stdout(self) -> None:
        assert self._proc is not None and self._proc.stdout is not None
        for line in self._proc.stdout:
            line = line.strip()
            if not line:
                continue
            try:
                event = json.loads(line)
            except json.JSONDecodeError:
                continue
            self._events.put(event)

    # ------------------------------------------------------------------
    # Prompt
    # ------------------------------------------------------------------

    def prompt(
        self,
        text: str,
        timeout: float = 240.0,
        progress_fn: Callable[[str], None] | None = None,
    ) -> str:
        """Send a prompt, wait for agent_end, return accumulated assistant text."""
        if not self.is_alive():
            self.start()

        assert self._proc is not None and self._proc.stdin is not None

        # Drain any stale events from before this prompt
        while True:
            try:
                self._events.get_nowait()
            except queue.Empty:
                break

        cmd = {"id": f"req-{int(time.time() * 1000)}", "type": "prompt", "message": text}
        try:
            self._proc.stdin.write(json.dumps(cmd) + "\n")
            self._proc.stdin.flush()
        except Exception as exc:
            raise PiSessionError(f"Failed to send prompt: {exc}") from exc

        # Collect text deltas until agent_end
        accumulated: list[str] = []
        deadline = time.monotonic() + timeout

        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                raise PiSessionError(f"Timeout after {timeout}s")
            try:
                event = self._events.get(timeout=min(remaining, 5.0))
            except queue.Empty:
                if not self.is_alive():
                    raise PiSessionError("pi process died")
                continue

            etype = event.get("type")

            if etype == "response" and event.get("command") == "prompt":
                if not event.get("success", False):
                    raise PiSessionError(f"Prompt rejected: {event}")
                continue

            if etype == "message_update":
                ev = event.get("assistantMessageEvent", {})
                if ev.get("type") == "text_delta":
                    accumulated.append(ev.get("delta", ""))
                else:
                    progress = _summarize_progress_event(ev)
                    if progress and progress_fn:
                        progress_fn(progress)
                continue

            progress = _summarize_progress_event(event)
            if progress and progress_fn:
                progress_fn(progress)

            if etype == "agent_end":
                break

            # Other events (turn_start, tool_execution_*, etc.) are ignored

        return "".join(accumulated).strip() or "[empty response]"


def _summarize_progress_event(event: dict) -> str | None:
    """Return a compact human-readable progress line for known Pi RPC events."""
    event_type = str(event.get("type") or "")
    if event_type == "message":
        message = event.get("message") or {}
        for item in message.get("content") or []:
            if item.get("type") == "toolCall":
                name = item.get("name") or "tool"
                args = item.get("arguments") or {}
                if isinstance(args, dict):
                    command = args.get("command") or args.get("cmd")
                    if command:
                        return f"Using {name}: {command}"
                return f"Using {name}"
            if item.get("type") == "toolResult":
                name = item.get("toolName") or "tool"
                if item.get("isError"):
                    return f"{name} failed"
                return f"{name} finished"

    if event_type in {"tool_call", "tool_use", "tool_execution_start"}:
        name = event.get("name") or event.get("toolName") or event.get("tool_name") or "tool"
        args = event.get("arguments") or event.get("input") or event.get("args") or {}
        if isinstance(args, dict):
            command = args.get("command") or args.get("cmd")
            if command:
                return f"Using {name}: {command}"
        return f"Using {name}"

    if event_type in {"tool_result", "tool_execution_end"}:
        name = event.get("name") or event.get("toolName") or event.get("tool_name") or "tool"
        if event.get("isError"):
            return f"{name} failed"
        return f"{name} finished"

    return None
