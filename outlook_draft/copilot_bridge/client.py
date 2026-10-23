"""Unix socket client shared by CLI and MCP. No automatic retries."""
from __future__ import annotations

import os
import socket
import stat
from collections.abc import Iterator
from pathlib import Path

from ..errors import CopilotError
from .protocol import MAX_MESSAGE, VERSION, decode_message, encode_message, socket_path, timeout_seconds


class BridgeClient:
    def __init__(self, path: Path | None = None):
        self.path = path or socket_path()

    def request(self, message: dict, *, timeout: float = 30) -> Iterator[dict]:
        try:
            info = self.path.lstat()
        except FileNotFoundError:
            raise CopilotError("Copilot bridge is not running. Open Chrome and load the extension. Run 'outlook-cli copilot extension install' for setup.") from None
        if not stat.S_ISSOCK(info.st_mode) or info.st_uid != os.getuid() or stat.S_IMODE(info.st_mode) != 0o600:
            raise CopilotError("Unsafe Copilot bridge socket. It must belong to this user with mode 0600.")
        try:
            yield from self._request(message, timeout)
        except (OSError, EOFError) as exc:
            raise CopilotError("Copilot bridge disconnected or timed out. The document may still be changing. Do not retry automatically.") from exc

    def _request(self, message: dict, timeout: float) -> Iterator[dict]:
        with socket.socket(socket.AF_UNIX, socket.SOCK_STREAM) as connection:
            connection.settimeout(timeout)
            connection.connect(str(self.path))
            connection.sendall(encode_message(message) + b"\n")
            with connection.makefile("rb") as stream:
                hello = self._read(stream)
                if hello.get("error"):
                    raise CopilotError(hello["error"])
                if hello.get("type") != "hello" or hello.get("protocol") != VERSION or hello.get("host") != VERSION:
                    raise CopilotError("Unsupported Copilot bridge protocol. Update the host and extension.")
                while True:
                    message = self._read(stream)
                    if message.get("error"):
                        raise CopilotError(str(message["error"]))
                    if message.get("done"):
                        yield {"type": "result", "result": message.get("result") or {}}
                        return
                    if isinstance(message.get("event"), dict):
                        yield message["event"]

    @staticmethod
    def _read(stream) -> dict:
        line = stream.readline(MAX_MESSAGE + 2)
        if not line:
            raise EOFError
        if not line.endswith(b"\n"):
            raise CopilotError("Invalid socket message framing.")
        return decode_message(line)

    def status(self) -> dict:
        for event in self.request({"type": "status"}):
            if event["type"] == "result":
                return {"host": VERSION, **event["result"]}
        raise CopilotError("The bridge did not return status.")

    def inspect(self, tab_id: int) -> dict:
        if type(tab_id) is not int or tab_id < 0:
            raise CopilotError("A non-negative Chrome tab ID is required.")
        for event in self.request({"type": "inspect", "tabId": tab_id}):
            if event["type"] == "result":
                return event["result"]
        raise CopilotError("The bridge did not return diagnostics.")

    def ask(self, prompt: str, *, file: dict, app: str, new_chat: bool = False,
            open_tab: bool = False, timeout: float = 300) -> Iterator[dict]:
        timeout = timeout_seconds(timeout)
        yield from self.request({
            "type": "ask", "prompt": prompt, "file": file, "app": app,
            "newChat": new_chat, "open": open_tab, "timeout": timeout,
        }, timeout=timeout + 160)
