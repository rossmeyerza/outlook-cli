"""Chrome native host. stdout is reserved for native message framing."""
from __future__ import annotations

import fcntl
import logging
import os
import socket
import stat
import struct
import sys
import threading
import time
import uuid
from pathlib import Path
from typing import BinaryIO

from ..errors import CopilotError
from .protocol import (
    EXTENSION_ID,
    MAX_MESSAGE,
    VERSION,
    decode_message,
    encode_message,
    read_native,
    socket_path,
    timeout_seconds,
    write_native,
)

log = logging.getLogger(__name__)


def same_uid(connection: socket.socket) -> bool:
    credentials = connection.getsockopt(socket.SOL_SOCKET, socket.SO_PEERCRED, struct.calcsize("3i"))
    _, uid, _ = struct.unpack("3i", credentials)
    return uid == os.getuid()


def secure_directory(path: Path) -> None:
    path.mkdir(mode=0o700, parents=True, exist_ok=True)
    info = path.lstat()
    if not stat.S_ISDIR(info.st_mode) or info.st_uid != os.getuid():
        raise CopilotError("The Copilot bridge directory must belong to this user and must not be a symlink.")
    path.chmod(0o700)


def acquire_socket(path: Path, *, lock_timeout: float = 3) -> tuple[socket.socket, BinaryIO]:
    secure_directory(path.parent)
    flags = os.O_CREAT | os.O_RDWR | os.O_NOFOLLOW
    fd = os.open(str(path) + ".lock", flags, 0o600)
    lock = os.fdopen(fd, "wb")
    try:
        # Chrome can start the replacement before the old host finishes its
        # accept timeout. Wait briefly, but never unlink a live host's socket.
        deadline = time.monotonic() + lock_timeout
        while True:
            try:
                fcntl.flock(lock, fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except BlockingIOError:
                if time.monotonic() >= deadline:
                    raise
                time.sleep(0.05)
        if path.exists() or path.is_symlink():
            info = path.lstat()
            if not stat.S_ISSOCK(info.st_mode) or info.st_uid != os.getuid():
                raise CopilotError("Refusing to replace an unsafe Copilot bridge socket path.")
            path.unlink()
        listener = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
        listener.bind(str(path))
        path.chmod(0o600)
        listener.listen(8)
        return listener, lock
    except BaseException:
        lock.close()
        raise


class NativeHost:
    def __init__(self, stdin: BinaryIO, stdout: BinaryIO):
        self.stdin = stdin
        self.stdout = stdout
        self.write_lock = threading.Lock()
        self.routes_lock = threading.Lock()
        self.routes: dict[str, tuple[socket.socket, threading.Lock, threading.Event]] = {}
        self.connected = threading.Event()
        self.stopped = threading.Event()
        self.extension_version: str | None = None

    def send_native(self, value: dict) -> None:
        with self.write_lock:
            write_native(self.stdout, value)

    def send_socket(self, connection: socket.socket, lock: threading.Lock, value: dict) -> None:
        with lock:
            connection.sendall(encode_message(value) + b"\n")

    def read_browser(self, listener: socket.socket) -> None:
        try:
            self._browser_loop()
        except (EOFError, OSError, CopilotError):
            log.info("Chrome native channel closed.")
        finally:
            self.stopped.set()
            self.connected.clear()
            listener.close()
            with self.routes_lock:
                routes = list(self.routes.values())
            for connection, lock, done in routes:
                try:
                    self.send_socket(connection, lock, {"done": True, "error": "Chrome disconnected. The document may still be changing. Do not retry automatically."})
                except OSError:
                    log.debug("Socket client disconnected.")
                done.set()

    def _browser_loop(self) -> None:
        while not self.stopped.is_set():
            message = read_native(self.stdin)
            if message.get("type") == "hello":
                if message.get("protocol") != VERSION:
                    raise CopilotError("Unsupported extension protocol.")
                self.extension_version = str(message.get("extension", "unknown"))
                self.connected.set()
                continue
            request_id = message.get("id")
            with self.routes_lock:
                route = self.routes.get(request_id)
            if route:
                connection, lock, done = route
                try:
                    self.send_socket(connection, lock, message)
                except OSError:
                    done.set()
                if message.get("done"):
                    done.set()

    def serve_client(self, connection: socket.socket) -> None:
        request_id = uuid.uuid4().hex
        registered = False
        done = threading.Event()
        lock = threading.Lock()
        try:
            if not same_uid(connection):
                return
            connection.settimeout(5)
            with connection.makefile("rb") as stream:
                raw = stream.readline(MAX_MESSAGE + 2)
                if not raw.endswith(b"\n"):
                    raise CopilotError("Invalid socket message framing.")
                message = decode_message(raw)
                wait = self.validate_request(message)
                if not self.connected.wait(5):
                    raise CopilotError("The Chrome extension is not connected. Reload it in chrome://extensions.")
                connection.settimeout(10)
                self.send_socket(connection, lock, {"type": "hello", "host": VERSION, "extension": self.extension_version, "protocol": VERSION})
                with self.routes_lock:
                    self.routes[request_id] = (connection, lock, done)
                registered = True
                self.send_native({**message, "id": request_id})
                self._wait_client(connection, done, wait)
        except (OSError, CopilotError) as exc:
            try:
                self.send_socket(connection, lock, {"done": True, "error": str(exc)})
            except OSError:
                log.debug("Socket client disconnected before receiving the error.")
        finally:
            with self.routes_lock:
                self.routes.pop(request_id, None)
            if registered and not done.is_set() and not self.stopped.is_set():
                try:
                    self.send_native({"type": "cancel", "id": request_id})
                except OSError:
                    log.debug("Could not send cancellation to Chrome.")
            connection.close()

    @staticmethod
    def validate_request(message: dict) -> float:
        kind = message.get("type")
        if kind == "status":
            return 30
        if kind == "inspect":
            tab_id = message.get("tabId")
            if type(tab_id) is not int or tab_id < 0:
                raise CopilotError("A non-negative Chrome tab ID is required.")
            return 30
        if kind != "ask":
            raise CopilotError("Unknown bridge request.")
        prompt = message.get("prompt")
        if not isinstance(prompt, str) or not prompt.strip() or len(prompt) > 40000:
            raise CopilotError("A prompt must contain 1 to 40000 characters.")
        if message.get("app") not in {"excel", "powerpoint", "word"}:
            raise CopilotError("Choose Excel, PowerPoint or Word.")
        if not isinstance(message.get("file"), dict):
            raise CopilotError("A resolved file is required.")
        return timeout_seconds(message.get("timeout", 300)) + 150

    @staticmethod
    def _wait_client(connection: socket.socket, done: threading.Event, timeout: float) -> None:
        import select
        import time

        deadline = time.monotonic() + timeout
        while not done.wait(0.2):
            if time.monotonic() > deadline:
                raise CopilotError("Bridge request timed out. The document may still be changing. Do not retry automatically.")
            readable, _, _ = select.select([connection], [], [], 0)
            if readable:
                if not connection.recv(1, socket.MSG_PEEK):
                    raise CopilotError("Client disconnected.")
                raise CopilotError("Only one request per connection is allowed.")

    def run(self, path: Path | None = None) -> None:
        path = path or socket_path()
        listener, lock = acquire_socket(path)
        listener.settimeout(1)
        reader = threading.Thread(target=self.read_browser, args=(listener,), daemon=True)
        reader.start()
        try:
            self.send_native({"type": "hello", "host": VERSION, "protocol": VERSION})
            while not self.stopped.is_set():
                try:
                    connection, _ = listener.accept()
                except TimeoutError:
                    continue
                except OSError:
                    if self.stopped.is_set():
                        break
                    raise
                threading.Thread(target=self.serve_client, args=(connection,), daemon=True).start()
        finally:
            self.stopped.set()
            listener.close()
            path.unlink(missing_ok=True)
            lock.close()


def run() -> None:
    if f"chrome-extension://{EXTENSION_ID}/" not in sys.argv[1:]:
        raise CopilotError("This command must be started by the Copilot Chrome extension.")
    NativeHost(sys.stdin.buffer, sys.stdout.buffer).run()
