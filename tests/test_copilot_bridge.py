from __future__ import annotations

import io
import json
import os
import socket
import struct
import threading
from contextlib import contextmanager

import pytest

from outlook_draft.copilot_bridge import app, install, native_host, protocol
from outlook_draft.copilot_bridge.client import ASK_GRACE_SECONDS, BridgeClient, BridgeTimeout
from outlook_draft.errors import CopilotError


def test_native_message_round_trip_and_partial_reads():
    class Partial(io.BytesIO):
        def read(self, n=-1):
            return super().read(min(n, 2))

    data = io.BytesIO()
    protocol.write_native(data, {"type": "hello", "text": "café"})
    assert protocol.read_native(Partial(data.getvalue())) == {"v": 1, "type": "hello", "text": "café"}


@pytest.mark.parametrize("data", [
    struct.pack("=I", 0), struct.pack("=I", protocol.MAX_MESSAGE + 1),
    struct.pack("=I", 5) + b"x", b"\x01\x00", struct.pack("=I", 1),
])
def test_invalid_native_frames(data):
    with pytest.raises(CopilotError):
        protocol.read_native(io.BytesIO(data))


def test_native_eof():
    with pytest.raises(EOFError):
        protocol.read_native(io.BytesIO())


@pytest.mark.parametrize("data", [b"[]", b'{}', b'{"v":2}', b'"text"', b'\xff', b'{'])
def test_invalid_protocol(data):
    with pytest.raises(CopilotError):
        protocol.decode_message(data)


@pytest.mark.parametrize("value", [0, -1, True, float("nan"), float("inf"), 3601, "300"])
def test_invalid_timeout(value):
    with pytest.raises(CopilotError):
        protocol.timeout_seconds(value)


def test_socket_permissions_lock_and_same_uid(tmp_path):
    path = tmp_path / "bridge" / "host.sock"
    listener, lock = native_host.acquire_socket(path)
    try:
        assert path.stat().st_mode & 0o777 == 0o600
        assert path.parent.stat().st_mode & 0o777 == 0o700
        with pytest.raises(BlockingIOError):
            native_host.acquire_socket(path, lock_timeout=0)
        with socket.socket(socket.AF_UNIX, socket.SOCK_STREAM) as client:
            client.connect(str(path))
            peer, _ = listener.accept()
            with peer:
                assert native_host.same_uid(peer)
    finally:
        listener.close()
        lock.close()


def test_replacement_host_waits_for_previous_lock(tmp_path):
    path = tmp_path / "bridge" / "host.sock"
    old_listener, old_lock = native_host.acquire_socket(path)
    released = threading.Event()

    def release():
        released.wait(0.1)
        old_listener.close()
        old_lock.close()

    thread = threading.Thread(target=release)
    thread.start()
    listener, lock = native_host.acquire_socket(path, lock_timeout=2)
    try:
        assert old_lock.closed
        assert path.is_socket()
    finally:
        listener.close()
        lock.close()
        thread.join(2)


def test_different_uid_is_rejected():
    class Peer:
        def getsockopt(self, *args):
            return struct.pack("3i", 1, os.getuid() + 1, 1)
    assert not native_host.same_uid(Peer())


def test_unsafe_socket_and_directory_refused(tmp_path):
    target = tmp_path / "host.sock"
    target.write_text("keep")
    with pytest.raises(CopilotError, match="unsafe"):
        native_host.acquire_socket(target)
    assert target.read_text() == "keep"
    link = tmp_path / "link"
    link.symlink_to(tmp_path, target_is_directory=True)
    with pytest.raises(CopilotError):
        native_host.secure_directory(link)


@contextmanager
def fake_bridge(tmp_path, messages, hold=None):
    path = tmp_path / "host.sock"
    server = socket.socket(socket.AF_UNIX, socket.SOCK_STREAM)
    server.bind(str(path))
    path.chmod(0o600)
    server.listen(1)
    requests = []

    def serve():
        connection, _ = server.accept()
        with connection, connection.makefile("rb") as stream:
            requests.append(protocol.decode_message(stream.readline()))
            for message in messages:
                connection.sendall(protocol.encode_message(message) + b"\n")
            if hold:
                hold.wait(5)  # A stuck bridge: the connection stays open with no result.
        server.close()

    thread = threading.Thread(target=serve)
    thread.start()
    try:
        yield BridgeClient(path), requests
    finally:
        thread.join(timeout=5)
        server.close()
        assert not thread.is_alive()


HELLO = {"type": "hello", "host": 1, "protocol": 1, "extension": "0.1.0"}


def test_client_status(tmp_path):
    with fake_bridge(tmp_path, [HELLO, {"done": True, "result": {"tabs": [{"app": "excel"}]}}]) as (client, requests):
        assert client.status()["tabs"] == [{"app": "excel"}]
        assert requests[0]["type"] == "status"


def test_client_ask_events(tmp_path):
    with fake_bridge(tmp_path, [HELLO, {"event": {"type": "text", "text": "Hi"}},
                               {"done": True, "result": {"status": "completed"}}]) as (client, requests):
        events = list(client.ask("test", file={"id": "i"}, app="excel"))
        assert events == [{"type": "text", "text": "Hi"}, {"type": "result", "result": {"status": "completed"}}]
        assert requests[0]["prompt"] == "test"


def test_version_mismatch(tmp_path):
    with (
        fake_bridge(tmp_path, [{**HELLO, "protocol": 2}]) as (client, _),
        pytest.raises(CopilotError, match="protocol"),
    ):
        client.status()


def test_closed_chrome_and_disconnect(tmp_path):
    with pytest.raises(CopilotError, match="not running"):
        BridgeClient(tmp_path / "missing").status()
    with fake_bridge(tmp_path, [HELLO]) as (client, _), pytest.raises(CopilotError, match="Do not retry"):
        client.status()


def test_install(tmp_path):
    result = install.install(config_home=tmp_path / "config", data_dir=tmp_path / "data")
    from pathlib import Path
    manifest = json.loads(Path(result["manifest"]).read_text())
    assert manifest["allowed_origins"] == [f"chrome-extension://{protocol.EXTENSION_ID}/"]
    assert Path(result["launcher"]).stat().st_mode & 0o777 == 0o700
    ext_manifest = json.loads((Path(result["extension"]) / "manifest.json").read_text())
    assert "<all_urls>" not in json.dumps(ext_manifest)
    assert "localhost" not in json.dumps(ext_manifest)
    import base64
    import hashlib
    public = base64.b64decode(ext_manifest["key"])
    extension_id = "".join(chr(ord("a") + int(c, 16)) for c in hashlib.sha256(public).hexdigest()[:32])
    assert extension_id == protocol.EXTENSION_ID


@pytest.mark.parametrize(("name", "expected"), [("Book.XLSX", "excel"), ("a.xlsm", "excel"), ("deck.pptx", "powerpoint")])
def test_app_detection(name, expected):
    assert app.detect_app(name) == expected


def test_unsupported_apps():
    assert app.detect_app("a.docx") == "word"
    for name in ["a.pdf", "a"]:
        with pytest.raises(CopilotError):
            app.detect_app(name)
    assert app.detect_app("file", "excel") == "excel"


@pytest.mark.parametrize("url", ["http://tenant.sharepoint.com/f", "https://evil.test/f", "https://tenant.sharepoint.com.evil.test/f", "https://user@tenant.sharepoint.com/f"])
def test_invalid_file_urls(url):
    with pytest.raises(CopilotError):
        app._valid_file_url(url)


def test_ask_result_and_save_evidence(monkeypatch):
    class Bridge:
        def ask(self, *args, **kwargs):
            yield {"type": "text", "text": "Hi"}
            yield {"type": "result", "result": {"status": "completed", "applied": True, "text": "Hi"}}
    monkeypatch.setattr(app, "wait_for_save", lambda *args: True)
    events = []
    result = app.ask_app(None, "prompt", {"app": "excel"}, bridge=Bridge(), wait_saved=True, on_event=events.append)
    assert result["saved"] is True
    assert "co-author" in result["saveEvidence"]
    assert events == [{"type": "text", "text": "Hi"}]


def test_native_host_end_to_end_with_browser_pipes(tmp_path, monkeypatch):
    from_browser_read, from_browser_write = os.pipe()
    to_browser_read, to_browser_write = os.pipe()
    stdin = os.fdopen(from_browser_read, "rb", buffering=0)
    stdout = os.fdopen(to_browser_write, "wb", buffering=0)
    browser_in = os.fdopen(to_browser_read, "rb", buffering=0)
    browser_out = os.fdopen(from_browser_write, "wb", buffering=0)
    ready = threading.Event()
    acquire = native_host.acquire_socket

    def bind(path):
        result = acquire(path)
        ready.set()
        return result

    monkeypatch.setattr(native_host, "acquire_socket", bind)
    host = native_host.NativeHost(stdin, stdout)
    path = tmp_path / "bridge" / "host.sock"
    host_thread = threading.Thread(target=host.run, args=(path,), daemon=True)
    host_thread.start()
    assert ready.wait(5)

    def browser():
        assert protocol.read_native(browser_in)["type"] == "hello"
        protocol.write_native(browser_out, {"type": "hello", "protocol": 1, "extension": "test"})
        request = protocol.read_native(browser_in)
        protocol.write_native(browser_out, {"id": request["id"], "done": True, "result": {"tabs": [{"app": "excel"}]}})

    browser_thread = threading.Thread(target=browser, daemon=True)
    browser_thread.start()
    try:
        assert BridgeClient(path).status()["tabs"] == [{"app": "excel"}]
        browser_thread.join(5)
        assert not browser_thread.is_alive()
    finally:
        browser_out.close()
        host_thread.join(5)
        stdin.close()
        stdout.close()
        browser_in.close()
    assert not host_thread.is_alive()
    assert not path.exists()


def test_inspect_is_read_only_and_validates_tab_id(tmp_path):
    with fake_bridge(tmp_path, [HELLO, {"done": True, "result": {"frames": []}}]) as (client, requests):
        assert client.inspect(7) == {"frames": []}
        assert requests[0] == {"v": 1, "type": "inspect", "tabId": 7}
    assert native_host.NativeHost.validate_request({"type": "inspect", "tabId": 7}) == 30
    for bad in [True, -1, "7", None]:
        with pytest.raises(CopilotError):
            native_host.NativeHost.validate_request({"type": "inspect", "tabId": bad})
        with pytest.raises(CopilotError):
            BridgeClient(tmp_path / "missing").inspect(bad)


def test_host_request_validation():
    valid = {"v": 1, "type": "ask", "prompt": "Hello", "app": "excel", "file": {}}
    # The host waits longer than the client, so the client deadline ends a stuck request first.
    assert native_host.NativeHost.validate_request(valid) == 360
    assert native_host.NativeHost.validate_request(valid) > 300 + ASK_GRACE_SECONDS
    assert native_host.NativeHost.validate_request({"type": "reset", "tabId": 7}) == 30
    for extra in [{"prompt": ""}, {"app": "pdf"}, {"file": "x"}, {"timeout": float("nan")}]:
        with pytest.raises(CopilotError):
            native_host.NativeHost.validate_request({**valid, **extra})


def test_client_deadline_ends_a_silent_bridge(tmp_path):
    import time
    hold = threading.Event()
    with fake_bridge(tmp_path, [HELLO, {"event": {"type": "phase", "phase": "wait_answer"}}], hold) as (client, _):
        started = time.monotonic()
        events = []
        try:
            with pytest.raises(BridgeTimeout, match="time limit"):
                for event in client.request({"type": "status"}, timeout=5, deadline=time.monotonic() + 0.3):
                    events.append(event)
        finally:
            hold.set()
        assert time.monotonic() - started < 3
        assert events == [{"type": "phase", "phase": "wait_answer"}]


def test_client_reset(tmp_path):
    result = {"tabId": 7, "cleared": {"request": False, "uncertain": True}}
    with fake_bridge(tmp_path, [HELLO, {"done": True, "result": result}]) as (client, requests):
        assert client.reset(7) == result
        assert requests[0] == {"v": 1, "type": "reset", "tabId": 7}
    for bad in [True, -1, "7"]:
        with pytest.raises(CopilotError):
            BridgeClient(tmp_path / "missing").reset(bad)


class _InterruptedBridge:
    def __init__(self, events, error):
        self.events, self.error = events, error

    def ask(self, *args, **kwargs):
        yield from self.events
        raise self.error


def test_timeout_after_send_reports_phase_and_unknown_edits():
    bridge = _InterruptedBridge([{"type": "phase", "phase": "send"}, {"type": "sent"},
                                 {"type": "phase", "phase": "wait_answer"}], BridgeTimeout("no result"))
    seen = []
    result = app.ask_app(None, "prompt", {"app": "powerpoint"}, bridge=bridge, on_event=seen.append)
    assert result["status"] == "timeout"
    assert result["phase"] == "wait_answer"
    assert result["sent"] is True
    assert result["applied"] is None
    assert [e["type"] for e in seen] == ["phase", "sent", "phase"]


def test_error_before_send_reports_phase_and_no_edits():
    bridge = _InterruptedBridge([{"type": "phase", "phase": "wait_pane"}], CopilotError("disconnected"))
    result = app.ask_app(None, "prompt", {"app": "powerpoint"}, bridge=bridge)
    assert (result["status"], result["phase"], result["sent"], result["applied"]) == ("error", "wait_pane", False, False)


def test_error_before_any_phase_is_raised():
    with pytest.raises(CopilotError, match="Open the file"):
        app.ask_app(None, "prompt", {"app": "excel"}, bridge=_InterruptedBridge([], CopilotError("Open the file in Chrome first")))
