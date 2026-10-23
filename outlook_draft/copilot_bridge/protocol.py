"""Local bridge protocol. No Office tokens or raw frames belong here."""
from __future__ import annotations

import json
import math
import struct
from pathlib import Path
from typing import BinaryIO

from .. import config
from ..errors import CopilotError

VERSION = 1
MAX_MESSAGE = 1024 * 1024
HOST_NAME = "com.outlook_cli.copilot"
EXTENSION_ID = "lijjlmfmgilgjikalapbblbadhabmnae"


def socket_path() -> Path:
    # This bridge is browser-wide, not tied to the CLI's active account profile.
    return config.DATA_DIR / "copilot-bridge.sock"


def timeout_seconds(value: object) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise CopilotError("Timeout must be a number of seconds.")
    if not math.isfinite(value) or not 10 <= value <= 3600:
        raise CopilotError("Timeout must be between 10 and 3600 seconds.")
    return float(value)


def decode_message(data: bytes) -> dict:
    if len(data) > MAX_MESSAGE:
        raise CopilotError("Bridge message is too large.")
    try:
        value = json.loads(data)
    except (ValueError, UnicodeError) as exc:
        raise CopilotError("Invalid bridge JSON.") from exc
    if not isinstance(value, dict) or value.get("v") != VERSION:
        raise CopilotError("Unsupported Copilot bridge protocol. Update the host and extension.")
    return value


def encode_message(value: dict) -> bytes:
    data = json.dumps({**value, "v": VERSION}, ensure_ascii=False, allow_nan=False).encode()
    if len(data) > MAX_MESSAGE:
        raise CopilotError("Bridge message is too large.")
    return data


def read_exact(stream: BinaryIO, size: int) -> bytes:
    chunks = bytearray()
    while len(chunks) < size:
        part = stream.read(size - len(chunks))
        if not part:
            if not chunks:
                raise EOFError
            raise CopilotError("Truncated native message.")
        chunks.extend(part)
    return bytes(chunks)


def read_native(stream: BinaryIO) -> dict:
    length = struct.unpack("=I", read_exact(stream, 4))[0]
    if not 0 < length <= MAX_MESSAGE:
        raise CopilotError("Invalid native message length.")
    try:
        return decode_message(read_exact(stream, length))
    except EOFError as exc:
        raise CopilotError("Truncated native message.") from exc


def write_native(stream: BinaryIO, value: dict) -> None:
    data = encode_message(value)
    stream.write(struct.pack("=I", len(data)) + data)
    stream.flush()
