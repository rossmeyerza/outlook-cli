"""Shared fakes for MCP tool tests."""

from __future__ import annotations

from contextlib import contextmanager
from typing import Any

import pytest


class FakeClient:
    """Records every call; return values are looked up by method name.

    A return value may be a callable, in which case it receives the call arguments.
    """

    def __init__(self, **returns: Any) -> None:
        self.returns = returns
        self.calls: list[tuple[str, tuple, dict]] = []

    def __getattr__(self, name: str):
        def method(*args: Any, **kwargs: Any) -> Any:
            self.calls.append((name, args, kwargs))
            result = self.returns.get(name)
            return result(*args, **kwargs) if callable(result) else result

        return method

    def names(self) -> list[str]:
        return [name for name, _, _ in self.calls]

    def call(self, name: str) -> tuple[tuple, dict]:
        for call_name, args, kwargs in self.calls:
            if call_name == name:
                return args, kwargs
        raise AssertionError(f"{name} was not called; calls: {self.names()}")


def use_client(monkeypatch: pytest.MonkeyPatch, module: Any, attr: str, fake: FakeClient) -> FakeClient:
    """Make `module.<attr>()` yield the fake client."""

    @contextmanager
    def factory():
        yield fake

    monkeypatch.setattr(module, attr, factory)
    return fake
