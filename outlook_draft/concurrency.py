"""Bounded, ordered work for independent Graph reads."""
from __future__ import annotations

from collections.abc import Callable, Iterable, Iterator
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from itertools import islice
from typing import TypeVar

Input = TypeVar("Input")
Output = TypeVar("Output")
WORKERS = 8


@contextmanager
def ordered_map(function: Callable[[Input], Output], items: Iterable[Input]) -> Iterator[Iterator[Output]]:
    """Keep input order and limit speculative reads when a search stops early."""
    with ThreadPoolExecutor(max_workers=WORKERS) as executor:
        def results() -> Iterator[Output]:
            source = iter(items)
            while batch := list(islice(source, WORKERS)):
                yield from executor.map(function, batch)

        yield results()
