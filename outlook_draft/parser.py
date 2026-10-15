"""Shared argparse options."""
from __future__ import annotations

import argparse


def add_output_args(parser: argparse.ArgumentParser) -> None:
    parser.add_argument(
        "--json",
        dest="json",
        action="store_true",
        default=argparse.SUPPRESS,
        help="Output machine-readable JSON",
    )
    parser.add_argument(
        "--table",
        dest="json",
        action="store_false",
        default=argparse.SUPPRESS,
        help="Output human-readable tables (default)",
    )

def _add_yes_arg(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--yes", action="store_true", help="Confirm sending. Without it, nothing is sent.")

