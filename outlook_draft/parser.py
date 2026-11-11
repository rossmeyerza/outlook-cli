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

def add_email_format_args(parser: argparse.ArgumentParser) -> None:
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--format", dest="body_format", choices=["markdown", "html"], default="markdown",
                       help="Body format: Markdown (default) or HTML")
    group.add_argument("--html", dest="body_format", action="store_const", const="html", help=argparse.SUPPRESS)


def email_is_html(args: argparse.Namespace) -> bool:
    """Accept old internal namespaces without exposing a third composition format."""
    from .email_markdown import is_html_format

    fallback = "html" if getattr(args, "html", False) else "markdown"
    return is_html_format(getattr(args, "body_format", fallback))


def _add_yes_arg(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--yes", action="store_true", help="Confirm sending. Without it, nothing is sent.")

