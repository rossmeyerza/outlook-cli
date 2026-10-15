"""CLI registration for completion."""
from __future__ import annotations

import argparse

from .. import completion
def cmd_completion(args: argparse.Namespace) -> None:
    completion.print_completion(args.shell, args._parser)


def register(sub) -> None:
    p_completion = sub.add_parser(
        "completion",
        help="Print a shell completion script",
    )
    p_completion.add_argument("shell", choices=list(completion.SHELLS))
    p_completion.set_defaults(func=cmd_completion, _parser=sub._parser)
