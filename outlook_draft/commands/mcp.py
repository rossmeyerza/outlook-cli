"""CLI registration for mcp."""
from __future__ import annotations

import argparse


def cmd_mcp(args: argparse.Namespace) -> None:
    """Run the MCP server on stdio. The active account comes from --account."""
    from ..mcp.server import run  # lazy: the mcp SDK is only needed for this command

    run()


def register(sub) -> None:
    p_mcp = sub.add_parser("mcp", help="Run the MCP server on stdio (no send tools)")
    p_mcp.set_defaults(func=cmd_mcp)
