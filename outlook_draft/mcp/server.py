"""FastMCP app for outlook-cli.

stdout carries the MCP protocol. Nothing in this package may print to it.
"""

from __future__ import annotations

import logging
import sys

from mcp.server.fastmcp import FastMCP

from . import auth, calendar, contacts, drafts, files, mail, tasks, teams

INSTRUCTIONS = (
    "Outlook mail, drafts, calendar, contacts, tasks, Teams (read-only) and OneDrive/SharePoint "
    "files for Ross's Microsoft 365 account. Sending is not available here; create drafts.\n\n"
    "All IDs are full Outlook or Graph IDs returned by the list and search tools. "
    "If a tool reports an expired token, ask the user to run 'outlook-cli auth' in a terminal."
)

MODULES = (mail, drafts, calendar, contacts, tasks, teams, files, auth)


def build_server() -> FastMCP:
    """Create the server and register every phase 1 tool."""
    server = FastMCP("outlook", instructions=INSTRUCTIONS)
    for module in MODULES:
        for tool in module.TOOLS:
            server.tool()(tool)
    return server


def run() -> None:
    """Serve over stdio. Logs go to stderr only."""
    logging.basicConfig(stream=sys.stderr, level=logging.WARNING, format="%(levelname)s %(name)s: %(message)s")
    build_server().run(transport="stdio")
