"""FastMCP app for outlook-cli.

stdout carries the MCP protocol. Nothing in this package may print to it.
"""

from __future__ import annotations

import logging
import sys

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from . import auth, calendar, contacts, copilot, drafts, files, mail, sending, tasks, teams

INSTRUCTIONS = (
    "Outlook mail, drafts, calendar, contacts, tasks, Teams, OneDrive/SharePoint files "
    "and Microsoft 365 Copilot for Ross's Microsoft 365 account. Send tools preview unless confirm=true.\n\n"
    "Set confirm=true only when the user explicitly requests sending that message. "
    "Prefer drafts. Never automatically retry a send after an uncertain result. "
    "Do not send Teams messages directly from the Teams gateway.\n\n"
    "All IDs are full Outlook or Graph IDs returned by the list and search tools. "
    "If a tool reports an expired token, ask the user to run 'outlook-cli auth' in a terminal."
)

MODULES = (mail, drafts, calendar, contacts, tasks, teams, files, copilot, auth, sending)


def build_server() -> FastMCP:
    """Create the server and register all tools."""
    server = FastMCP("outlook", instructions=INSTRUCTIONS)
    for module in MODULES:
        for tool in module.TOOLS:
            if module is sending or module is copilot:
                server.tool(annotations=ToolAnnotations(
                    readOnlyHint=False, destructiveHint=True, idempotentHint=False, openWorldHint=True,
                ))(tool)
            else:
                server.tool()(tool)
    return server


def run() -> None:
    """Serve over stdio. Logs go to stderr only."""
    logging.basicConfig(stream=sys.stderr, level=logging.WARNING, format="%(levelname)s %(name)s: %(message)s")
    build_server().run(transport="stdio")
