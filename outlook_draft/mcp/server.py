"""FastMCP app for outlook-cli.

stdout carries the MCP protocol. Nothing in this package may print to it.
"""

from __future__ import annotations

import logging
import os
import sys

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

from ..safety import command_allowed, current_policy
from . import auth, calendar, contacts, copilot, drafts, files, mail, mail_ops, mail_organise, sending, tasks, teams

INSTRUCTIONS = (
    "Outlook mail, drafts, calendar, contacts, tasks, Teams, OneDrive/SharePoint files "
    "and Microsoft 365 Copilot for Ross's Microsoft 365 account. Send tools preview unless confirm=true.\n\n"
    "Set confirm=true only when the user explicitly approves that tool's send or destructive change. "
    "Prefer drafts. Never automatically retry a send or destructive change after an uncertain result. "
    "Do not send Teams messages directly from the Teams gateway.\n\n"
    "All IDs are full Outlook or Graph IDs returned by the list and search tools. "
    "If a tool reports an expired token, ask the user to run 'outlook-cli auth' in a terminal."
)

MODULES = (mail, mail_ops, mail_organise, drafts, calendar, contacts, tasks, teams, files, copilot, auth, sending)

READ_ONLY_TOOLS = {
    "mail_search", "mail_unread", "mail_read", "mail_attachments", "mail_download_attachments",
    "mail_links", "mail_folders", "mail_query", "mail_conversation", "mail_categories", "mail_rules",
    "draft_list", "draft_get", "cal_agenda", "cal_get", "cal_availability", "cal_find_time",
    "contact_search", "task_list", "teams_list", "teams_search", "teams_messages",
    "teams_attachments", "teams_download_attachments", "files_sites", "files_libraries",
    "files_list", "files_search", "files_download", "auth_status",
    "mail_download_links", "cal_rooms", "teams_get", "teams_self",
}


def _tool_enabled(name: str) -> bool:
    active = current_policy()
    readonly, no_send = active.readonly, active.no_send
    if readonly and name not in READ_ONLY_TOOLS:
        return False
    if no_send and name in {"mail_send", "draft_send", "teams_send", "teams_new", "copilot_ask", "copilot_app"}:
        return False
    return command_allowed(name, os.environ.get("OUTLOOK_CLI_ENABLE_TOOLS") or None,
                           os.environ.get("OUTLOOK_CLI_DISABLE_TOOLS") or None)


def build_server() -> FastMCP:
    """Create the server and register all tools."""
    server = FastMCP("outlook", instructions=INSTRUCTIONS)
    for module in MODULES:
        for tool in module.TOOLS:
            if not _tool_enabled(tool.__name__):
                continue
            if tool.__name__ in {"mail_bulk", "mail_trash", "mail_categorise", "draft_update", "draft_forward",
                                  "cal_delete", "cal_cancel", "cal_update", "contact_create", "contact_update",
                                  "task_update", "task_delete", "files_upload", "files_mkdir", "files_rename", "files_move"} or (
                module is mail_organise and tool.__name__ not in READ_ONLY_TOOLS
            ) or module is sending or module is copilot:
                server.tool(annotations=ToolAnnotations(
                    readOnlyHint=False, destructiveHint=True, idempotentHint=False, openWorldHint=True,
                ))(tool)
            elif tool is mail.mail_create_folder:
                server.tool(annotations=ToolAnnotations(
                    readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=True,
                ))(tool)
            else:
                server.tool()(tool)
    return server


def run() -> None:
    """Serve over stdio. Logs go to stderr only."""
    logging.basicConfig(stream=sys.stderr, level=logging.WARNING, format="%(levelname)s %(name)s: %(message)s")
    build_server().run(transport="stdio")
