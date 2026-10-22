"""Microsoft 365 Copilot tool."""

from __future__ import annotations

from ..commands.copilot import resolve_copilot_file
from .clients import copilot_client, files_client
from .shaping import JsonDict


def copilot_ask(
    prompt: str,
    files: list[str] | None = None,
    site: str | None = None,
    library: str | None = None,
    conversation_id: str | None = None,
) -> JsonDict:
    """Ask Microsoft 365 Copilot a question. `files` are SharePoint/OneDrive URLs (such as webUrl from files_search), drive item IDs or paths (OneDrive unless `site` is given); Copilot reads them as context. For a follow-up, pass conversationId from an earlier result; Copilot keeps earlier files. NOT read-only: Copilot can use its own tools to create drafts and events, change Teams and SharePoint data, and more. Ask only for information; use the guarded tools here to send or change data. The result lists `toolCalls` and any `pendingConfirmations`, which are never confirmed. An answer can take a few minutes."""
    resolved = []
    if files:
        with files_client() as gc:
            resolved = [resolve_copilot_file(gc, ref, site=site, library=library) for ref in files]
    with copilot_client() as client:
        answer = client.ask(prompt, files=resolved, conversation_id=conversation_id)
    result = answer.to_dict()
    result["files"] = [{"name": f.name, "id": f.item_id, "driveId": f.drive_id, "webUrl": f.web_url} for f in resolved]
    return result


TOOLS = [copilot_ask]
