"""Microsoft 365 Copilot tool."""

from __future__ import annotations

from ..commands.copilot import resolve_copilot_file
from ..copilot_bridge.app import ask_app, resolve_app_file
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


def copilot_app(
    prompt: str,
    file: str,
    app: str | None = None,
    site: str | None = None,
    library: str | None = None,
    new_chat: bool = False,
    open_tab: bool = False,
    wait_saved: bool = False,
) -> JsonDict:
    """EXPERIMENTAL: Send a prompt to Copilot in a live Excel, PowerPoint or Word web editor in the user's Chrome. This can edit the live document. Co-authors see edits at once. SharePoint version history can restore earlier versions. Use only when the user asks for this operation. The Chrome extension and native host must be installed. Never approves actions or changes the Allow editing setting. Approval detection still needs live verification. Word uses the observed OfficeAgent chat protocol, but automated Word reads and edits still need live acceptance. A timeout or disconnect does not undo edits: never retry automatically. `saved` uses Graph modification time, not proof that these specific edits were saved."""
    with files_client() as gc:
        resolved = resolve_app_file(gc, file, site=site, library=library, app=app)
        result = ask_app(gc, prompt, resolved, new_chat=new_chat, open_tab=open_tab, wait_saved=wait_saved)
    if resolved["shared"]:
        result["warning"] = "This file is outside the user's own OneDrive. Other people can see edits at once."
    return result


TOOLS = [copilot_ask, copilot_app]
