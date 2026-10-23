"""Resolve an Office file and collect a live editor request."""
from __future__ import annotations

import time
from collections.abc import Callable
from datetime import UTC, datetime
from urllib.parse import parse_qs, urlsplit

from ..commands.copilot import resolve_copilot_file
from ..errors import CopilotError
from ..outlook_client import OutlookClient
from .client import BridgeClient
from .protocol import timeout_seconds

APPS = {"xlsx": "excel", "xlsm": "excel", "pptx": "powerpoint", "docx": "word"}


def detect_app(name: str, override: str | None = None) -> str:
    app = override or APPS.get(name.rsplit(".", 1)[-1].lower())
    if app not in {"excel", "powerpoint", "word"}:
        raise CopilotError("Could not detect an app. Use an .xlsx, .xlsm, .pptx or .docx file, or --app.")
    return app


def _valid_file_url(value: str) -> None:
    url = urlsplit(value)
    if url.scheme != "https" or not url.hostname or not url.hostname.endswith(".sharepoint.com") or url.username:
        raise CopilotError("The file must have an HTTPS SharePoint or OneDrive web URL.")


def resolve_app_file(gc: OutlookClient, ref: str, *, site: str | None = None,
                     library: str | None = None, app: str | None = None) -> dict:
    resolved = resolve_copilot_file(gc, ref, site=site, library=library)
    chosen_app = detect_app(resolved.name, app)
    _valid_file_url(resolved.web_url)
    item = gc.drive_item_with_sharepoint_ids(resolved.item_id, resolved.drive_id)
    ids = item.get("sharepointIds") or {}
    unique_id = ids.get("listItemUniqueId")
    if not unique_id:
        query = parse_qs(urlsplit(resolved.web_url).query)
        unique_id = next((v[0] for k, v in query.items() if k.lower() == "sourcedoc"), None)
    if not unique_id:
        raise CopilotError("Graph did not return listItemUniqueId. The bridge cannot safely identify the open tab.")
    own_drive = gc._request("GET", "/me/drive", params={"$select": "id"}).json()
    return {
        "name": resolved.name, "id": resolved.item_id, "driveId": resolved.drive_id,
        "webUrl": resolved.web_url, "uniqueId": unique_id, "app": chosen_app,
        "shared": resolved.drive_id != own_drive.get("id"),
    }


def _modified_at(gc: OutlookClient, file: dict) -> datetime | None:
    item = gc._request("GET", f"/drives/{file['driveId']}/items/{file['id']}",
                       params={"$select": "lastModifiedDateTime"}).json()
    value = item.get("lastModifiedDateTime")
    if not value:
        return None
    try:
        return datetime.fromisoformat(value.replace("Z", "+00:00"))
    except (ValueError, AttributeError):
        raise CopilotError("Graph returned an invalid lastModifiedDateTime.") from None


def wait_for_save(gc: OutlookClient, file: dict, started: datetime, deadline: float) -> bool:
    while True:
        modified = _modified_at(gc, file)
        if modified and modified > started:
            return True
        if time.monotonic() >= deadline:
            return False
        time.sleep(min(5, max(0, deadline - time.monotonic())))


def _action_evidence(event: dict) -> dict:
    if event.get("app") == "excel" and "writes" in event:
        event = dict(event)
        # Generated scripts contain helper functions which may never run.
        # Source text alone cannot prove that a workbook was changed.
        event["mayWrite"] = event.pop("writes")
    return event


def _normalise_result(result: dict, file: dict) -> None:
    result["actions"] = [_action_evidence(action) for action in result.get("actions", [])]
    if file["app"] in {"excel", "word"}:
        result["applied"] = None
        result["appliedEvidence"] = "Not verified. Function execution and source-text patterns do not prove document changes."


def ask_app(gc: OutlookClient, prompt: str, file: dict, *, new_chat: bool = False,
            open_tab: bool = False, wait_saved: bool = False, timeout: float = 300,
            on_event: Callable[[dict], None] | None = None, bridge: BridgeClient | None = None) -> dict:
    timeout = timeout_seconds(timeout)
    if not isinstance(prompt, str) or not prompt.strip() or len(prompt) > 40000:
        raise CopilotError("A prompt must contain 1 to 40000 characters.")
    started = datetime.now(UTC)
    deadline = time.monotonic() + timeout
    result = None
    client = bridge or BridgeClient()
    for event in client.ask(prompt, file=file, app=file["app"], new_chat=new_chat, open_tab=open_tab, timeout=timeout):
        if event.get("type") == "result":
            result = event["result"]
        elif on_event:
            on_event(_action_evidence(event))
    if result is None:
        raise CopilotError("The bridge did not return a result.")
    _normalise_result(result, file)
    result["file"] = file
    result["saved"] = None
    if wait_saved and result.get("status") == "completed":
        result["saved"] = wait_for_save(gc, file, started, deadline)
        result["saveEvidence"] = "Graph lastModifiedDateTime changed after the request began. A co-author or another tab can also cause this change."
    return result
