"""OneDrive and SharePoint file tools."""

from __future__ import annotations

from pathlib import Path

from ..commands.files import (
    _document_library_drives,
    _download_destination,
    _download_share_url,
    _drive_summary,
    _find_od_item,
    _find_sp_item_by_path,
    _item_summary,
    _match_library,
    _resolve_site_drives,
    _site_summary,
    _write_download,
)
from ..commands.mail import _safe_filename
from ..links import looks_like_share_url
from .clients import clamp, files_client
from .shaping import JsonDict, require_absolute_dir
from .validation import require_name


def files_sites() -> list[JsonDict]:
    """List SharePoint sites (Microsoft 365 groups) you are a member of. Read-only."""
    with files_client() as gc:
        groups = gc.sp_list_sites()
    return [_site_summary(g) for g in sorted(groups, key=lambda g: g["displayName"].lower())]


def files_libraries(site: str) -> list[JsonDict]:
    """List document libraries in a SharePoint site (partial, case-insensitive site name). Read-only."""
    with files_client() as gc:
        _, _, drives = _resolve_site_drives(gc, site)
    return [_drive_summary(d) for d in _document_library_drives(drives)]


def files_list(path: str = "", site: str | None = None, library: str | None = None) -> JsonDict:
    """List a folder in OneDrive, or in a SharePoint site when `site` is given. With a site and no path or library, lists the site's libraries. Read-only."""
    with files_client() as gc:
        if site and not path and not library:
            _, _, drives = _resolve_site_drives(gc, site)
            libraries = _document_library_drives(drives)
            return {
                "location": f"SharePoint: {site}",
                "kind": "libraries",
                "items": [_drive_summary(d) for d in libraries],
            }
        if site:
            drive, item = _find_sp_item_by_path(gc, site, path, library)
            children = gc.sp_list_children(drive["id"], item["id"])
            library_name = drive.get("name", "")
            location = f"SharePoint: {site} / {library_name}"
        else:
            item = gc.od_item_by_path(path)
            children = gc.od_list_children(item["id"])
            library_name, location = "", "OneDrive"
    return {
        "location": location,
        "kind": "items",
        "items": [_item_summary(i, library=library_name) for i in children],
    }


def files_search(
    query: str, site: str | None = None, library: str | None = None, count: int = 20
) -> list[JsonDict]:
    """Search document libraries of a SharePoint site (site is required). Read-only. Returns items with webUrl."""
    if not site:
        raise ValueError("files_search requires 'site' (a SharePoint site name).")
    limit = clamp(count)
    results: list[JsonDict] = []
    with files_client() as gc:
        _, _, drives = _resolve_site_drives(gc, site)
        candidates = _document_library_drives(drives)
        if library:
            candidates = [_match_library(candidates, library)]
        per_drive = limit if len(candidates) == 1 else max(limit, 25)
        for drive in candidates:
            for item in gc.sp_search(drive["id"], query, per_drive):
                results.append(_item_summary(item, library=drive.get("name", "")))
    return results[:limit]


def files_download(
    item: str,
    dest: str,
    site: str | None = None,
    library: str | None = None,
    overwrite: bool = False,
) -> JsonDict:
    """Download a file to an absolute local path or directory. Accepts SharePoint/OneDrive URLs, paths and item IDs. A share URL identifies its own site and library. Writes a local file without remote changes."""
    require_absolute_dir(dest, "dest")
    with files_client() as gc:
        if looks_like_share_url(item):
            found, data = _download_share_url(gc, item)
        elif site:
            drive, found = _find_sp_item_by_path(gc, site, item, library)
            _reject_folder(found)
            data = gc.sp_download(drive["id"], found["id"])
        else:
            found = _find_od_item(gc, item)
            _reject_folder(found)
            data = gc.od_download(found["id"])
    output = _download_destination(dest, _safe_filename(found["name"]))
    _write_download(output, data, overwrite=overwrite)
    return {"name": found["name"], "path": str(output), "size": len(data)}


def _reject_folder(item: JsonDict) -> None:
    if "folder" in item:
        raise ValueError("Download supports files, not folders.")


def _file_context(client, path: str, site: str | None, library: str | None, *, folder: bool = False) -> JsonDict:
    if library and not site:
        raise ValueError("A library requires an explicit SharePoint site.")
    if any(part in {".", ".."} for part in path.split("/")) or "\\" in path:
        raise ValueError("Use a path relative to the selected drive without dot segments or backslashes.")
    if site:
        drive, item = _find_sp_item_by_path(client, site, path, library)
        context = {"drive_id": drive["id"], "library": drive.get("name", ""), "site": site, "item": item}
    else:
        context = {"drive_id": None, "library": None, "site": None, "item": client.od_item_by_path(path)}
    if folder and "folder" not in context["item"]:
        raise ValueError("The destination must be an existing folder.")
    return context


def _file_call(client, context: JsonDict, action: str, *args, **kwargs):
    drive = context["drive_id"]
    if drive is not None:
        return getattr(client, "sp_" + action)(drive, *args, **kwargs)
    return getattr(client, "od_" + action)(*args, **kwargs)


def _file_plan(operation: str, context: JsonDict) -> JsonDict:
    item = context["item"]
    if operation in {"rename", "move"} and "folder" in item and not (item.get("parentReference") or {}).get("id"):
        raise ValueError("Cannot rename or move a drive or library root.")
    return {"operation": operation, "site": context["site"], "library": context["library"],
            "driveId": context["drive_id"], "target": _item_summary(context["item"]), "confirmed": False}


def _existing_child(client, context: JsonDict, name: str) -> JsonDict | None:
    children = _file_call(client, context, "list_children", context["item"]["id"])
    return next((item for item in children if item.get("name", "").casefold() == name.casefold()), None)


def files_upload(file: str, dest: str = "", site: str | None = None, library: str | None = None,
                 overwrite: bool = False, confirm: bool = False) -> JsonDict:
    """Preview a local file upload to an existing OneDrive or SharePoint folder. Set confirm=true only after approval of the file and exact destination. Replacement also requires overwrite=true. No folder or share link is created. Never retry uncertain writes automatically."""
    if type(overwrite) is not bool:
        raise ValueError("overwrite must be a boolean")
    source = Path(file).expanduser().absolute()
    if not source.is_file():
        raise ValueError("file must refer to an existing regular local file")
    require_name(source.name)
    with files_client() as client:
        context = _file_context(client, dest, site, library, folder=True)
        existing = _existing_child(client, context, source.name)
        result = {**_file_plan("upload", context), "source": str(source), "name": source.name,
                  "size": source.stat().st_size, "existing": _item_summary(existing) if existing else None,
                  "overwrite": overwrite, "canApply": not existing or (overwrite is True and "folder" not in existing), "uploaded": False}
        if confirm is True:
            if not result["canApply"]:
                raise ValueError("The destination exists. Approve overwrite=true for a file replacement; folders cannot be replaced.")
            uploaded = _file_call(client, context, "upload", context["item"]["id"], source.name, source, overwrite=overwrite)
            result.update(confirmed=True, uploaded=True, item=_item_summary(uploaded))
    return result


def files_mkdir(path: str, site: str | None = None, library: str | None = None, confirm: bool = False) -> JsonDict:
    """Preview one new folder in an existing parent. Set confirm=true only after approval of the exact site, library and path. Does not create missing parent folders. Never retry uncertain writes automatically."""
    parent, _, name = path.strip("/").rpartition("/")
    require_name(name)
    with files_client() as client:
        context = _file_context(client, parent, site, library, folder=True)
        if _existing_child(client, context, name):
            raise ValueError("The destination already exists.")
        result = {**_file_plan("mkdir", context), "name": name, "created": False}
        if confirm is True:
            item = _file_call(client, context, "mkdir", context["item"]["id"], name)
            result.update(confirmed=True, created=True, item=_item_summary(item))
    return result


def files_rename(path: str, name: str, site: str | None = None, library: str | None = None,
                 confirm: bool = False) -> JsonDict:
    """Preview a file or folder rename. The new name must not contain a path. Set confirm=true only after approval of the exact item and new name. Never retry uncertain writes automatically."""
    require_name(name)
    if not path.strip("/"):
        raise ValueError("Cannot rename the drive root.")
    with files_client() as client:
        context = _file_context(client, path, site, library)
        result = {**_file_plan("rename", context), "name": name, "renamed": False}
        if confirm is True:
            item = _file_call(client, context, "rename", context["item"]["id"], name)
            result.update(confirmed=True, renamed=True, item=_item_summary(item))
    return result


def files_move(path: str, dest: str, site: str | None = None, library: str | None = None,
               confirm: bool = False) -> JsonDict:
    """Preview a file or folder move within one drive. The destination folder must exist. Set confirm=true only after approval of the exact source and destination. Never retry uncertain writes automatically."""
    if not path.strip("/"):
        raise ValueError("Cannot move the drive root.")
    with files_client() as client:
        source = _file_context(client, path, site, library)
        destination = _file_context(client, dest, site, source["library"], folder=True)
        if source["drive_id"] != destination["drive_id"]:
            raise ValueError("Cross-drive moves are not supported.")
        existing = _existing_child(client, destination, source["item"]["name"])
        if existing and existing["id"] != source["item"]["id"]:
            raise ValueError("An item with this name already exists in the destination.")
        result = {**_file_plan("move", source), "destination": _item_summary(destination["item"]), "moved": False}
        if confirm is True:
            item = _file_call(client, source, "move", source["item"]["id"], destination["item"]["id"])
            result.update(confirmed=True, moved=True, item=_item_summary(item))
    return result


TOOLS = [files_sites, files_libraries, files_list, files_search, files_download, files_upload, files_mkdir, files_rename, files_move]
