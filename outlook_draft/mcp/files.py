"""OneDrive and SharePoint file tools."""

from __future__ import annotations

from ..commands.files import (
    _document_library_drives,
    _download_destination,
    _drive_summary,
    _find_sp_item_by_path,
    _item_summary,
    _match_library,
    _resolve_site_drives,
    _site_summary,
    _write_download,
)
from .clients import clamp, files_client
from .shaping import JsonDict, require_absolute_dir


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
    limit = clamp(count, high=100)
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
    """Download a file to an absolute local path (file or directory). `item` is a path, or for SharePoint also a drive item ID. OneDrive is used when `site` is omitted. Writes a file."""
    require_absolute_dir(dest, "dest")
    with files_client() as gc:
        if site:
            drive, found = _find_sp_item_by_path(gc, site, item, library)
            _reject_folder(found)
            data = gc.sp_download(drive["id"], found["id"])
        else:
            found = gc.od_item_by_path(item)
            _reject_folder(found)
            data = gc.od_download(found["id"])
    output = _download_destination(dest, found["name"])
    _write_download(output, data, overwrite=overwrite)
    return {"name": found["name"], "path": str(output), "size": len(data)}


def _reject_folder(item: JsonDict) -> None:
    if "folder" in item:
        raise ValueError("Download supports files, not folders.")


TOOLS = [files_sites, files_libraries, files_list, files_search, files_download]
