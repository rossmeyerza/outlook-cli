"""OneDrive and SharePoint file operations via Microsoft Graph API."""

from __future__ import annotations

import logging
import math
from pathlib import Path
from typing import Any

import httpx
from rich.console import Console
from rich import box
from rich.table import Table

from .. import config
from ..errors import OutlookAPIError
from ..token_manager import TokenManager

log = logging.getLogger(__name__)
console = Console()

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
UPLOAD_CHUNK_SIZE = 10 * 320 * 1024  # 3.2 MB
ITEM_SELECT = "id,name,size,lastModifiedDateTime,file,folder,parentReference"


# ── Graph client (sync, reuses existing token infrastructure) ─────

class _GraphClient:
    """Thin sync Graph client using the captured Graph bearer token."""

    def __init__(self):
        self._tm = TokenManager(
            token_domain=config.GRAPH_TOKEN_DOMAIN,
            token_label="Microsoft Graph",
        )
        self._client: httpx.Client | None = None

    def _ensure(self) -> httpx.Client:
        if self._client is None or self._client.is_closed:
            self._client = httpx.Client(
                timeout=httpx.Timeout(60.0, connect=10.0),
                follow_redirects=True,
            )
        return self._client

    def _headers(self) -> dict[str, str]:
        return {"Authorization": f"Bearer {self._tm.token}"}

    def _request(
        self,
        method: str,
        url: str,
        *,
        params: dict | None = None,
        json_body: dict | None = None,
        data: bytes | None = None,
        extra_headers: dict | None = None,
        max_retries: int = 2,
    ) -> httpx.Response:
        import time
        client = self._ensure()
        headers = {**self._headers(), **(extra_headers or {})}

        for attempt in range(max_retries + 1):
            try:
                resp = client.request(
                    method, url, params=params, json=json_body,
                    content=data, headers=headers,
                )
                if resp.status_code in (200, 201, 202, 204):
                    return resp
                if resp.status_code == 401:
                    self._tm.force_reload()
                    headers = {**self._headers(), **(extra_headers or {})}
                    continue
                if resp.status_code == 429 or resp.status_code >= 500:
                    wait = float(resp.headers.get("Retry-After", str(2 ** attempt)))
                    time.sleep(wait)
                    continue
                raise OutlookAPIError(resp.status_code, resp.text[:500])
            except httpx.HTTPError as e:
                if attempt < max_retries:
                    time.sleep(2 ** attempt)
                    continue
                raise OutlookAPIError(0, str(e)) from e
        raise OutlookAPIError(0, "Max retries exceeded")

    def close(self):
        if self._client and not self._client.is_closed:
            self._client.close()

    # ── OneDrive ──────────────────────────────────────────────────────

    def od_root_id(self) -> str:
        resp = self._request("GET", f"{GRAPH_BASE}/me/drive/root",
                             params={"$select": "id"})
        return resp.json()["id"]

    def od_item_by_path(self, path: str) -> dict[str, Any]:
        """Resolve a path relative to OneDrive root."""
        path = path.strip("/")
        if not path:
            resp = self._request("GET", f"{GRAPH_BASE}/me/drive/root",
                                 params={"$select": ITEM_SELECT})
        else:
            resp = self._request("GET", f"{GRAPH_BASE}/me/drive/root:/{path}",
                                 params={"$select": ITEM_SELECT})
        return resp.json()

    def od_list_children(self, item_id: str) -> list[dict]:
        url = f"{GRAPH_BASE}/me/drive/items/{item_id}/children"
        params: dict | None = {"$select": ITEM_SELECT, "$top": "200"}
        items = []
        while url:
            resp = self._request("GET", url, params=params)
            data = resp.json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            params = None
        return items

    def od_upload(self, parent_id: str, name: str, local_path: Path) -> dict:
        data = local_path.read_bytes()
        size = len(data)
        if size < 4 * 1024 * 1024:
            url = f"{GRAPH_BASE}/me/drive/items/{parent_id}:/{name}:/content"
            resp = self._request("PUT", url, data=data,
                                 extra_headers={"Content-Type": "application/octet-stream"})
            return resp.json()
        return self._chunked_upload_od(parent_id, name, data)

    def od_download(self, item_id: str) -> bytes:
        url = f"{GRAPH_BASE}/me/drive/items/{item_id}/content"
        resp = self._request("GET", url)
        return resp.content

    def _chunked_upload_od(self, parent_id: str, name: str, data: bytes) -> dict:
        url = f"{GRAPH_BASE}/me/drive/items/{parent_id}:/{name}:/createUploadSession"
        sess = self._request("POST", url, json_body={
            "item": {"@microsoft.graph.conflictBehavior": "replace", "name": name}
        }).json()
        return self._upload_chunks(sess["uploadUrl"], data)

    def od_mkdir(self, parent_id: str, name: str) -> dict:
        url = f"{GRAPH_BASE}/me/drive/items/{parent_id}/children"
        resp = self._request("POST", url, json_body={
            "name": name, "folder": {},
            "@microsoft.graph.conflictBehavior": "fail"
        })
        return resp.json()

    def od_rename(self, item_id: str, new_name: str) -> dict:
        url = f"{GRAPH_BASE}/me/drive/items/{item_id}"
        resp = self._request("PATCH", url, json_body={"name": new_name})
        return resp.json()

    def od_move(self, item_id: str, new_parent_id: str, new_name: str | None = None) -> dict:
        url = f"{GRAPH_BASE}/me/drive/items/{item_id}"
        body: dict = {"parentReference": {"id": new_parent_id}}
        if new_name:
            body["name"] = new_name
        resp = self._request("PATCH", url, json_body=body)
        return resp.json()

    # ── SharePoint ────────────────────────────────────────────────────

    def sp_list_sites(self) -> list[dict]:
        """List SharePoint sites via group membership."""
        url = f"{GRAPH_BASE}/me/memberOf/microsoft.graph.group"
        params: dict | None = {
            "$select": "id,displayName,description,mail,groupTypes",
            "$top": "999",
        }
        groups = []
        while url:
            resp = self._request("GET", url, params=params)
            data = resp.json()
            groups.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            params = None
        # Only M365 Unified groups (have SharePoint sites)
        return [
            g for g in groups
            if "Unified" in g.get("groupTypes", [])
        ]

    def sp_get_site(self, group_id: str) -> dict:
        resp = self._request("GET", f"{GRAPH_BASE}/groups/{group_id}/sites/root",
                             params={"$select": "id,displayName,webUrl"})
        return resp.json()

    def sp_list_drives(self, site_id: str) -> list[dict]:
        resp = self._request("GET", f"{GRAPH_BASE}/sites/{site_id}/drives",
                             params={"$select": "id,name,driveType,webUrl"})
        return resp.json().get("value", [])

    def sp_item_by_path(self, drive_id: str, path: str) -> dict:
        path = path.strip("/")
        if not path:
            resp = self._request("GET", f"{GRAPH_BASE}/drives/{drive_id}/root",
                                 params={"$select": ITEM_SELECT})
        else:
            resp = self._request("GET", f"{GRAPH_BASE}/drives/{drive_id}/root:/{path}",
                                 params={"$select": ITEM_SELECT})
        return resp.json()

    def sp_list_children(self, drive_id: str, item_id: str) -> list[dict]:
        url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/children"
        params: dict | None = {"$select": ITEM_SELECT, "$top": "200"}
        items = []
        while url:
            resp = self._request("GET", url, params=params)
            data = resp.json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
            params = None
        return items

    def sp_upload(self, drive_id: str, parent_id: str, name: str, local_path: Path) -> dict:
        data = local_path.read_bytes()
        size = len(data)
        if size < 4 * 1024 * 1024:
            url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_id}:/{name}:/content"
            resp = self._request("PUT", url, data=data,
                                 extra_headers={"Content-Type": "application/octet-stream"})
            return resp.json()
        return self._chunked_upload_sp(drive_id, parent_id, name, data)

    def sp_download(self, drive_id: str, item_id: str) -> bytes:
        url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}/content"
        resp = self._request("GET", url)
        return resp.content

    def _chunked_upload_sp(self, drive_id: str, parent_id: str, name: str, data: bytes) -> dict:
        url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_id}:/{name}:/createUploadSession"
        sess = self._request("POST", url, json_body={
            "item": {"@microsoft.graph.conflictBehavior": "replace", "name": name}
        }).json()
        return self._upload_chunks(sess["uploadUrl"], data)

    def sp_mkdir(self, drive_id: str, parent_id: str, name: str) -> dict:
        url = f"{GRAPH_BASE}/drives/{drive_id}/items/{parent_id}/children"
        resp = self._request("POST", url, json_body={
            "name": name, "folder": {},
            "@microsoft.graph.conflictBehavior": "fail"
        })
        return resp.json()

    def sp_rename(self, drive_id: str, item_id: str, new_name: str) -> dict:
        url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
        resp = self._request("PATCH", url, json_body={"name": new_name})
        return resp.json()

    def sp_move(self, drive_id: str, item_id: str, new_parent_id: str,
                new_name: str | None = None) -> dict:
        url = f"{GRAPH_BASE}/drives/{drive_id}/items/{item_id}"
        body: dict = {"parentReference": {"driveId": drive_id, "id": new_parent_id}}
        if new_name:
            body["name"] = new_name
        resp = self._request("PATCH", url, json_body=body)
        return resp.json()

    # ── Shared helpers ────────────────────────────────────────────────

    def _upload_chunks(self, upload_url: str, data: bytes) -> dict:
        """Upload data in chunks to a pre-created upload session URL."""
        import time
        total = len(data)
        chunks = math.ceil(total / UPLOAD_CHUNK_SIZE)
        result = {}
        for i in range(chunks):
            offset = i * UPLOAD_CHUNK_SIZE
            chunk = data[offset: offset + UPLOAD_CHUNK_SIZE]
            end = offset + len(chunk) - 1
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {offset}-{end}/{total}",
            }
            resp = self._ensure().put(upload_url, content=chunk, headers=headers)
            if resp.status_code in (200, 201, 202):
                result = resp.json()
            else:
                raise OutlookAPIError(resp.status_code, resp.text[:300])
            console.print(f"[dim]  chunk {i+1}/{chunks} uploaded[/]")
        return result


# ── Site resolution helper ────────────────────────────────────────

def _resolve_site(gc: _GraphClient, site_hint: str) -> tuple[str, str, str]:
    """Resolve a site hint to (group_id, site_id, drive_id).

    site_hint is matched case-insensitively against group displayName.
    Returns the default document library drive.
    """
    groups = gc.sp_list_sites()
    hint_lower = site_hint.lower()
    matches = [g for g in groups if hint_lower in g["displayName"].lower()]

    if not matches:
        names = [g["displayName"] for g in groups]
        raise ValueError(
            f"No site matching {site_hint!r}. Available: {', '.join(names)}"
        )
    if len(matches) > 1:
        names = [g["displayName"] for g in matches]
        raise ValueError(
            f"Ambiguous site {site_hint!r}, matched: {', '.join(names)}. Be more specific."
        )

    group = matches[0]
    site = gc.sp_get_site(group["id"])
    site_id = site["id"]

    drives = gc.sp_list_drives(site_id)
    # Prefer the default document library (driveType == documentLibrary)
    doc_drives = [d for d in drives if d.get("driveType") == "documentLibrary"]
    drive = doc_drives[0] if doc_drives else drives[0]

    return group["id"], site_id, drive["id"]


# ── Formatting helpers ────────────────────────────────────────────

def _fmt_size(size: int | None) -> str:
    if size is None:
        return ""
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024:
            return f"{size:.0f} {unit}"
        size /= 1024
    return f"{size:.1f} TB"


def _print_items(items: list[dict], path: str, location: str) -> None:
    table = Table(box=box.SIMPLE, show_header=True, header_style="bold")
    table.add_column("Name")
    table.add_column("Type", width=6)
    table.add_column("Size", justify="right", width=10)
    table.add_column("Modified", width=20)

    # Folders first, then files
    folders = [i for i in items if "folder" in i]
    files = [i for i in items if "file" in i]

    for item in sorted(folders, key=lambda x: x["name"].lower()):
        table.add_row(
            f"[bold blue]{item['name']}[/]",
            "dir",
            "",
            item.get("lastModifiedDateTime", "")[:10],
        )
    for item in sorted(files, key=lambda x: x["name"].lower()):
        table.add_row(
            item["name"],
            "file",
            _fmt_size(item.get("size")),
            item.get("lastModifiedDateTime", "")[:10],
        )

    console.print(f"[dim]{location}[/] [cyan]{path or '/'}[/]")
    console.print(table)
    console.print(f"[dim]{len(folders)} folder(s), {len(files)} file(s)[/]")


def _download_destination(dest: str, remote_name: str) -> Path:
    output = Path(dest or ".").expanduser()
    if output.exists() and output.is_dir():
        return output / remote_name
    if dest.endswith("/") or dest.endswith("\\"):
        output.mkdir(parents=True, exist_ok=True)
        return output / remote_name
    output.parent.mkdir(parents=True, exist_ok=True)
    return output


def _write_download(path: Path, data: bytes, *, overwrite: bool) -> None:
    if path.exists() and not overwrite:
        raise FileExistsError(f"{path} already exists. Use --overwrite to replace it.")
    path.write_bytes(data)


# ── Command handlers ──────────────────────────────────────────────

def cmd_files_sites(args) -> None:
    """List SharePoint sites."""
    gc = _GraphClient()
    try:
        groups = gc.sp_list_sites()
    except OutlookAPIError as e:
        console.print(f"[red]Error:[/] {e}")
        return
    finally:
        gc.close()

    if not groups:
        console.print("No SharePoint sites found.")
        return

    table = Table(box=box.SIMPLE, show_header=True, header_style="bold")
    table.add_column("Name")
    table.add_column("Email")
    table.add_column("Description")

    for g in sorted(groups, key=lambda x: x["displayName"].lower()):
        table.add_row(
            g["displayName"],
            g.get("mail", ""),
            (g.get("description") or "")[:60],
        )
    console.print(table)
    console.print(f"[dim]{len(groups)} site(s)[/]")


def cmd_files_list(args) -> None:
    """List files in OneDrive or a SharePoint site."""
    path: str = getattr(args, "path", "") or ""
    site: str | None = getattr(args, "site", None)

    gc = _GraphClient()
    try:
        if site:
            _, site_id, drive_id = _resolve_site(gc, site)
            item = gc.sp_item_by_path(drive_id, path)
            children = gc.sp_list_children(drive_id, item["id"])
            _print_items(children, path, f"SharePoint: {site}")
        else:
            item = gc.od_item_by_path(path)
            children = gc.od_list_children(item["id"])
            _print_items(children, path, "OneDrive")
    except (OutlookAPIError, ValueError) as e:
        console.print(f"[red]Error:[/] {e}")
    finally:
        gc.close()


def cmd_files_upload(args) -> None:
    """Upload a file to OneDrive or SharePoint."""
    local_path = Path(args.file)
    remote_path: str = args.dest or ""
    site: str | None = getattr(args, "site", None)

    if not local_path.exists():
        console.print(f"[red]File not found:[/] {local_path}")
        return

    name = local_path.name
    size_mb = local_path.stat().st_size / (1024 * 1024)
    console.print(f"[cyan]Uploading[/] {name} ({size_mb:.1f} MB)...")

    gc = _GraphClient()
    try:
        if site:
            _, site_id, drive_id = _resolve_site(gc, site)
            parent = gc.sp_item_by_path(drive_id, remote_path)
            result = gc.sp_upload(drive_id, parent["id"], name, local_path)
        else:
            parent = gc.od_item_by_path(remote_path)
            result = gc.od_upload(parent["id"], name, local_path)
        console.print(f"[green]Uploaded[/] → {result.get('name')} "
                      f"({_fmt_size(result.get('size'))})")
    except (OutlookAPIError, ValueError) as e:
        console.print(f"[red]Error:[/] {e}")
    finally:
        gc.close()


def cmd_files_download(args) -> None:
    """Download a file from OneDrive or SharePoint."""
    remote_path: str = args.path
    dest: str = getattr(args, "dest", ".") or "."
    site: str | None = getattr(args, "site", None)
    overwrite: bool = bool(getattr(args, "overwrite", False))

    gc = _GraphClient()
    try:
        if site:
            _, site_id, drive_id = _resolve_site(gc, site)
            item = gc.sp_item_by_path(drive_id, remote_path)
            if "folder" in item:
                raise ValueError("Download currently supports files, not folders.")
            data = gc.sp_download(drive_id, item["id"])
        else:
            item = gc.od_item_by_path(remote_path)
            if "folder" in item:
                raise ValueError("Download currently supports files, not folders.")
            data = gc.od_download(item["id"])

        output_path = _download_destination(dest, item["name"])
        _write_download(output_path, data, overwrite=overwrite)
        console.print(f"[green]Downloaded[/] {item['name']} → {output_path} "
                      f"({_fmt_size(len(data))})")
    except (OutlookAPIError, ValueError, FileExistsError, OSError) as e:
        console.print(f"[red]Error:[/] {e}")
    finally:
        gc.close()


def cmd_files_mkdir(args) -> None:
    """Create a folder in OneDrive or SharePoint."""
    folder_path: str = args.path
    site: str | None = getattr(args, "site", None)

    # Split into parent path and new folder name
    parts = folder_path.rstrip("/").rsplit("/", 1)
    parent_path = parts[0] if len(parts) == 2 else ""
    name = parts[-1]

    gc = _GraphClient()
    try:
        if site:
            _, site_id, drive_id = _resolve_site(gc, site)
            parent = gc.sp_item_by_path(drive_id, parent_path)
            result = gc.sp_mkdir(drive_id, parent["id"], name)
        else:
            parent = gc.od_item_by_path(parent_path)
            result = gc.od_mkdir(parent["id"], name)
        console.print(f"[green]Created[/] folder: {result.get('name')}")
    except (OutlookAPIError, ValueError) as e:
        console.print(f"[red]Error:[/] {e}")
    finally:
        gc.close()


def cmd_files_rename(args) -> None:
    """Rename a file or folder."""
    item_path: str = args.path
    new_name: str = args.name
    site: str | None = getattr(args, "site", None)

    gc = _GraphClient()
    try:
        if site:
            _, site_id, drive_id = _resolve_site(gc, site)
            item = gc.sp_item_by_path(drive_id, item_path)
            result = gc.sp_rename(drive_id, item["id"], new_name)
        else:
            item = gc.od_item_by_path(item_path)
            result = gc.od_rename(item["id"], new_name)
        console.print(f"[green]Renamed[/] → {result.get('name')}")
    except (OutlookAPIError, ValueError) as e:
        console.print(f"[red]Error:[/] {e}")
    finally:
        gc.close()


def cmd_files_move(args) -> None:
    """Move a file or folder to a different location."""
    item_path: str = args.path
    dest_path: str = args.dest
    site: str | None = getattr(args, "site", None)

    gc = _GraphClient()
    try:
        if site:
            _, site_id, drive_id = _resolve_site(gc, site)
            item = gc.sp_item_by_path(drive_id, item_path)
            dest = gc.sp_item_by_path(drive_id, dest_path)
            result = gc.sp_move(drive_id, item["id"], dest["id"])
        else:
            item = gc.od_item_by_path(item_path)
            dest = gc.od_item_by_path(dest_path)
            result = gc.od_move(item["id"], dest["id"])
        console.print(f"[green]Moved[/] → {result.get('name')}")
    except (OutlookAPIError, ValueError) as e:
        console.print(f"[red]Error:[/] {e}")
    finally:
        gc.close()
