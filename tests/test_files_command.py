from __future__ import annotations

from pathlib import Path

import pytest

from outlook_draft.commands.files import (
    _download_destination,
    _find_od_item,
    _find_sp_item_by_path,
    _GraphClient,
    _item_summary,
    _match_library,
    _write_download,
)
from outlook_draft.errors import OutlookAPIError, TokenExpiredError


class FakeGraphClient:
    def __init__(
        self,
        items_by_drive: dict[str, dict[str, dict]],
        *,
        search_by_drive: dict[str, list[dict]] | None = None,
        items_by_id: dict[str, dict[str, dict]] | None = None,
    ):
        self.items_by_drive = items_by_drive
        self.search_by_drive = search_by_drive or {}
        self.items_by_id = items_by_id or {}

    def sp_list_sites(self) -> list[dict]:
        return [{"id": "group-1", "displayName": "MAP PaidMedia", "groupTypes": ["Unified"]}]

    def sp_get_site(self, group_id: str) -> dict:
        return {"id": "site-1", "displayName": "MAP PaidMedia"}

    def sp_list_drives(self, site_id: str) -> list[dict]:
        return [
            {"id": "drive-1", "name": "Documents", "driveType": "documentLibrary"},
            {"id": "drive-2", "name": "Campaign Assets", "driveType": "documentLibrary"},
        ]

    def sp_item_by_path(self, drive_id: str, path: str) -> dict:
        try:
            return self.items_by_drive[drive_id][path]
        except KeyError as exc:
            raise OutlookAPIError(404, "not found") from exc

    def sp_item_by_id(self, drive_id: str, item_id: str) -> dict:
        try:
            return self.items_by_id[drive_id][item_id]
        except KeyError as exc:
            raise OutlookAPIError(404, "not found") from exc

    def sp_search(self, drive_id: str, query: str, top: int) -> list[dict]:
        return [
            item for item in self.search_by_drive.get(drive_id, [])
            if query.casefold() in item.get("name", "").casefold()
        ][:top]


def test_download_destination_defaults_to_current_directory() -> None:
    assert _download_destination("", "report.pdf") == Path("report.pdf")


def test_download_destination_uses_existing_directory(tmp_path: Path) -> None:
    target_dir = tmp_path / "downloads"
    target_dir.mkdir()

    assert _download_destination(str(target_dir), "report.pdf") == target_dir / "report.pdf"


def test_download_destination_creates_trailing_slash_directory(tmp_path: Path) -> None:
    target_dir = tmp_path / "downloads"

    assert _download_destination(f"{target_dir}/", "report.pdf") == target_dir / "report.pdf"
    assert target_dir.is_dir()


def test_write_download_requires_overwrite_for_existing_file(tmp_path: Path) -> None:
    output = tmp_path / "report.pdf"
    output.write_bytes(b"old")

    with pytest.raises(FileExistsError):
        _write_download(output, b"new", overwrite=False)

    _write_download(output, b"new", overwrite=True)

    assert output.read_bytes() == b"new"


def test_match_library_uses_case_insensitive_partial_name() -> None:
    drive = _match_library(
        [
            {"id": "drive-1", "name": "Documents"},
            {"id": "drive-2", "name": "Campaign Assets"},
        ],
        "campaign",
    )

    assert drive["id"] == "drive-2"


def test_find_sharepoint_path_searches_all_document_libraries() -> None:
    gc = FakeGraphClient({
        "drive-1": {},
        "drive-2": {"Briefs": {"id": "item-1", "name": "Briefs", "folder": {}}},
    })

    drive, item = _find_sp_item_by_path(gc, "paidmedia", "Briefs")

    assert drive["name"] == "Campaign Assets"
    assert item["id"] == "item-1"


def test_find_sharepoint_path_requires_library_when_ambiguous() -> None:
    gc = FakeGraphClient({
        "drive-1": {"General": {"id": "item-1", "name": "General", "folder": {}}},
        "drive-2": {"General": {"id": "item-2", "name": "General", "folder": {}}},
    })

    with pytest.raises(ValueError, match="multiple libraries"):
        _find_sp_item_by_path(gc, "paidmedia", "General")

    drive, item = _find_sp_item_by_path(gc, "paidmedia", "General", "campaign")

    assert drive["name"] == "Campaign Assets"
    assert item["id"] == "item-2"


def test_find_sharepoint_root_file_falls_back_to_exact_search() -> None:
    root_file = {
        "id": "01AGUTGK2IEY4QAP4T2FAZYPYA7R5YGM2X",
        "name": "Root File.pptx",
        "file": {},
        "parentReference": {"path": "/drives/drive-1/root:"},
    }
    gc = FakeGraphClient(
        {"drive-1": {}, "drive-2": {}},
        search_by_drive={"drive-1": [root_file]},
    )

    drive, item = _find_sp_item_by_path(gc, "paidmedia", "Root File.pptx", "Documents")

    assert drive["name"] == "Documents"
    assert item["id"] == "01AGUTGK2IEY4QAP4T2FAZYPYA7R5YGM2X"


def test_find_sharepoint_item_accepts_drive_item_id() -> None:
    item_id = "01AGUTGK2IEY4QAP4T2FAZYPYA7R5YGM2X"
    root_file = {"id": item_id, "name": "Root File.pptx", "file": {}}
    gc = FakeGraphClient(
        {"drive-1": {}, "drive-2": {}},
        items_by_id={"drive-1": {item_id: root_file}},
    )

    drive, item = _find_sp_item_by_path(gc, "paidmedia", item_id, "Documents")

    assert drive["name"] == "Documents"
    assert item["name"] == "Root File.pptx"


def test_item_summary_includes_path_and_web_url() -> None:
    item = {
        "id": "item-1",
        "name": "report.pdf",
        "size": 1234,
        "lastModifiedDateTime": "2026-06-22T10:00:00Z",
        "webUrl": "https://contoso.sharepoint.com/report.pdf",
        "file": {},
        "parentReference": {"path": "/drives/drive-1/root:/General/Reports"},
    }

    summary = _item_summary(item, library="Documents")

    assert summary["type"] == "file"
    assert summary["path"] == "Documents/General/Reports/report.pdf"
    assert summary["parentPath"] == "Documents/General/Reports"
    assert summary["webUrl"] == "https://contoso.sharepoint.com/report.pdf"


def test_graph_client_converts_token_failure_to_api_error() -> None:
    class ExpiredTokenManager:
        def get_token(self, *, auto_reauth: bool = False) -> str:
            raise TokenExpiredError("Graph token has expired")

        def run_reauth(self, *, headless: bool = True) -> bool:
            return False

    gc = _GraphClient()
    gc._tm = ExpiredTokenManager()  # type: ignore[assignment]

    with pytest.raises(OutlookAPIError) as excinfo:
        gc._request("GET", "https://graph.microsoft.com/v1.0/me")

    assert excinfo.value.status == 401
    assert "outlook-cli auth" in str(excinfo.value)


class FakeOneDriveClient:
    def __init__(self, by_id: dict[str, dict], by_path: dict[str, dict]):
        self.by_id = by_id
        self.by_path = by_path

    def od_item_by_id(self, item_id: str) -> dict:
        try:
            return self.by_id[item_id]
        except KeyError as exc:
            raise OutlookAPIError(404, "not found") from exc

    def od_item_by_path(self, path: str) -> dict:
        try:
            return self.by_path[path]
        except KeyError as exc:
            raise OutlookAPIError(404, "not found") from exc


def test_find_onedrive_item_accepts_drive_item_id() -> None:
    item = {"id": "01LT2YGQYBBGUCIETYMVA2PBZIDTQ5ZQMH", "name": "dummy_document.docx"}
    gc = FakeOneDriveClient({item["id"]: item}, {})

    assert _find_od_item(gc, item["id"]) is item


def test_find_onedrive_item_falls_back_to_path() -> None:
    item = {"id": "item-1", "name": "README2026ARCHIVEDNOTES"}
    gc = FakeOneDriveClient({}, {"README2026ARCHIVEDNOTES": item})

    assert _find_od_item(gc, "README2026ARCHIVEDNOTES") is item


def test_find_onedrive_item_uses_path_for_normal_names() -> None:
    item = {"id": "item-2", "name": "report.pdf"}
    gc = FakeOneDriveClient({}, {"Documents/report.pdf": item})

    assert _find_od_item(gc, "Documents/report.pdf") is item


# ── Share URLs and sites outside the user's groups ───────────────

SHARE_URL = "https://tenant.sharepoint.com/:p:/r/sites/Team/Shared%20Documents/Deck.pptx?d=w1&web=1"


class ShareClient:
    def __init__(self, item: dict):
        self.item = item
        self.calls: list[str] = []

    def get_share_drive_item(self, url: str) -> dict:
        self.calls.append(url)
        return self.item

    def download_share_url(self, url: str) -> bytes:
        return b"deck"

    def close(self) -> None:
        pass


def _download(monkeypatch, client, path, dest, **extra):
    import argparse

    from outlook_draft.commands import files

    monkeypatch.setattr(files, "_GraphClient", lambda: client)
    args = argparse.Namespace(path=path, dest=str(dest), site=extra.get("site"), library=None,
                              overwrite=False, no_spinner=True)
    files.cmd_files_download(args)


def test_download_accepts_a_sharepoint_url(monkeypatch, tmp_path: Path, capsys) -> None:
    client = ShareClient({"id": "item", "name": "Deck.pptx", "file": {}})
    _download(monkeypatch, client, SHARE_URL, f"{tmp_path}/")
    assert (tmp_path / "Deck.pptx").read_bytes() == b"deck"
    assert client.calls == [SHARE_URL]
    assert "Downloaded" in capsys.readouterr().out


def test_download_url_refuses_folders_and_other_hosts(monkeypatch, tmp_path: Path, capsys) -> None:
    _download(monkeypatch, ShareClient({"id": "f", "name": "Folder", "folder": {}}), SHARE_URL, f"{tmp_path}/")
    assert "not folders" in capsys.readouterr().out
    other = ShareClient({"id": "x", "name": "x.pptx"})
    _download(monkeypatch, other, "https://example.com/x.pptx", f"{tmp_path}/")
    assert "SharePoint and OneDrive URLs" in capsys.readouterr().out
    assert other.calls == []
    assert list(tmp_path.iterdir()) == []


class SearchingGraphClient(FakeGraphClient):
    def __init__(self, sites: list[dict]):
        super().__init__({})
        self.sites = sites
        self.site_urls: list[str] = []

    def sp_search_sites(self, query: str) -> list[dict]:
        return self.sites

    def sp_site_by_url(self, url: str) -> dict:
        self.site_urls.append(url)
        return {"id": "site-url", "displayName": "From URL"}


AGENTIC = {"id": "site-9", "displayName": "WT-P-U-Consulting-AgenticLabs",
           "webUrl": "https://tenant.sharepoint.com/sites/WT-P-U-Consulting-AgenticLabs"}


def test_site_outside_the_users_groups_is_found_by_search() -> None:
    from outlook_draft.commands.files import _resolve_site_drives

    other = {"id": "site-8", "displayName": "AgenticLabs Archive", "webUrl": "https://tenant.sharepoint.com/sites/AgenticLabsArchive"}
    group, site_id, drives = _resolve_site_drives(SearchingGraphClient([AGENTIC, other]), "WT-P-U-Consulting-AgenticLabs")
    assert site_id == "site-9"
    assert group["id"] is None
    assert drives[0]["id"] == "drive-1"


def test_site_search_reports_ambiguous_and_missing_sites() -> None:
    from outlook_draft.commands.files import _resolve_site_drives

    other = {"id": "site-8", "displayName": "AgenticLabs Archive", "webUrl": "https://tenant.sharepoint.com/sites/AgenticLabsArchive"}
    with pytest.raises(ValueError, match="Ambiguous site"):
        _resolve_site_drives(SearchingGraphClient([AGENTIC, other]), "AgenticLabs")
    with pytest.raises(ValueError, match="Pass the site URL"):
        _resolve_site_drives(SearchingGraphClient([]), "Nowhere")


def test_site_url_resolves_directly() -> None:
    from outlook_draft.commands.files import _resolve_site_drives

    client = SearchingGraphClient([])
    _, site_id, _ = _resolve_site_drives(client, "https://tenant.sharepoint.com/sites/WT-P-U-Consulting-AgenticLabs/Shared%20Documents")
    assert site_id == "site-url"
    assert client.site_urls == ["https://tenant.sharepoint.com/sites/WT-P-U-Consulting-AgenticLabs/Shared%20Documents"]


def test_site_by_url_builds_the_graph_path() -> None:
    from types import SimpleNamespace

    from outlook_draft.outlook_client import OutlookClient

    client = object.__new__(OutlookClient)
    seen = []
    client._request = lambda method, url, **kw: seen.append(url) or SimpleNamespace(json=lambda: {"id": "s"})
    assert client.sp_site_by_url("https://tenant.sharepoint.com/sites/Team/Shared%20Documents/x")["id"] == "s"
    assert seen == ["https://graph.microsoft.com/v1.0/sites/tenant.sharepoint.com:/sites/Team"]
    for bad in ["http://tenant.sharepoint.com/sites/Team", "https://tenant.sharepoint.com/", "https://tenant.sharepoint.com/other/Team"]:
        with pytest.raises(ValueError):
            client.sp_site_by_url(bad)
