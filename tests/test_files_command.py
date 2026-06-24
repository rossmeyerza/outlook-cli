from __future__ import annotations

from pathlib import Path

import pytest

from outlook_draft.errors import OutlookAPIError, TokenExpiredError
from outlook_draft.commands.files import (
    _GraphClient,
    _find_sp_item_by_path,
    _item_summary,
    _match_library,
    _download_destination,
    _write_download,
)


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
