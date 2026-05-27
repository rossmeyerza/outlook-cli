from __future__ import annotations

from pathlib import Path

import pytest

from outlook_draft.commands.files import _download_destination, _write_download


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
