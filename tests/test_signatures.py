from __future__ import annotations

from pathlib import Path

import pytest

from outlook_draft import config
from outlook_draft.signatures import load_signature, sanitize_signature_html


def test_resolve_config_path_uses_default_when_env_missing(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("SIGNATURE_NEW_FILE", raising=False)
    default = Path("/tmp/default-signature.html")

    assert config.resolve_config_path("SIGNATURE_NEW_FILE", default) == default


def test_resolve_config_path_resolves_relative_paths_from_project_root(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("SIGNATURE_NEW_FILE", "custom/signature.html")
    default = Path("/tmp/default-signature.html")

    assert config.resolve_config_path("SIGNATURE_NEW_FILE", default) == (
        config._project_root / "custom/signature.html"
    )


def test_resolve_config_path_expands_absolute_paths(monkeypatch: pytest.MonkeyPatch, tmp_path: Path) -> None:
    custom = tmp_path / "signature.html"
    monkeypatch.setenv("SIGNATURE_NEW_FILE", str(custom))

    assert config.resolve_config_path("SIGNATURE_NEW_FILE", Path("unused.html")) == custom


def test_sanitize_signature_html_removes_outlook_noise() -> None:
    html = '''
    <div id="x" class="y" style="color: var(--darkColor_black, black)">
      <span></span>
      <img src="blob:https://outlook.office.com/example">
      <span>Hello</span>
    </div>
    '''

    sanitized = sanitize_signature_html(html)

    assert 'id="x"' not in sanitized
    assert 'class="y"' not in sanitized
    assert "blob:" not in sanitized
    assert "var(--darkColor_black" not in sanitized
    assert "<span></span>" not in sanitized
    assert "Hello" in sanitized


def test_load_signature_sanitizes_file(tmp_path: Path) -> None:
    signature = tmp_path / "signature.html"
    signature.write_text('<div id="x"><span>Hello</span></div>', encoding="utf-8")

    assert load_signature(signature) == "<div><span>Hello</span></div>"


def test_load_signature_raises_for_missing_file(tmp_path: Path) -> None:
    with pytest.raises(FileNotFoundError, match="Signature file not found"):
        load_signature(tmp_path / "missing.html")
