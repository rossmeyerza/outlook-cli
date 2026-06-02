from __future__ import annotations

import argparse
import html
import json
import re
import sys
from html.parser import HTMLParser
from pathlib import Path
from typing import Any, Callable

from rich.table import Table

from ..cache import (
    NOTES_NOTEBOOK_CACHE,
    NOTES_PAGE_CACHE,
    NOTES_SECTION_CACHE,
    load_cache,
    save_cache,
)
from ..errors import OutlookAPIError


JsonDict = dict[str, Any]


def _ctx(args: argparse.Namespace) -> dict[str, Any]:
    return args._notes_ctx


def _console(args: argparse.Namespace):
    return _ctx(args)["console"]


def _get_graph_client(args: argparse.Namespace):
    return _ctx(args)["get_graph_client"]()


def _format_datetime(args: argparse.Namespace, value: str) -> str:
    return _ctx(args)["format_datetime"](value)


def build_ctx(
    *,
    console: Any,
    get_graph_client: Callable[[], Any],
    format_datetime: Callable[[str], str],
) -> dict[str, Any]:
    return {
        "console": console,
        "get_graph_client": get_graph_client,
        "format_datetime": format_datetime,
    }


class _TextExtractor(HTMLParser):
    def __init__(self) -> None:
        super().__init__(convert_charrefs=True)
        self.parts: list[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        if tag.lower() == "br":
            self.parts.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if tag.lower() in {"p", "div", "li", "tr", "h1", "h2", "h3", "h4", "h5", "h6"}:
            self.parts.append("\n")

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

    def text(self) -> str:
        text = "".join(self.parts)
        text = re.sub(r"[ \t]+", " ", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        return text.strip()


def _html_to_text(value: str) -> str:
    parser = _TextExtractor()
    parser.feed(value)
    return parser.text()


def _plain_text_to_html(text: str) -> str:
    text = text.replace("\r\n", "\n").strip()
    if not text:
        return "<p></p>"
    parts = []
    for block in text.split("\n\n"):
        block = block.strip()
        if not block:
            continue
        parts.append(f"<p>{html.escape(block).replace(chr(10), '<br>')}</p>")
    return "".join(parts) or "<p></p>"


def _fallback_markdown_to_html(text: str) -> str:
    lines = text.replace("\r\n", "\n").splitlines()
    out: list[str] = []
    list_items: list[str] = []
    paragraph: list[str] = []

    def flush_paragraph() -> None:
        if paragraph:
            out.append(f"<p>{html.escape(' '.join(paragraph))}</p>")
            paragraph.clear()

    def flush_list() -> None:
        if list_items:
            out.append("<ul>" + "".join(f"<li>{html.escape(item)}</li>" for item in list_items) + "</ul>")
            list_items.clear()

    for line in lines:
        stripped = line.strip()
        if not stripped:
            flush_paragraph()
            flush_list()
            continue
        heading = re.match(r"^(#{1,6})\s+(.+)$", stripped)
        if heading:
            flush_paragraph()
            flush_list()
            level = len(heading.group(1))
            out.append(f"<h{level}>{html.escape(heading.group(2))}</h{level}>")
            continue
        item = re.match(r"^[-*]\s+(.+)$", stripped)
        if item:
            flush_paragraph()
            list_items.append(item.group(1))
            continue
        flush_list()
        paragraph.append(stripped)

    flush_paragraph()
    flush_list()
    return "".join(out) or "<p></p>"


def _markdown_to_html(text: str) -> str:
    try:
        import markdown  # type: ignore

        return markdown.markdown(text, extensions=["extra", "sane_lists"])
    except Exception:
        return _fallback_markdown_to_html(text)


def _load_text_arg(args: argparse.Namespace, *, value_attr: str, file_attr: str, label: str) -> str:
    value = getattr(args, value_attr, None)
    file_value = getattr(args, file_attr, None)
    if value and file_value:
        _console(args).print(f"[red]Use either --{label} or --{label}-file, not both.[/]")
        sys.exit(1)
    if file_value:
        path = Path(file_value)
        if not path.exists():
            _console(args).print(f"[red]File not found: {path}[/]")
            sys.exit(1)
        return path.read_text(encoding="utf-8")
    if value is not None:
        return value
    _console(args).print(f"[red]Provide --{label} or --{label}-file.[/]")
    sys.exit(1)


def _load_body_html(args: argparse.Namespace) -> str:
    body = _load_text_arg(args, value_attr="body", file_attr="body_file", label="body")
    modes = [bool(getattr(args, "markdown", False)), bool(getattr(args, "html", False))]
    if sum(modes) > 1:
        _console(args).print("[red]Use only one body format flag.[/]")
        sys.exit(1)
    if getattr(args, "html", False):
        return body
    if getattr(args, "markdown", False):
        return _markdown_to_html(body)
    return _plain_text_to_html(body)


def _resolve_cached_ref(cache_path: Path, ref: str, label: str, console: Any) -> str:
    cached = load_cache(cache_path)
    if ref.isdigit():
        idx = int(ref)
        if cached and 1 <= idx <= len(cached):
            return cached[idx - 1]["Id"]
        if not cached:
            console.print(f"[red]No cached {label}s. Run a notes list command first.[/]")
        else:
            console.print(f"[red]{label.title()} index {idx} out of range. Only {len(cached)} cached.[/]")
        sys.exit(1)
    if len(ref) < 40 and cached:
        matches = [item for item in cached if item["Id"].endswith(ref)]
        if len(matches) == 1:
            return matches[0]["Id"]
        if len(matches) > 1:
            console.print(f"[red]Ambiguous {label} ID suffix '{ref}'. Use a longer ID.[/]")
            sys.exit(1)
    return ref


def _resolve_notebook_id(args: argparse.Namespace, ref: str) -> str:
    return _resolve_cached_ref(NOTES_NOTEBOOK_CACHE, ref, "notebook", _console(args))


def _resolve_section_id(args: argparse.Namespace, ref: str) -> str:
    return _resolve_cached_ref(NOTES_SECTION_CACHE, ref, "section", _console(args))


def _resolve_page_id(args: argparse.Namespace, ref: str) -> str:
    return _resolve_cached_ref(NOTES_PAGE_CACHE, ref, "page", _console(args))


def _web_url(item: JsonDict) -> str:
    return ((item.get("links") or {}).get("oneNoteWebUrl") or {}).get("href") or item.get("webUrl") or ""


def _exit_api_error(args: argparse.Namespace, exc: OutlookAPIError) -> None:
    if exc.status in {0, 401, 403}:
        _console(args).print(
            "[red]OneNote Graph request failed.[/] "
            "Check that the Graph token includes Notes.ReadWrite with "
            "`outlook-cli auth scopes`, then re-authenticate with `outlook-cli auth --headed` if needed."
        )
    else:
        _console(args).print(f"[red]OneNote Graph request failed: {exc}[/]")
    sys.exit(1)


def _simplify_notebook(item: JsonDict, index: int) -> JsonDict:
    return {
        "index": index,
        "id": item.get("id", ""),
        "name": item.get("displayName", ""),
        "lastModifiedDateTime": item.get("lastModifiedDateTime", ""),
        "webUrl": _web_url(item),
    }


def _simplify_section(item: JsonDict, index: int) -> JsonDict:
    parent = item.get("parentNotebook") or {}
    return {
        "index": index,
        "id": item.get("id", ""),
        "name": item.get("displayName", ""),
        "notebook": parent.get("displayName", ""),
        "lastModifiedDateTime": item.get("lastModifiedDateTime", ""),
        "webUrl": _web_url(item),
    }


def _simplify_page(item: JsonDict, index: int) -> JsonDict:
    parent = item.get("parentSection") or {}
    return {
        "index": index,
        "id": item.get("id", ""),
        "title": item.get("title", ""),
        "section": parent.get("displayName", ""),
        "lastModifiedDateTime": item.get("lastModifiedDateTime", ""),
        "webUrl": _web_url(item),
    }


def cmd_notebooks(args: argparse.Namespace) -> None:
    client = _get_graph_client(args)
    try:
        items = client.list_onenote_notebooks(top=args.top)
    except OutlookAPIError as exc:
        _exit_api_error(args, exc)
    finally:
        client.close()
    save_cache(NOTES_NOTEBOOK_CACHE, items, id_key="id")
    rows = [_simplify_notebook(item, i) for i, item in enumerate(items, 1)]
    if args.json:
        print(json.dumps(rows, indent=2))
        return
    table = Table(title=f"OneNote notebooks ({len(rows)})")
    table.add_column("#", justify="right")
    table.add_column("Name")
    table.add_column("Modified")
    table.add_column("ID suffix")
    for row in rows:
        table.add_row(str(row["index"]), row["name"], _format_datetime(args, row["lastModifiedDateTime"]), row["id"][-8:])
    _console(args).print(table)


def cmd_sections(args: argparse.Namespace) -> None:
    notebook_id = _resolve_notebook_id(args, args.notebook) if args.notebook else None
    client = _get_graph_client(args)
    try:
        items = client.list_onenote_sections(notebook_id=notebook_id, top=args.top)
    except OutlookAPIError as exc:
        _exit_api_error(args, exc)
    finally:
        client.close()
    save_cache(NOTES_SECTION_CACHE, items, id_key="id")
    rows = [_simplify_section(item, i) for i, item in enumerate(items, 1)]
    if args.json:
        print(json.dumps(rows, indent=2))
        return
    table = Table(title=f"OneNote sections ({len(rows)})")
    table.add_column("#", justify="right")
    table.add_column("Section")
    table.add_column("Notebook")
    table.add_column("Modified")
    table.add_column("ID suffix")
    for row in rows:
        table.add_row(
            str(row["index"]),
            row["name"],
            row["notebook"],
            _format_datetime(args, row["lastModifiedDateTime"]),
            row["id"][-8:],
        )
    _console(args).print(table)


def cmd_pages(args: argparse.Namespace) -> None:
    section_id = _resolve_section_id(args, args.section) if args.section else None
    client = _get_graph_client(args)
    try:
        items = client.list_onenote_pages(section_id=section_id, top=args.top)
    except OutlookAPIError as exc:
        _exit_api_error(args, exc)
    finally:
        client.close()
    save_cache(NOTES_PAGE_CACHE, items, id_key="id")
    rows = [_simplify_page(item, i) for i, item in enumerate(items, 1)]
    if args.json:
        print(json.dumps(rows, indent=2))
        return
    table = Table(title=f"OneNote pages ({len(rows)})")
    table.add_column("#", justify="right")
    table.add_column("Title")
    table.add_column("Section")
    table.add_column("Modified")
    table.add_column("ID suffix")
    for row in rows:
        table.add_row(
            str(row["index"]),
            row["title"],
            row["section"],
            _format_datetime(args, row["lastModifiedDateTime"]),
            row["id"][-8:],
        )
    _console(args).print(table)


def cmd_read(args: argparse.Namespace) -> None:
    page_id = _resolve_page_id(args, args.page)
    client = _get_graph_client(args)
    try:
        page = client.get_onenote_page(page_id)
        content = client.get_onenote_page_content(page_id)
    except OutlookAPIError as exc:
        _exit_api_error(args, exc)
    finally:
        client.close()
    if args.json:
        print(json.dumps({"page": _simplify_page(page, 1), "html": content, "text": _html_to_text(content)}, indent=2))
        return
    if args.html:
        print(content)
        return
    _console(args).print(f"[bold]{page.get('title') or '(untitled)'}[/]\n")
    _console(args).print(_html_to_text(content))


def cmd_create(args: argparse.Namespace) -> None:
    section_id = _resolve_section_id(args, args.section) if args.section else None
    body_html = _load_body_html(args)
    client = _get_graph_client(args)
    try:
        page = client.create_onenote_page(title=args.title, html_body=body_html, section_id=section_id)
    except OutlookAPIError as exc:
        _exit_api_error(args, exc)
    finally:
        client.close()
    if args.json:
        print(json.dumps(_simplify_page(page, 1), indent=2))
        return
    _console(args).print(f"[green]Created OneNote page:[/] {page.get('title') or args.title}")
    if _web_url(page):
        _console(args).print(_web_url(page))


def cmd_append(args: argparse.Namespace) -> None:
    page_id = _resolve_page_id(args, args.page)
    body_html = _load_body_html(args)
    client = _get_graph_client(args)
    try:
        client.append_onenote_page_content(page_id, body_html)
    except OutlookAPIError as exc:
        _exit_api_error(args, exc)
    finally:
        client.close()
    result = {"ok": True, "page_id": page_id, "action": "append"}
    if args.json:
        print(json.dumps(result, indent=2))
        return
    _console(args).print("[green]Appended content to OneNote page.[/]")


_REPLACEABLE_RE = re.compile(
    r"<(?P<tag>p|li|h[1-6])\b(?P<attrs>[^>]*)>(?P<inner>.*?)</(?P=tag)>",
    re.IGNORECASE | re.DOTALL,
)
_TITLE_RE = re.compile(r"<title\b[^>]*>(?P<inner>.*?)</title>", re.IGNORECASE | re.DOTALL)
_ID_RE = re.compile(r"""\bid=["'](?P<id>[^"']+)["']""", re.IGNORECASE)


def _candidate_content(tag: str, text: str) -> str:
    if tag.lower() == "title":
        return text
    escaped = html.escape(text).replace("\n", "<br>")
    return f"<{tag.lower()}>{escaped}</{tag.lower()}>"


def _replace_change_from_html(page_html: str, old: str, new: str) -> dict[str, str]:
    if not old:
        raise ValueError("old text must not be empty")

    matches: list[tuple[str, str, str]] = []

    title_match = _TITLE_RE.search(page_html)
    if title_match:
        title_text = _html_to_text(title_match.group("inner"))
        if old in title_text:
            for _ in range(title_text.count(old)):
                matches.append(("title", "title", title_text))

    for match in _REPLACEABLE_RE.finditer(page_html):
        id_match = _ID_RE.search(match.group("attrs"))
        if not id_match:
            continue
        text = _html_to_text(match.group("inner"))
        count = text.count(old)
        for _ in range(count):
            matches.append((id_match.group("id"), match.group("tag"), text))

    if not matches:
        raise ValueError("old text did not match the page content exactly")
    if len(matches) > 1:
        raise ValueError(f"old text matched {len(matches)} places; refine it so it matches exactly once")

    target, tag, text = matches[0]
    replaced = text.replace(old, new, 1)
    return {
        "target": target,
        "action": "replace",
        "content": _candidate_content(tag, replaced),
    }


def cmd_replace(args: argparse.Namespace) -> None:
    page_id = _resolve_page_id(args, args.page)
    old = _load_text_arg(args, value_attr="old", file_attr="old_file", label="old")
    new = _load_text_arg(args, value_attr="new", file_attr="new_file", label="new")
    client = _get_graph_client(args)
    try:
        page_html = client.get_onenote_page_content(page_id, include_ids=True)
        change = _replace_change_from_html(page_html, old, new)
        client.patch_onenote_page_content(page_id, [change])
    except OutlookAPIError as exc:
        _exit_api_error(args, exc)
    except ValueError as exc:
        _console(args).print(f"[red]Could not replace text: {exc}[/]")
        sys.exit(1)
    finally:
        client.close()
    result = {"ok": True, "page_id": page_id, "action": "replace", "target": change["target"]}
    if args.json:
        print(json.dumps(result, indent=2))
        return
    _console(args).print(f"[green]Replaced exact text in OneNote page.[/] Target: {change['target']}")
