"""Fetch OWA email signatures via intercepted API responses.

OWA loads signature HTML from OutlookCloudSettings on inbox load.
We intercept those responses with Playwright instead of scraping the DOM.
"""

from __future__ import annotations

import json
from urllib.parse import parse_qs, urlparse

from rich.console import Console

from .. import config
from ..signatures import sanitize_signature_html

console = Console()

_SETTINGS_URL = "OutlookCloudSettings/settings/account"


def _fetch_signatures_via_browser(headless: bool) -> dict[str, str]:
    """Open OWA and intercept signature API responses.

    Returns a dict with keys 'new' and/or 'reply', values being HTML strings.
    """
    from playwright.sync_api import sync_playwright
    from .. import config as cfg
    from ..auth import _enter_email, _enter_password, _wait_for_mfa

    browser_state = cfg.SESSION_DIR / "browser_state.json"
    captured_by_name: dict[str, list[dict]] = {}  # settingname -> list of items

    def on_response(resp):
        if _SETTINGS_URL not in resp.url:
            return
        try:
            qs = parse_qs(urlparse(resp.url).query)
            name = qs.get("settingname", [""])[0]
            body = resp.json()
            items = body if isinstance(body, list) else [body]
            # Only keep HTML items (type=String with HTML content)
            html_items = [
                item for item in items
                if isinstance(item, dict)
                and item.get("type") == "String"
                and "<" in item.get("value", "")
            ]
            if html_items:
                captured_by_name[name] = html_items
        except Exception:
            pass

    launch_opts: dict = {"headless": headless}
    if not headless:
        launch_opts["args"] = ["--start-maximized"]

    with sync_playwright() as p:
        browser = p.chromium.launch(**launch_opts)
        ctx_opts: dict = {"viewport": {"width": 1280, "height": 900}} if headless else {"no_viewport": True}

        if browser_state.exists():
            ctx_opts["storage_state"] = str(browser_state)
            console.print("[dim]Reusing saved browser session[/]")

        context = browser.new_context(**ctx_opts)
        page = context.new_page()
        page.on("response", on_response)

        page.goto("https://outlook.office.com/mail/", wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(3000)

        # Re-auth if session expired
        if "login.microsoft" in page.url:
            console.print("[yellow]Browser session expired, re-authenticating...[/]")
            if not cfg.MS_EMAIL or not cfg.MS_PASSWORD:
                console.print("[red]MS_EMAIL/MS_PASSWORD not set in .env[/]")
                context.close()
                browser.close()
                return {}
            _enter_email(page, cfg.MS_EMAIL)
            _enter_password(page, cfg.MS_PASSWORD)
            _wait_for_mfa(page)

        # Wait for OWA to fire its settings requests (they fire on load)
        page.wait_for_timeout(6000)

        # If we didn't catch them on inbox load, open settings to trigger them
        if not captured_by_name:
            console.print("[dim]Settings not seen on inbox load, opening settings panel...[/]")
            try:
                page.locator('[aria-label*="Setting"]').first.click(timeout=5000)
                page.wait_for_timeout(2000)
                page.get_by_role("tab", name="Compose and reply").first.click(timeout=5000)
                page.wait_for_timeout(4000)
            except Exception:
                pass

        context.close()
        browser.close()

    if not captured_by_name:
        return {}

    # Now we need to map the captured HTML by setting name to new/reply
    # The setting names are the display names of the signatures as configured in OWA
    # We need to look up which name is assigned to new vs reply
    # Those mappings come from the roaming_signature_list call, but we can also
    # fetch them directly via the token
    import requests
    try:
        tokens_file = cfg.TOKENS_FILE
        token_data = json.loads(tokens_file.read_text())
        token = token_data["tokens"].get("outlook.office.com", "")
        headers = {"Authorization": f"Bearer {token}"}

        meta_resp = requests.get(
            "https://outlook.office.com/ows/v1/OutlookCloudSettings/settings/",
            params={"settingname": "roaming_signature_list,roaming_new_signature,roaming_reply_signature"},
            headers=headers,
            timeout=15,
        )
        meta = meta_resp.json() if meta_resp.status_code == 200 else []
    except Exception:
        meta = []

    new_name = None
    reply_name = None
    for item in meta if isinstance(meta, list) else []:
        n = (item.get("name") or "").lower()
        if n == "roaming_new_signature":
            new_name = item.get("value", "")
        elif n == "roaming_reply_signature":
            reply_name = item.get("value", "")

    result: dict[str, str] = {}

    # Match by name if we got the metadata
    if new_name and new_name in captured_by_name:
        result["new"] = captured_by_name[new_name][0]["value"]
    if reply_name and reply_name in captured_by_name:
        result["reply"] = captured_by_name[reply_name][0]["value"]

    # Fallback: if metadata unavailable, use whatever we captured
    if not result and captured_by_name:
        names = list(captured_by_name.keys())
        result["new"] = captured_by_name[names[0]][0]["value"]
        if len(names) >= 2:
            result["reply"] = captured_by_name[names[1]][0]["value"]
        console.print(f"[yellow]Could not determine which is new/reply; used first two captured: {names[:2]}[/]")

    return result


def cmd_signature_fetch(args) -> None:
    """Fetch OWA signatures and save to configured files."""
    headless = not getattr(args, "headed", False)

    console.print("[cyan]Fetching signatures from OWA...[/]")

    sigs = _fetch_signatures_via_browser(headless)

    if not sigs:
        console.print("[red]No signatures found.[/]")
        console.print("Try running [bold]outlook-cli auth[/] first, then retry.")
        console.print("Or use [bold]--headed[/] to run with a visible browser window.")
        return

    saved = []

    if "new" in sigs:
        html = sanitize_signature_html(sigs["new"])
        config.SIGNATURE_NEW_FILE.write_text(html, encoding="utf-8")
        saved.append(("New message", config.SIGNATURE_NEW_FILE))

    if "reply" in sigs:
        html = sanitize_signature_html(sigs["reply"])
        config.SIGNATURE_REPLY_FILE.write_text(html, encoding="utf-8")
        saved.append(("Reply/forward", config.SIGNATURE_REPLY_FILE))

    for label, path in saved:
        console.print(f"[green]Saved[/] {label} signature → {path}")

    if len(sigs) == 1 and "reply" not in sigs:
        console.print(
            "[yellow]Only a new-message signature was found. "
            "If you have a separate reply signature, check OWA settings.[/]"
        )
