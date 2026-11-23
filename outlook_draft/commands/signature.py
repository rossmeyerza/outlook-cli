"""Fetch OWA email signatures via intercepted API responses.

OWA loads signatures and their assignments from OutlookCloudSettings.
The signature settings panel can trigger these responses.
"""

from __future__ import annotations

import logging
from urllib.parse import parse_qs, urlparse

from rich.console import Console

from .. import config
from ..auth import _on_outlook
from ..errors import CliError
from ..signatures import sanitize_signature_html

console = Console()

_SETTINGS_PATH = "/ows/v1/outlookcloudsettings/settings"
_ASSIGNMENTS = {"roaming_new_signature": "new", "roaming_reply_signature": "reply"}
_SIGNATURE_LIST = "roaming_signature_list"
_OUTLOOK_MAIL = "https://outlook.cloud.microsoft/mail/"
_SIGN_IN_REQUIRED = (
    "The saved Outlook session could not sign in without user action. "
    "Run 'outlook-cli auth' in a terminal, then retry 'outlook-cli signature fetch'."
)


log = logging.getLogger(__name__)


def _open_outlook(page) -> None:
    from playwright.sync_api import Error

    def open_mail():
        page.goto(_OUTLOOK_MAIL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(6000)

    try:
        open_mail()
        url = urlparse(page.url)
        if (
            url.scheme == "https"
            and (url.hostname or "").endswith(".okta.com")
            and url.path.rstrip("/").lower() == "/app/userhome"
        ):
            # Start SSO again after the sign-out returns to Okta home.
            open_mail()
    except Error:
        if not _on_outlook(page.url):
            raise CliError(_SIGN_IN_REQUIRED) from None
        raise CliError("Could not load Outlook Web. Check the connection and retry.") from None
    if not _on_outlook(page.url):
        raise CliError(_SIGN_IN_REQUIRED)


def _capture_response(html_by_name: dict[str, str], assignments: dict[str, str]):
    from playwright.sync_api import Error

    def on_response(resp):
        url = urlparse(resp.url)
        path = url.path.rstrip("/").lower()
        if not _on_outlook(resp.url) or path not in {_SETTINGS_PATH, f"{_SETTINGS_PATH}/account"}:
            return
        try:
            body = resp.json()
        except (Error, ValueError):
            log.debug("Could not read the Outlook settings response.")
            return
        query_name = parse_qs(url.query).get("settingname", [""])[0]
        for item in body if isinstance(body, list) else [body]:
            if not isinstance(item, dict) or item.get("type") != "String":
                continue
            name, value = item.get("name"), item.get("value")
            if not isinstance(value, str):
                continue
            kind = _ASSIGNMENTS.get(name.lower()) if isinstance(name, str) else None
            if kind:
                assignments[kind] = value
            elif path.endswith("/account") and value.lstrip().startswith("<"):
                if item.get("parentSetting") == _SIGNATURE_LIST:
                    for key in (item.get("secondaryKey"), name):
                        if isinstance(key, str) and key:
                            html_by_name[key] = value
                elif query_name:
                    html_by_name[query_name] = value

    return on_response


def _open_signature_settings(page) -> None:
    from playwright.sync_api import Error

    console.print("[dim]Opening Outlook signature settings...[/]")
    try:
        settings = page.get_by_role("button", name="Settings", exact=True)
        if settings.count():
            settings.click(timeout=5000)
        else:
            # Outlook can hide its toolbar button at the headless viewport size.
            page.locator('button[aria-label="Settings"]').evaluate("e => e.click()", timeout=5000)
        account = page.get_by_text("Account", exact=True).filter(visible=True).first
        try:
            account.click(timeout=5000)
            page.get_by_text("Signatures", exact=True).filter(visible=True).first.click(timeout=5000)
        except Error:
            page.get_by_role("tab", name="Compose and reply").first.click(timeout=5000)
        page.wait_for_timeout(6000)
    except Error:
        log.debug("Could not open Outlook signature settings.")


def _assigned_signatures(html_by_name: dict[str, str], assignments: dict[str, str]) -> dict[str, str]:
    return {kind: html_by_name[name] for kind, name in assignments.items() if name in html_by_name}


def _capture_signatures(headless: bool) -> dict[str, str]:
    from playwright.sync_api import sync_playwright

    config.ensure_dirs()
    browser_state = config.SESSION_DIR / "browser_state.json"
    if not browser_state.exists():
        raise CliError(_SIGN_IN_REQUIRED)
    html_by_name: dict[str, str] = {}
    assignments: dict[str, str] = {}

    launch_opts: dict = {"headless": headless}
    if not headless:
        launch_opts["args"] = ["--start-maximized"]

    with sync_playwright() as p:
        browser = p.chromium.launch(**launch_opts)
        ctx_opts: dict = {"viewport": {"width": 1280, "height": 900}} if headless else {"no_viewport": True}

        console.print("[dim]Reusing saved browser session[/]")
        try:
            context = browser.new_context(storage_state=str(browser_state), **ctx_opts)
            try:
                page = context.new_page()
                page.on("response", _capture_response(html_by_name, assignments))
                _open_outlook(page)
                if len(_assigned_signatures(html_by_name, assignments)) < 2:
                    _open_signature_settings(page)
                    if not _on_outlook(page.url):
                        raise CliError(_SIGN_IN_REQUIRED)
            finally:
                context.close()
        finally:
            browser.close()
    return _assigned_signatures(html_by_name, assignments)


def cmd_signature_fetch(args) -> None:
    """Fetch OWA signatures and save to configured files."""
    headless = not getattr(args, "headed", False)

    console.print("[cyan]Fetching signatures from OWA...[/]")

    sigs = _capture_signatures(headless)

    if not sigs:
        raise CliError("No assigned signatures were captured. Check Settings > Account > Signatures in Outlook Web.")

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


def register(sub) -> None:
    p_sig = sub.add_parser("signature", help="Manage email signatures")
    sig_sub = p_sig.add_subparsers(dest="command", required=True)
    cmd_sig_fetch = sig_sub.add_parser(
        "fetch",
        help="Fetch your OWA signature and save to configured files",
    )
    cmd_sig_fetch.add_argument(
        "--headed",
        action="store_true",
        help="Run with a visible browser window (useful if headless fails)",
    )
    cmd_sig_fetch.set_defaults(func=cmd_signature_fetch)
