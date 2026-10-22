"""Microsoft SSO token capture via Playwright.

Supports both headed (interactive) and headless modes.
Headless mode enters credentials automatically, displays the MFA
challenge number in the console, and waits for push approval.

Supports Okta Verify push and other MFA providers that display a
numeric challenge on screen.

Tokens are cached to disk and reused until expired.
Self-contained auth for outlook-draft-cli.

Usage:
  python auth.py                # headed, auto-fill credentials
  python auth.py --headless     # fully headless, MFA number in console
  python auth.py --force         # force re-auth even if tokens are valid
"""

import base64
import json
import logging
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from urllib.parse import parse_qs, urlsplit

import httpx
from playwright.sync_api import sync_playwright
from rich.console import Console

from . import config

console = Console()

config.ensure_dirs()

# Microsoft API domains we look for when intercepting requests.
KNOWN_DOMAINS = [
    "graph.microsoft.com",
    "outlook.office365.com",
    "outlook.office.com",
]


# The Microsoft 365 Copilot web app gets a token for this audience.
COPILOT_CHAT_URL = "https://m365.cloud.microsoft/chat"
COPILOT_AUDIENCE = "https://substrate.office.com/sydney"

# Buttons on Microsoft sign-in pages that can appear after SSO:
# "Do you trust <domain>?", "Stay signed in?" and the account picker.
LOGIN_INTERSTITIAL_SELECTORS = (
    "input#idSIButton9:visible",
    'button:has-text("Continue"):visible',
    'input[value="Continue"]:visible',
    'input[value="Yes"]:visible',
    "div[data-test-id]:visible >> nth=0",
)


log = logging.getLogger(__name__)


def _token_claims(token: str) -> dict:
    """Decode JWT claims without verification."""
    try:
        payload = token.split(".")[1]
        payload += "=" * (4 - len(payload) % 4)
        return json.loads(base64.urlsafe_b64decode(payload))
    except Exception:
        log.debug("Best-effort operation failed", exc_info=True)
        return {}


def _token_scopes(token: str) -> set[str]:
    """Decode JWT scopes without verification."""
    return set(_token_claims(token).get("scp", "").split())


def _save_tokens(tokens: dict[str, str]):
    """Save all captured tokens keyed by domain."""
    config.ensure_dirs()
    data = {
        "tokens": tokens,
        "captured_at": datetime.now().isoformat(),
    }
    config.TOKENS_FILE.write_text(json.dumps(data, indent=2))


def _load_tokens() -> dict[str, str]:
    """Load saved tokens from disk."""
    if not config.TOKENS_FILE.exists():
        return {}
    try:
        data = json.loads(config.TOKENS_FILE.read_text())
        return data.get("tokens", {})
    except Exception:
        log.debug("Best-effort operation failed", exc_info=True)
        return {}


def _test_token(token: str, url: str) -> bool:
    """Quick check if a token works against a URL."""
    try:
        resp = httpx.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            timeout=10,
            follow_redirects=True,
        )
        return resp.status_code == 200
    except Exception:
        log.debug("Best-effort operation failed", exc_info=True)
        return False


def _make_request_interceptor(captured: dict[str, str]):
    """Create a request handler that captures bearer tokens."""
    def on_request(request):
        auth = request.headers.get("authorization", "")
        if not auth.startswith("Bearer "):
            return
        url = request.url
        for domain in KNOWN_DOMAINS:
            if domain in url:
                new_token = auth.removeprefix("Bearer ")
                token_domain = domain
                if token_domain not in captured:
                    captured[token_domain] = new_token
                    console.print(f"[dim]  Captured token for {token_domain}[/]")
                else:
                    new_scopes = _token_scopes(new_token)
                    old_scopes = _token_scopes(captured[token_domain])
                    if len(new_scopes) > len(old_scopes):
                        captured[token_domain] = new_token
                        console.print(f"[dim]  Updated token for {token_domain} (broader scopes)[/]")
                break
    return on_request


def _find_mfa_number(page) -> str | None:
    """Find the numeric MFA challenge on the login page (Okta Verify and similar)."""
    return page.evaluate("""() => {
        // Try known Okta selectors first
        for (const sel of [
            '[data-se="challenge-number"]',
            '[data-se="number-challenge"]',
            '[class*="number-challenge"]',
            '[class*="challenge-number"]',
        ]) {
            const el = document.querySelector(sel);
            if (el) {
                const t = el.textContent?.trim() ?? '';
                if (/^\\d+$/.test(t)) return t;
            }
        }

        // Fallback: find the largest prominent number on the page
        let best = null;
        let bestSize = 0;
        for (const el of document.querySelectorAll('h1,h2,h3,span,div,p,strong,b')) {
            const t = el.textContent?.trim() ?? '';
            if (!/^\\d{1,3}$/.test(t)) continue;
            const rect = el.getBoundingClientRect();
            if (rect.width === 0 || rect.height === 0) continue;
            const size = parseFloat(getComputedStyle(el).fontSize);
            if (size > bestSize) { bestSize = size; best = t; }
        }
        return best;
    }""")


def _enter_email(page, email: str) -> bool:
    """Enter email on the Microsoft login page."""
    try:
        page.wait_for_selector('input[type="email"]', timeout=10000)
        page.fill('input[type="email"]', email)
        page.click('input[type="submit"]')
        console.print("[dim]  Entered email[/]")
        page.wait_for_timeout(3000)
        return True
    except Exception as e:
        log.debug("Best-effort operation failed", exc_info=True)
        console.print(f"[red]Failed to enter email: {e}[/]")
        return False


def _body_text(page) -> str:
    try:
        value = page.locator("body").inner_text(timeout=2000)
        return value or ""
    except Exception:
        log.debug("Best-effort operation failed", exc_info=True)
        return ""


def _has_okta_password_error(page) -> bool:
    text = _body_text(page).lower()
    return "unable to sign in" in text or "incorrect password" in text or "password is incorrect" in text


def _enter_password(page, password: str) -> bool:
    """Enter password on the Microsoft/SSO login page."""
    try:
        # Try Okta-style selector first, then generic
        for selector in [
            '[name="credentials.passcode"]',
            'input[type="password"]:visible',
        ]:
            try:
                page.wait_for_selector(selector, timeout=15000)
                page.wait_for_timeout(500)
                page.fill(selector, password)
                break
            except Exception:
                log.debug("Best-effort operation failed", exc_info=True)
                continue
        else:
            console.print("[red]Could not find password field[/]")
            return False

        # Click submit
        for btn_sel in [
            'button[type="submit"]',
            'input[type="submit"]',
            'input[value="Sign in"]',
            '#okta-signin-submit',  # Okta-hosted login
        ]:
            try:
                btn = page.locator(btn_sel).first
                if btn.is_visible(timeout=2000):
                    btn.click()
                    break
            except Exception:
                log.debug("Best-effort operation failed", exc_info=True)
                continue

        console.print("[dim]  Entered password[/]")
        page.wait_for_timeout(3000)
        if _has_okta_password_error(page):
            raise RuntimeError(
                "Outlook Okta rejected the password before MFA. "
                "Update MS_PASSWORD in ~/.config/outlook-cli/.env."
            )
        return True
    except RuntimeError:
        raise
    except Exception as e:
        log.debug("Best-effort operation failed", exc_info=True)
        console.print(f"[red]Failed to enter password: {e}[/]")
        return False


def _wait_for_mfa(page) -> None:
    """Detect MFA challenge, display number, and wait for approval."""
    console.print("[dim]Waiting for MFA challenge...[/]")
    mfa_number = None

    for _ in range(30):
        page.wait_for_timeout(500)
        if _has_okta_password_error(page):
            raise RuntimeError(
                "Outlook Okta rejected the password before MFA. "
                "Update MS_PASSWORD in ~/.config/outlook-cli/.env."
            )
        mfa_number = _find_mfa_number(page)
        if mfa_number:
            break
        # Check if we already got through (no MFA needed)
        try:
            url = page.url
            if "outlook.office.com" in url and "login" not in url and "authn" not in url:
                return
        except Exception:
            log.debug("Best-effort operation failed", exc_info=True)
            pass

    if mfa_number:
        console.print(
            f"\n[bold yellow]MFA Verification: tap"
            f" [bold white on blue] {mfa_number} [/]"
            f" in your authenticator app[/]\n"
        )
    else:
        console.print("[yellow]No MFA number was shown. Approve the MFA push notification on your phone if one appears...[/]")

    # Wait for approval and redirect
    console.print("[dim]Waiting for MFA approval (up to 3 minutes)...[/]")
    try:
        page.wait_for_url("**/outlook.office.com/**", timeout=180000)
        page.wait_for_timeout(8000)
    except Exception as e:
        log.debug("Best-effort operation failed", exc_info=True)
        console.print(f"[yellow]Timeout waiting for Outlook: {e}[/]")


def _extract_storage_tokens(page, captured: dict[str, str]) -> None:
    """Fallback: extract MSAL tokens from browser storage."""
    try:
        storage_tokens = page.evaluate("""() => {
            const results = [];
            for (const store of [sessionStorage, localStorage]) {
                for (let i = 0; i < store.length; i++) {
                    const key = store.key(i);
                    try {
                        const parsed = JSON.parse(store.getItem(key));
                        if (parsed && parsed.secret && parsed.credentialType === 'AccessToken') {
                            results.push({
                                secret: parsed.secret,
                                target: parsed.target || '',
                            });
                        }
                    } catch {}
                }
            }
            return results;
        }""")
        if storage_tokens:
            console.print(f"[dim]  Found {len(storage_tokens)} token(s) in browser storage[/]")
            for t in storage_tokens:
                target = t.get("target", "")
                secret = t.get("secret", "")
                for domain in KNOWN_DOMAINS:
                    if domain in target:
                        captured[domain] = secret
                        break
                else:
                    if secret and "graph.microsoft.com" not in captured:
                        captured["graph.microsoft.com"] = secret
    except Exception:
        log.debug("Best-effort operation failed", exc_info=True)
        pass


def _wait_for_login_or_tokens(page, captured: dict[str, str], timeout_ms: int = 20000) -> bool:
    """Wait until the SSO redirect settles. Return True if we land on a login page.

    A fixed short wait raced the redirect: if it was slow, the login steps were
    skipped and no tokens were captured.
    """
    waited = 0
    while waited < timeout_ms:
        if "login.microsoft" in page.url:
            return True
        if captured:
            return False
        page.wait_for_timeout(500)
        waited += 500
    return "login.microsoft" in page.url


def capture_tokens_via_browser(headless: bool = False) -> dict[str, str]:
    """Open Outlook Web in Playwright and intercept all bearer tokens.

    Args:
        headless: If True, run fully headless (no visible browser window).
                  Requires MS_EMAIL and MS_PASSWORD in .env.

    Returns a dict of {domain: token} for each Microsoft API domain seen.
    """
    if headless:
        if not config.MS_EMAIL or not config.MS_PASSWORD:
            console.print("[red]Headless mode requires MS_EMAIL and MS_PASSWORD in .env[/]")
            return {}
        console.print("[cyan]Starting headless login...[/]")
    else:
        console.print("[cyan]Opening browser for Microsoft login...[/]")

    captured = {}

    with sync_playwright() as p:
        if headless:
            browser = p.chromium.launch(headless=True)
            context = browser.new_context(viewport={"width": 1280, "height": 720})
        else:
            browser = p.chromium.launch(headless=False, args=["--start-maximized"])
            context = browser.new_context(no_viewport=True)

        page = context.new_page()
        page.on("request", _make_request_interceptor(captured))

        # Navigate to Outlook calendar (triggers SSO + lots of API calls)
        page.goto("https://outlook.office.com/calendar", wait_until="domcontentloaded")
        on_login_page = _wait_for_login_or_tokens(page, captured)

        if on_login_page and config.MS_EMAIL:
            _enter_email(page, config.MS_EMAIL)

        if on_login_page and config.MS_PASSWORD:
            _enter_password(page, config.MS_PASSWORD)

        if on_login_page:
            _wait_for_mfa(page)

        # If no tokens captured yet, navigate around to trigger API calls
        if not captured:
            console.print("[dim]Navigating to capture tokens...[/]")
            try:
                page.goto(
                    "https://outlook.office.com/calendar/view/week",
                    wait_until="domcontentloaded",
                )
                page.wait_for_timeout(8000)
            except Exception:
                log.debug("Best-effort operation failed", exc_info=True)
                pass

        # Fallback: extract MSAL tokens from browser storage
        if not captured:
            _extract_storage_tokens(page, captured)

        if captured:
            copilot_token = _capture_copilot_in_context(context)
            if copilot_token:
                captured[config.COPILOT_TOKEN_DOMAIN] = copilot_token
                console.print(f"[dim]  Captured token for {config.COPILOT_TOKEN_DOMAIN}[/]")
            else:
                console.print("[yellow]No Microsoft 365 Copilot token captured. 'copilot' commands will not work.[/]")

        # Save browser session state so Playwright can reuse it without MFA
        storage_path = config.SESSION_DIR / "browser_state.json"
        try:
            context.storage_state(path=str(storage_path))
            console.print(f"[dim]Browser session saved to {storage_path}[/]")
        except Exception as e:
            log.debug("Best-effort operation failed", exc_info=True)
            console.print(f"[dim]Could not save browser session: {e}[/]")

        context.close()
        browser.close()

    if captured:
        _save_tokens(captured)
        console.print(f"[green]Captured tokens for: {', '.join(captured.keys())}[/]")
    else:
        console.print("[red]Could not capture any bearer tokens.[/]")

    return captured


def _click_login_interstitial(page) -> None:
    for selector in LOGIN_INTERSTITIAL_SELECTORS:
        try:
            button = page.locator(selector).first
            if button.count() and button.is_visible(timeout=500):
                button.click()
                page.wait_for_timeout(2000)
                return
        except Exception:
            log.debug("Best-effort operation failed", exc_info=True)


def _capture_copilot_in_context(context, timeout_ms: int = 30000) -> str | None:
    """Open the Copilot web app in a signed-in context and capture its token.

    The app sends the token in request headers to substrate.office.com and in
    the `access_token` parameter of its StreamHub WebSocket URL.
    """
    found: list[str] = []

    def consider(token: str) -> None:
        if _token_claims(token).get("aud") == COPILOT_AUDIENCE:
            found.append(token)

    def on_request(request) -> None:
        auth = request.headers.get("authorization", "")
        if auth.startswith("Bearer "):
            consider(auth.removeprefix("Bearer "))

    def on_websocket(websocket) -> None:
        for token in parse_qs(urlsplit(websocket.url).query).get("access_token", []):
            consider(token)

    page = context.new_page()
    page.on("request", on_request)
    page.on("websocket", on_websocket)
    try:
        page.goto(COPILOT_CHAT_URL, wait_until="domcontentloaded", timeout=60000)
        waited = 0
        while not found and waited < timeout_ms:
            page.wait_for_timeout(1000)
            waited += 1000
            if "login.microsoftonline.com" in page.url:
                _click_login_interstitial(page)
    except Exception:
        log.debug("Copilot token capture failed", exc_info=True)
    finally:
        page.close()
    if not found:
        return None
    return max(found, key=lambda token: _token_claims(token).get("exp", 0))


def _capture_copilot_with_saved_session(headless: bool) -> str | None:
    storage_path = config.SESSION_DIR / "browser_state.json"
    if not storage_path.exists():
        return None
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=headless)
        context = browser.new_context(storage_state=str(storage_path), viewport={"width": 1280, "height": 900})
        token = _capture_copilot_in_context(context)
        if token:
            try:
                context.storage_state(path=str(storage_path))
            except Exception:
                log.debug("Best-effort operation failed", exc_info=True)
        context.close()
        browser.close()
    return token


def capture_copilot_token(headless: bool = True) -> str | None:
    """Refresh only the Copilot token from the saved browser session.

    No password or MFA step runs. If the saved session has expired, this
    returns None and the caller must run the full login. The capture runs in
    a worker thread because the Playwright sync API fails inside an asyncio
    event loop, such as the MCP server loop.
    """
    with ThreadPoolExecutor(max_workers=1) as pool:
        token = pool.submit(_capture_copilot_with_saved_session, headless).result()
    if token:
        tokens = _load_tokens()
        tokens[config.COPILOT_TOKEN_DOMAIN] = token
        _save_tokens(tokens)
    return token


def get_tokens(headless: bool = False) -> dict[str, str]:
    """Get valid tokens, using cache or opening browser.

    Returns a dict of {domain: token}.
    """
    tokens = _load_tokens()
    if tokens:
        test_domain = (
            "graph.microsoft.com" if "graph.microsoft.com" in tokens
            else next(iter(tokens), None)
        )
        if test_domain:
            test_url = (
                f"https://{test_domain}/v1.0/me" if "graph" in test_domain
                else f"https://{test_domain}/api/v2.0/me"
            )
            if _test_token(tokens[test_domain], test_url):
                console.print(f"[green]Using saved tokens ({len(tokens)} domain(s))[/]")
                return tokens
            console.print("[yellow]Saved tokens expired, re-authenticating...[/]")

    return capture_tokens_via_browser(headless=headless)


def get_graph_token() -> str | None:
    """Convenience: get just the Graph API token, or None."""
    tokens = get_tokens()
    return tokens.get("graph.microsoft.com")


if __name__ == "__main__":
    import sys

    headless = "--headless" in sys.argv
    force = "--force" in sys.argv

    if force:
        capture_tokens_via_browser(headless=headless)
    else:
        get_tokens(headless=headless)
