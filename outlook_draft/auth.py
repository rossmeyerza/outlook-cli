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
import re
from datetime import datetime

import requests
from playwright.sync_api import sync_playwright
from rich.console import Console

from . import config

console = Console()

config.ensure_dirs()
SESSION_DIR = config.SESSION_DIR
TOKENS_PATH = config.TOKENS_FILE
MS_EMAIL = config.MS_EMAIL
MS_PASSWORD = config.MS_PASSWORD

# Microsoft API domains we look for when intercepting requests.
KNOWN_DOMAINS = [
    "graph.microsoft.com",
    "outlook.office365.com",
    "outlook.office.com",
    "substrate.office.com",
]


def _token_scopes(token: str) -> set[str]:
    """Decode JWT scopes without verification."""
    try:
        payload = token.split(".")[1]
        payload += "=" * (4 - len(payload) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload))
        return set(claims.get("scp", "").split())
    except Exception:
        return set()


def _save_tokens(tokens: dict[str, str]):
    """Save all captured tokens keyed by domain."""
    config.ensure_dirs()
    data = {
        "tokens": tokens,
        "captured_at": datetime.now().isoformat(),
    }
    TOKENS_PATH.write_text(json.dumps(data, indent=2))


def _load_tokens() -> dict[str, str]:
    """Load saved tokens from disk."""
    if not TOKENS_PATH.exists():
        return {}
    try:
        data = json.loads(TOKENS_PATH.read_text())
        return data.get("tokens", {})
    except Exception:
        return {}


def _test_token(token: str, url: str) -> bool:
    """Quick check if a token works against a URL."""
    try:
        resp = requests.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            timeout=10,
        )
        return resp.status_code == 200
    except Exception:
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
                if domain not in captured:
                    captured[domain] = new_token
                    console.print(f"[dim]  Captured token for {domain}[/]")
                else:
                    new_scopes = _token_scopes(new_token)
                    old_scopes = _token_scopes(captured[domain])
                    if len(new_scopes) > len(old_scopes):
                        captured[domain] = new_token
                        console.print(f"[dim]  Updated token for {domain} (broader scopes)[/]")
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
        console.print(f"[red]Failed to enter email: {e}[/]")
        return False


def _body_text(page) -> str:
    try:
        value = page.locator("body").inner_text(timeout=2000)
        return value or ""
    except Exception:
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
        pass


def capture_tokens_via_browser(headless: bool = False) -> dict[str, str]:
    """Open Outlook Web in Playwright and intercept all bearer tokens.

    Args:
        headless: If True, run fully headless (no visible browser window).
                  Requires MS_EMAIL and MS_PASSWORD in .env.

    Returns a dict of {domain: token} for each Microsoft API domain seen.
    """
    if headless:
        if not MS_EMAIL or not MS_PASSWORD:
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
        page.wait_for_timeout(3000)

        # Check if we hit the login page
        current_url = page.url
        on_login_page = (
            "login.microsoftonline.com" in current_url
            or "login.microsoft" in current_url
        )

        if on_login_page and MS_EMAIL:
            _enter_email(page, MS_EMAIL)

        if on_login_page and MS_PASSWORD:
            _enter_password(page, MS_PASSWORD)

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
                pass

        # Fallback: extract MSAL tokens from browser storage
        if not captured:
            _extract_storage_tokens(page, captured)

        # Save browser session state so Playwright can reuse it without MFA
        storage_path = SESSION_DIR / "browser_state.json"
        try:
            context.storage_state(path=str(storage_path))
            console.print(f"[dim]Browser session saved to {storage_path}[/]")
        except Exception as e:
            console.print(f"[dim]Could not save browser session: {e}[/]")

        context.close()
        browser.close()

    if captured:
        _save_tokens(captured)
        console.print(f"[green]Captured tokens for: {', '.join(captured.keys())}[/]")
    else:
        console.print("[red]Could not capture any bearer tokens.[/]")

    return captured


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
