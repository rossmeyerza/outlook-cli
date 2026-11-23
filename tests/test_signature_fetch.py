from unittest.mock import Mock

import pytest
from playwright.sync_api import TimeoutError

from outlook_draft.commands import signature
from outlook_draft.errors import CliError

OUTLOOK = "https://outlook.cloud.microsoft/mail/"
LOGIN = "https://login.microsoftonline.com/common/oauth2/authorize"
OKTA_HOME = "https://example.okta.com/app/UserHome?session_hint=AUTHENTICATED"


def _redirect_page(*urls):
    page = Mock()
    redirects = iter(urls)

    def finish_redirect(_timeout):
        page.url = next(redirects)

    page.wait_for_timeout.side_effect = finish_redirect
    return page


@pytest.mark.parametrize("redirects", [(OUTLOOK,), (OKTA_HOME, OUTLOOK)])
def test_signature_fetch_opens_outlook_with_saved_session(redirects):
    page = _redirect_page(*redirects)

    signature._open_outlook(page)

    assert page.goto.call_count == len(redirects)
    assert all(call.args == (OUTLOOK,) for call in page.goto.call_args_list)
    page.locator.assert_not_called()


@pytest.mark.parametrize("redirects", [(LOGIN,), (OKTA_HOME, LOGIN), (OKTA_HOME, OKTA_HOME)])
def test_signature_fetch_stops_when_user_action_is_required(redirects):
    page = _redirect_page(*redirects)

    with pytest.raises(CliError, match="outlook-cli auth"):
        signature._open_outlook(page)

    assert page.goto.call_count == len(redirects)
    page.locator.assert_not_called()


def test_signature_fetch_hides_navigation_error_details():
    page = Mock(url=LOGIN)
    page.goto.side_effect = TimeoutError("private navigation details")

    with pytest.raises(CliError, match="outlook-cli auth") as error:
        signature._open_outlook(page)

    assert "private navigation details" not in str(error.value)


def test_signature_fetch_requires_saved_state(tmp_path, monkeypatch):
    monkeypatch.setattr(signature.config, "SESSION_DIR", tmp_path)
    monkeypatch.setattr(signature.config, "ensure_dirs", lambda: None)

    with pytest.raises(CliError, match="outlook-cli auth"):
        signature._capture_signatures(headless=True)


def _settings_response(path, body, *, host="outlook.cloud.microsoft"):
    response = Mock(url=f"https://{host}/ows/v1/OutlookCloudSettings/settings{path}")
    response.json.return_value = body
    return response


def test_signature_capture_matches_current_assignments_not_response_order():
    html, assignments = {}, {}
    capture = signature._capture_response(html, assignments)
    capture(_settings_response("/", [
        {"name": "Roaming_New_Signature", "type": "String", "value": "Example new"},
        {"name": "roaming_signature_list", "type": "BlobArray", "value": "Example list"},
        {"name": "Roaming_Reply_Signature", "type": "String", "value": "Example reply"},
    ]))
    for label, value in [("Example reply", "<p>Reply</p>"), ("Unused", "<p>Unused</p>"), ("Example new", "<p>New</p>")]:
        capture(_settings_response("/account?settingname=encoded", [
            {"name": "encoded", "nameBase64Encoded": True, "secondaryKey": label,
             "parentSetting": "roaming_signature_list", "type": "String", "value": value},
            {"name": "plain", "parentSetting": "roaming_signature_list", "type": "String", "value": "Plain text"},
        ]))

    assert signature._assigned_signatures(html, assignments) == {"new": "<p>New</p>", "reply": "<p>Reply</p>"}


def test_signature_capture_accepts_legacy_query_name_and_ignores_other_hosts():
    html, assignments = {}, {}
    capture = signature._capture_response(html, assignments)
    capture(_settings_response("/account?settingname=Legacy", {"type": "String", "value": "<p>Legacy</p>"},
                               host="outlook.office.com"))
    capture(_settings_response("/", {"name": "roaming_new_signature", "type": "String", "value": "Legacy"}))
    unrelated = _settings_response("/account?settingname=Legacy", {}, host="example.com")
    capture(unrelated)

    unrelated.json.assert_not_called()
    assert signature._assigned_signatures(html, assignments) == {"new": "<p>Legacy</p>"}
