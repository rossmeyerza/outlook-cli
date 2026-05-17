from __future__ import annotations

import os
import uuid
from datetime import datetime, timedelta

import pytest

from outlook_draft.calendar_time import local_timezone, outlook_datetime
from outlook_draft.errors import OutlookAPIError
from outlook_draft.outlook_client import OutlookClient
from outlook_draft.token_manager import TokenManager


pytestmark = pytest.mark.live


def _require_live() -> None:
    if os.environ.get("OUTLOOK_CLI_LIVE") != "1":
        pytest.skip("Set OUTLOOK_CLI_LIVE=1 to run live Outlook calendar tests")


def _test_window() -> tuple[str, str]:
    tz = local_timezone()
    start = datetime.now(tz).replace(hour=16, minute=37, second=0, microsecond=0) + timedelta(days=1)
    if start <= datetime.now(tz):
        start += timedelta(days=1)
    end = start + timedelta(minutes=10)
    return outlook_datetime(start), outlook_datetime(end)


def test_live_calendar_create_get_delete_roundtrip() -> None:
    """Create, fetch, delete, and verify deletion of a disposable Outlook event.

    This is intentionally gated because it mutates the real calendar.
    """
    _require_live()
    subject = f"outlook-cli live test safe to delete {uuid.uuid4()}"
    start, end = _test_window()
    client = OutlookClient(TokenManager())
    event_id = ""

    try:
        created = client.create_event(
            subject=subject,
            start_dt=start,
            end_dt=end,
            body="Created by the outlook-draft-cli live test. Safe to delete.",
        )
        event_id = created.get("Id") or created.get("id") or ""
        assert event_id, f"Created event did not include an ID: {created!r}"

        fetched = client.get_event(event_id)
        assert fetched.get("Subject") == subject
        assert fetched.get("Start", {}).get("DateTime", "").startswith(start)
        assert fetched.get("End", {}).get("DateTime", "").startswith(end)

        client.delete_event(event_id)
        event_id = ""

        with pytest.raises(OutlookAPIError):
            client.get_event(created.get("Id") or created.get("id"))
    finally:
        if event_id:
            try:
                client.delete_event(event_id)
            except OutlookAPIError:
                pass
        client.close()
