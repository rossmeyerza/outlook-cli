"""Full IDs must resolve without endpoint escapes or a partial-list dependency."""
from __future__ import annotations

from unittest.mock import Mock

import pytest
from mcp_fakes import FakeClient, use_client

from outlook_draft.errors import OutlookAPIError
from outlook_draft.mcp import calendar, contacts, mail, tasks


@pytest.mark.parametrize("event_id", [".", ".."])
@pytest.mark.parametrize("method, field, changes", [
    (calendar.cal_delete, "event_id", {}),
    (calendar.cal_cancel, "event_id", {}),
    (calendar.cal_update, "event_id", {"subject": "Changed"}),
    (contacts.contact_update, "contact_id", {"company": "Company"}),
    (tasks.task_update, "task_id", {"subject": "Changed"}),
    (tasks.task_delete, "task_id", {}),
])
def test_dot_segments_never_reach_authentication(monkeypatch, event_id, method, field, changes):
    for module in (calendar, contacts, tasks):
        monkeypatch.setattr(module, "outlook_client", Mock(side_effect=AssertionError("No authentication")))
    with pytest.raises(ValueError):
        method(**{field: event_id}, **changes, confirm=True)


def test_mail_move_accepts_an_id_outside_the_first_folder_page(monkeypatch):
    fake = use_client(monkeypatch, mail, "outlook_client", FakeClient(list_mail_folders=[],
        get_mail_folder={"Id": "folder-id", "DisplayName": "Target"}, move_message={"Id": "moved-id"}))
    assert mail.mail_move("message-id", "folder-id")["id"] == "moved-id"
    assert fake.call("move_message")[0] == ("message-id", "folder-id")


@pytest.mark.parametrize("status", [403, 404])
def test_failed_folder_lookup_never_moves_mail(monkeypatch, status):
    def fail(*args):
        raise OutlookAPIError(status, "No access")
    fake = use_client(monkeypatch, mail, "outlook_client", FakeClient(list_mail_folders=[], get_mail_folder=fail))
    with pytest.raises(OutlookAPIError if status == 403 else ValueError):
        mail.mail_move("message-id", "folder-id")
    assert "move_message" not in fake.names()
