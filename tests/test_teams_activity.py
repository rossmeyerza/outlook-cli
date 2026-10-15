"""Regression coverage for stale Graph chat update timestamps."""
from __future__ import annotations

import argparse
import io
from unittest.mock import Mock

import httpx
import pytest
from mcp_fakes import FakeClient, use_client
from rich.console import Console

from outlook_draft.commands import teams as cli_teams
from outlook_draft.errors import OutlookAPIError
from outlook_draft.mcp import teams
from outlook_draft.outlook_client import OutlookClient
from outlook_draft.teams_activity import TeamsChatList, activity_fields, activity_time


def chat(chat_id="person", timestamp="2026-10-15T14:53:00Z", updated="2026-08-13T12:22:25Z"):
    return {
        "id": chat_id, "chatType": "oneOnOne", "lastUpdatedDateTime": updated,
        "members": [{"displayName": "Vagelis Gkarilas", "email": "vagelis@example.test"}],
        "lastMessagePreview": {"createdDateTime": timestamp, "messageType": "message",
                               "from": {"user": {"displayName": "Vagelis Gkarilas"}},
                               "body": {"contentType": "html", "content": "<p>yo! outside the cost\nof the MBA</p>"}},
    }


def test_mcp_defaults_use_message_time_not_chat_update(monkeypatch):
    person = chat()
    old = chat("group", "2026-10-01T10:00:00Z", "2026-10-15T15:00:00Z")
    old["chatType"] = "group"
    fake = FakeClient(list_teams_chats=[old, person], list_teams_messages=[],
                      get_teams_chat=person, list_teams_chat_members=person["members"])
    use_client(monkeypatch, teams, "graph_client", fake)
    for results in (teams.teams_list(), teams.teams_search("Vagelis")):
        assert [item["id"] for item in results] == ["person", "group"]
        assert results[0]["lastMessageTime"] == "2026-10-15T14:53:00Z"
        assert results[0]["lastUpdated"] == "2026-08-13T12:22:25Z"
        assert results[0]["lastSender"] == "Vagelis Gkarilas"
        assert results[0]["preview"] == "yo! outside the cost of the MBA"
    assert "list_teams_chat_members" not in fake.names()
    results = teams.teams_search("Vagelis", one_on_one=True, include_messages=True)
    assert len(results) == 1 and results[0]["messages"] == []


@pytest.mark.parametrize("matches,other,has_more,count,incomplete", [
    (1, 0, False, 1, False),  # Exact scan and result limits, but nothing omitted.
    (1, 3, False, 3, False),
    (1, 3, False, 1, False),  # More candidates are not proof of more matches.
    (2, 0, False, 1, True),
    (2, 0, False, 2, False),
    (1, 0, True, 3, True),  # Graph has an unvisited page.
    (0, 2, True, 3, True),
    (0, 2, False, 3, False),
])
def test_scan_and_count_limits_require_evidence(monkeypatch, caplog, matches, other, has_more, count, incomplete):
    chats = [chat(str(index)) for index in range(matches)]
    for index in range(other):
        unmatched = chat(f"other-{index}")
        unmatched["members"] = [{"displayName": "Someone Else"}]
        chats.append(unmatched)
    chats = TeamsChatList(chats, has_more=has_more)
    use_client(monkeypatch, teams, "graph_client", FakeClient(list_teams_chats=chats))
    results = teams.teams_search("Vagelis", scan=len(chats), count=count, one_on_one=True)
    assert len(results) == min(matches, count)
    assert all(result["searchIncomplete"] is incomplete for result in results)
    assert ("truncated" in caplog.text) is incomplete

    client = Mock()
    client.list_teams_chats.return_value = chats
    monkeypatch.setattr(cli_teams.runtime, "get_graph_client", lambda: client)
    output = io.StringIO()
    monkeypatch.setattr(cli_teams.runtime, "console", Console(file=output))
    render, save = Mock(), Mock()
    monkeypatch.setattr(cli_teams, "_render_search_results", render)
    monkeypatch.setattr(cli_teams, "save_cache", save)
    args = argparse.Namespace(query=["Vagelis"], scan=len(chats), count=count, member_count=100,
                              no_spinner=True, one_on_one=True)
    cli_teams.cmd_teams_search(args)
    assert ("truncated" in output.getvalue()) is incomplete
    assert len(render.call_args.args[1]) == min(matches, count)
    assert len(save.call_args.args[1]) == min(matches, count)


def test_cli_search_does_not_resort_by_update_time(monkeypatch, tmp_path):
    client = Mock()
    client.list_teams_chats.return_value = [chat("old", "2026-10-01T10:00:00Z", "2026-10-15T15:00:00Z"), chat()]
    monkeypatch.setattr(cli_teams.runtime, "get_graph_client", lambda: client)
    monkeypatch.setattr(cli_teams, "TEAMS_CACHE", tmp_path / "cache.json")
    render = Mock()
    monkeypatch.setattr(cli_teams, "_render_search_results", render)
    args = argparse.Namespace(query=["Vagelis"], scan=200, count=20, member_count=100, no_spinner=True)
    cli_teams.cmd_teams_search(args)
    assert [row[0]["id"] for row in render.call_args.args[1]] == ["person", "old"]
    client.list_teams_chat_members.assert_not_called()


def test_graph_expands_previews_and_members_across_pages():
    client = OutlookClient(Mock())
    person, old = chat(), chat("old", "2026-10-01T10:00:00Z")
    client._request = Mock(side_effect=[
        httpx.Response(200, json={"value": [person], "@odata.nextLink": "https://graph.microsoft.com/v1.0/next"}),
        httpx.Response(200, json={"value": [old]}),
    ])
    results = client.list_teams_chats(top=2)
    assert results == [person, old]
    assert results.has_more is False
    calls = client._request.call_args_list
    assert calls[0].kwargs["params"]["$expand"] == "lastMessagePreview,members"
    assert calls[0].kwargs["params"]["$orderby"] == "lastMessagePreview/createdDateTime desc"
    assert calls[1].kwargs["params"] is None
    assert len(calls) == 2  # No per-chat message or member reads.


@pytest.mark.parametrize("size,next_link,top,has_more", [
    (1, None, 1, False),
    (1, "https://graph.microsoft.com/v1.0/next", 1, True),
    (2, None, 1, True),  # A returned page can extend beyond the requested cap.
    (0, None, 1, False),
])
def test_graph_chat_limit_tracks_remaining_data(size, next_link, top, has_more):
    client = OutlookClient(Mock())
    values = [chat(str(index)) for index in range(size)]
    client._request = Mock(return_value=httpx.Response(200, json={"value": values, "@odata.nextLink": next_link}))
    results = client.list_teams_chats(top)
    assert len(results) == min(top, size)
    assert isinstance(results, list)
    assert results.has_more is has_more
    client._request.assert_called_once()


def test_missing_preview_falls_back_only_for_missing_chat(caplog):
    client = OutlookClient(Mock())
    client._request = Mock(return_value=httpx.Response(200, json={"value": [chat(), {"id": "missing"}]}))
    client.list_teams_messages = Mock(return_value=[{"createdDateTime": "2026-10-15T16:00:00Z"}])
    results = client.list_teams_chats(2)
    assert results[0]["id"] == "missing"
    client.list_teams_messages.assert_called_once_with("missing", top=1)
    client.list_teams_messages.side_effect = OutlookAPIError(403, "denied")
    assert len(client.list_teams_chats(2)) == 2
    assert "Cannot read Teams activity" in caplog.text


@pytest.mark.parametrize("width", [80, 100, 180])
def test_search_table_fits_terminal_and_prints_stable_ids(monkeypatch, width):
    output = io.StringIO()
    monkeypatch.setattr(cli_teams.runtime, "console", Console(file=output, width=width, color_system=None))
    person = chat("19:stable-chat-id@unq.gbl.spaces")
    cli_teams._render_search_results("Vagelis", [(person, person["members"], ["member"])], 1, 0)
    text = output.getvalue()
    assert "Vagelis" in text and "Oct 15 14:53" in text
    assert "yo! outside" in text and person["id"] in text
    assert all(len(line) <= width for line in text.splitlines())


def test_empty_and_system_previews_do_not_invent_message_times():
    assert activity_time({"lastUpdatedDateTime": "stale"}) == ""
    fields = activity_fields({"lastMessagePreview": {"messageType": "systemEventMessage"}})
    assert fields["preview"] == "[system event]" and fields["lastSender"] == "System"
    fields = activity_fields({"lastMessagePreview": {"attachments": [{"id": "card"}]}})
    assert fields["preview"] == "[attachment or rich content]"


@pytest.mark.parametrize("message", [
    {"messageType": "unknownFutureValue", "eventDetail": {"@odata.type": "membersAdded"}},
    {"messageType": "unknownFutureValue", "body": {"contentType": "html", "content": "<systemEventMessage/>"}},
])
def test_graph_unknown_event_types_are_labelled(message, monkeypatch):
    assert activity_fields({"lastMessagePreview": message})["preview"] == "[system event]"
    assert cli_teams._teams_body(argparse.Namespace(), message) == "[system event]"
    assert teams._message_text(message) == "[system event]"
    assert cli_teams._teams_sender(message) == "System"
    output = io.StringIO()
    monkeypatch.setattr(cli_teams.runtime, "console", Console(file=output, width=100))
    person = chat()
    person["lastMessagePreview"] = message
    cli_teams._render_search_results("Vagelis", [(person, person["members"], [])], 1, 0)
    assert "[system event]" in output.getvalue()
