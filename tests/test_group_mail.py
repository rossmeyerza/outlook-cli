"""Regression checks for group inbox paging and monitoring."""
import json
from argparse import Namespace
from unittest.mock import Mock

import pytest

from outlook_draft import group_mail
from outlook_draft.commands import groups
from outlook_draft.errors import OutlookAPIError, OutlookDraftError

GROUP = "01234567-89ab-cdef-0123-456789abcdef"
PATH = f"/groups/{GROUP}/threads"
NEXT = f"https://graph.microsoft.com/v1.0{PATH}?$skip=2"
AFTER = "2026-11-18T14:00:00Z"


def response(*items, next_link=None):
    page = {"value": list(items)}
    if next_link:
        page["@odata.nextLink"] = next_link
    return Mock(**{"json.return_value": page})


def test_recent_posts_follows_pages_and_reads_replies():
    recent = {"id": "reply", "receivedDateTime": AFTER, "from": {"emailAddress": {"address": "a@example.test"}}}
    prior = {"id": "first", "receivedDateTime": "2026-11-17T15:00:00Z"}
    client = Mock()
    client._request.side_effect = [
        response({"id": "quiet", "lastDeliveredDateTime": prior["receivedDateTime"]}, next_link=NEXT),
        response({"id": "old/==", "topic": "Old subject", "lastDeliveredDateTime": AFTER}),
        response(prior, recent),
    ]

    result = group_mail.recent_posts(client, GROUP, group_mail.parse_timestamp(AFTER))
    assert result == [{
        "groupId": GROUP,
        "threadId": "old/==",
        "postId": "reply",
        "receivedDateTime": "2026-11-18T14:00:00+00:00",
        "subject": "Old subject",
        "sender": "a@example.test",
    }]
    requests = client._request.call_args_list
    assert requests[1].args == ("GET", NEXT)
    assert requests[1].kwargs == {"params": None}
    assert requests[2].args == ("GET", PATH + "/old%2F%3D%3D/posts")
    assert "body" not in requests[2].kwargs["params"]["$select"]


def test_group_pagination_rejects_foreign_origin():
    client = Mock()
    client._request.return_value = response(next_link="https://example.test/v1.0/groups")
    with pytest.raises(OutlookDraftError, match="invalid group inbox page"):
        list(group_mail.threads(client, GROUP))
    assert client._request.call_count == 1


def test_watch_retries_and_emits_each_post_once(monkeypatch, capsys):
    after = group_mail.parse_timestamp(AFTER)
    event = {
        "groupId": GROUP, "threadId": "thread", "postId": "first",
        "receivedDateTime": after.isoformat(), "subject": "Subject", "sender": "",
    }
    reply = {**event, "postId": "reply"}
    read = Mock(side_effect=[OutlookAPIError(429), [event], [event, reply]])
    monkeypatch.setattr(group_mail, "recent_posts", read)
    monkeypatch.setattr(groups, "datetime", Mock(**{"now.return_value": after}))
    monkeypatch.setattr(groups.time, "sleep", Mock(side_effect=[None, None, KeyboardInterrupt]))
    monkeypatch.setattr(groups.runtime, "error_console", Mock())
    client = Mock()
    monkeypatch.setattr(groups.runtime, "get_graph_client", lambda: client)

    groups._run(Namespace(command="watch", group_id=GROUP, after=after, poll=60, json=True))

    assert [json.loads(line) for line in capsys.readouterr().out.splitlines()] == [event, reply]
    assert all(call.args == (client, GROUP, after) for call in read.call_args_list)
    client.close.assert_called_once()
