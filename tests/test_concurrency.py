from __future__ import annotations

from threading import Barrier, Event
from unittest.mock import Mock

from outlook_draft.concurrency import ordered_map
from outlook_draft.errors import OutlookAPIError
from outlook_draft.mcp import teams
from outlook_draft.outlook_client import OutlookClient


def test_ordered_map_keeps_input_order_when_workers_finish_in_reverse():
    events = [Event() for _ in range(8)]
    finished = []

    def worker(index):
        if index < 7:
            assert events[index + 1].wait(5)
        finished.append(index)
        events[index].set()
        return index

    with ordered_map(worker, range(8)) as results:
        assert list(results) == list(range(8))
    assert finished == list(reversed(range(8)))


def test_ordered_map_stops_after_one_batch_on_early_match():
    called = []
    with ordered_map(lambda index: called.append(index) or index, range(100)) as results:
        assert next(results) == 0
    assert sorted(called) == list(range(8))


def test_client_initialization_is_shared_across_workers(monkeypatch):
    created = []
    barrier = Barrier(8)

    def factory(**kwargs):
        created.append(kwargs)
        return Mock(is_closed=False)

    monkeypatch.setattr("outlook_draft.outlook_client.httpx.Client", factory)
    client = OutlookClient(Mock())

    def worker(index):
        barrier.wait(timeout=5)
        return client._ensure_client()

    with ordered_map(worker, range(8)) as results:
        clients = list(results)
    assert len(created) == 1
    assert all(value is clients[0] for value in clients)


def test_mcp_list_and_search_preserve_stable_order_and_member_failures(monkeypatch):
    from contextlib import contextmanager

    class Client:
        def list_teams_chats(self, top):
            return [{"id": str(index), "lastUpdatedDateTime": "same"} for index in range(8)]

        def list_teams_chat_members(self, chat_id, top):
            if chat_id == "3":
                raise OutlookAPIError(403, "forbidden")
            return [{"displayName": "Match " + chat_id, "email": "other@example.test"}]

    @contextmanager
    def factory():
        yield Client()

    monkeypatch.setattr(teams, "graph_client", factory)
    assert [chat["id"] for chat in teams.teams_list(8)] == list(map(str, range(8)))
    assert [chat["id"] for chat in teams.teams_search("Match", count=8)] == ["0", "1", "2", "4", "5", "6", "7"]
