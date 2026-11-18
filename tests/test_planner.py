"""Planner plans, buckets and tasks: shared logic, Graph calls and MCP tools."""
from __future__ import annotations

import asyncio
from typing import Any

import pytest

from outlook_draft import planner
from outlook_draft.errors import OutlookAPIError
from outlook_draft.mcp import planner as mcp_planner
from outlook_draft.mcp.server import build_server
from tests.mcp_fakes import FakeClient, use_client

PLAN = "PlanId12345678"
BUCKET = "BucketId123456"
TASK = "TaskId1234567890"
USER = "7c77e651-e4b1-465d-af37-8279274d5a5c"
OTHER = "e43713f8-6197-4af3-9f90-f8c1911978d3"


class Response:
    def __init__(self, body: Any) -> None:
        self.body = body

    def json(self) -> Any:
        return self.body


class RecordingClient(planner.PlannerMixin):
    """Planner mixin over a scripted _request. Records every call."""

    def __init__(self, responses: dict[tuple[str, str], Any]) -> None:
        self.responses = responses
        self.calls: list[tuple[str, str, dict]] = []

    def _request(self, method: str, path: str, **kwargs: Any) -> Response:
        self.calls.append((method, path, kwargs))
        return Response(self.responses.get((method, path), {}))


def test_task_changes_maps_and_validates_fields():
    changes = planner.task_changes(title=" Brief ", bucket_id=BUCKET, due="2026-11-20 14:00",
                                   priority="urgent", percent_complete=50)
    assert changes == {"title": "Brief", "bucketId": BUCKET, "dueDateTime": "2026-11-20T14:00:00Z",
                       "priority": 1, "percentComplete": 50}
    assert planner.task_changes(clear_due=True) == {"dueDateTime": None}
    for bad in ({"title": "  "}, {"priority": "soon"}, {"percent_complete": 101}, {"percent_complete": True},
                {"due": "2026-11-20", "clear_due": True}, {"bucket_id": "3"}):
        with pytest.raises(ValueError):
            planner.task_changes(**bad)


def test_task_summary_names_status_and_priority():
    summary = planner.task_summary({"id": "t", "percentComplete": 100, "priority": 3,
                                    "assignments": {USER: {}}, "dueDateTime": None})
    assert (summary["status"], summary["priority"], summary["assignees"], summary["due"]) == (
        "Completed", "important", [USER], "")
    assert planner.task_summary({})["status"] == "NotStarted"


def test_create_task_previews_then_creates():
    fake = FakeClient(resolve_planner_user=lambda ref: {"id": USER, "displayName": "Ross"},
                      create_planner_task={"id": "new", "title": "Brief", "planId": PLAN})
    preview = planner.create_task(fake, plan_id=PLAN, bucket_id=BUCKET, title="Brief", assign=["ross@example.com"])
    assert preview["created"] is False
    assert "create_planner_task" not in fake.names()
    assert preview["task"]["assignments"][USER]["orderHint"] == " !"

    done = planner.create_task(fake, plan_id=PLAN, bucket_id=BUCKET, title="Brief", description="Notes", confirm=True)
    assert done["created"] is True and done["task"]["id"] == "new"
    assert fake.call("create_planner_task")[0][1] == "Notes"


def test_update_task_previews_then_patches_and_unassigns_with_null():
    fake = FakeClient(
        get_planner_task={"id": TASK, "title": "Old"},
        resolve_planner_user=lambda ref: {"id": USER if ref == "a@example.com" else OTHER, "displayName": ""},
        update_planner_task={"id": TASK, "title": "New", "percentComplete": 100},
    )
    preview = planner.update_task(fake, TASK, title="New", assign=["a@example.com"], unassign=["b@example.com"])
    assert preview["updated"] is False and "update_planner_task" not in fake.names()
    assignments = preview["changes"]["assignments"]
    assert assignments[USER]["orderHint"] == " !" and assignments[OTHER] is None

    done = planner.update_task(fake, TASK, percent_complete=100, confirm=True)
    assert done["updated"] is True and done["task"]["status"] == "Completed"
    with pytest.raises(ValueError):
        planner.update_task(fake, TASK)
    with pytest.raises(ValueError):
        planner.update_task(fake, "3", title="x")


def test_delete_task_and_bucket_preview_unless_confirmed():
    fake = FakeClient(get_planner_task={"id": TASK, "title": "Old"}, create_planner_bucket={"id": "b1"})
    assert planner.delete_task(fake, TASK)["deleted"] is False
    assert "delete_planner_task" not in fake.names()
    assert planner.delete_task(fake, TASK, confirm=True)["deleted"] is True
    assert planner.create_bucket(fake, PLAN, " Scoping ")["created"] is False
    assert "create_planner_bucket" not in fake.names()
    made = planner.create_bucket(fake, PLAN, " Scoping ", confirm=True)
    assert made == {"planId": PLAN, "name": "Scoping", "created": True, "id": "b1"}
    with pytest.raises(ValueError):
        planner.create_bucket(fake, PLAN, "  ")


def test_update_sends_the_current_etag_and_never_retries():
    client = RecordingClient({("GET", f"/planner/tasks/{TASK}"): {"@odata.etag": 'W/"t1"', "id": TASK},
                              ("GET", f"/planner/tasks/{TASK}/details"): {"@odata.etag": 'W/"d1"'}})
    client.update_planner_task(TASK, {"title": "New"}, description="Notes")
    patches = [(path, kw) for method, path, kw in client.calls if method == "PATCH"]
    assert patches[0][0] == f"/planner/tasks/{TASK}"
    assert patches[0][1]["extra_headers"]["If-Match"] == 'W/"t1"'
    assert patches[1][0] == f"/planner/tasks/{TASK}/details"
    assert patches[1][1]["extra_headers"]["If-Match"] == 'W/"d1"'
    assert patches[1][1]["json_body"] == {"description": "Notes"}
    assert all(kw.get("max_retries") == 0 for m, _, kw in client.calls if m == "PATCH")


def test_delete_sends_the_etag_and_create_never_retries():
    client = RecordingClient({("GET", f"/planner/tasks/{TASK}"): {"@odata.etag": 'W/"t1"'},
                              ("POST", "/planner/tasks"): {"id": TASK}})
    client.delete_planner_task(TASK)
    delete = next(kw for m, _, kw in client.calls if m == "DELETE")
    assert delete["extra_headers"]["If-Match"] == 'W/"t1"' and delete["max_retries"] == 0
    client.create_planner_task({"planId": PLAN, "title": "x"})
    assert next(kw for m, _, kw in client.calls if m == "POST")["max_retries"] == 0


def test_description_failure_names_the_created_task():
    class Failing(RecordingClient):
        def set_planner_description(self, task_id: str, description: str) -> None:
            raise OutlookAPIError(412, "etag changed")

    client = Failing({("POST", "/planner/tasks"): {"id": TASK}})
    with pytest.raises(OutlookAPIError, match=TASK):
        client.create_planner_task({"title": "x"}, "Notes")


def test_lists_follow_paging_and_pick_the_right_path():
    client = RecordingClient({("GET", "/planner/plans/" + PLAN + "/tasks"): {
        "value": [{"id": "1"}, {"id": "2"}], "@odata.nextLink": "https://graph.microsoft.com/next"},
        ("GET", "https://graph.microsoft.com/next"): {"value": [{"id": "3"}]}})
    assert [t["id"] for t in client.list_planner_tasks(plan_id=PLAN, top=3)] == ["1", "2", "3"]
    client.list_planner_tasks(bucket_id=BUCKET)
    client.list_planner_tasks()
    client.list_planner_plans(group_id="g1")
    paths = [path for _, path, _ in client.calls]
    assert f"/planner/buckets/{BUCKET}/tasks" in paths and "/me/planner/tasks" in paths
    assert "/groups/g1/planner/plans" in paths


def test_resolve_user_accepts_guid_and_email_and_rejects_names():
    client = RecordingClient({("GET", "/users/ross@example.com"): {"id": USER, "displayName": "Ross"}})
    assert client.resolve_planner_user(USER.upper()) == {"id": USER, "displayName": ""}
    assert client.resolve_planner_user("ross@example.com") == {"id": USER, "displayName": "Ross"}
    with pytest.raises(ValueError):
        client.resolve_planner_user("Ross Meyer")


def test_mcp_tools_use_the_graph_client(monkeypatch):
    fake = FakeClient(list_planner_buckets=[{"id": BUCKET, "name": "Scoping", "planId": PLAN}],
                      get_planner_task={"id": TASK, "title": "T", "percentComplete": 0},
                      get_planner_task_details={"description": "Notes", "checklist": {"a": {"title": "Do", "isChecked": True}}})
    use_client(monkeypatch, mcp_planner, "graph_client", fake)
    assert mcp_planner.planner_buckets(PLAN) == [{"id": BUCKET, "name": "Scoping", "planId": PLAN}]
    got = mcp_planner.planner_task_get(TASK)
    assert got["description"] == "Notes" and got["checklist"] == [{"title": "Do", "checked": True}]
    with pytest.raises(ValueError):
        mcp_planner.planner_buckets("1")
    assert mcp_planner.planner_task_delete(TASK)["deleted"] is False


WRITE_TOOLS = {"planner_task_create", "planner_task_update", "planner_task_delete", "planner_bucket_create",
               "planner_plan_pin", "planner_bucket_update", "planner_bucket_delete", "planner_member_add",
               "planner_member_remove", "planner_plan_unpin"}


def test_write_tools_default_to_preview_and_are_flagged_destructive():
    tools = {t.name: t for t in asyncio.run(build_server().list_tools())}
    for name in WRITE_TOOLS:
        assert tools[name].inputSchema["properties"]["confirm"]["default"] is False
        assert tools[name].annotations.destructiveHint is True
    assert {"planner_plans", "planner_groups", "planner_buckets", "planner_tasks", "planner_task_get"} <= set(tools)


def test_readonly_policy_hides_planner_write_tools(monkeypatch):
    monkeypatch.setenv("OUTLOOK_CLI_READONLY", "1")
    names = {t.name for t in asyncio.run(build_server().list_tools())}
    assert not names & WRITE_TOOLS
    assert {"planner_plans", "planner_tasks"} <= names


def test_groups_keep_only_microsoft_365_groups_filter_and_sort():
    groups = [{"id": "1", "displayName": "Zeta", "groupTypes": ["Unified"]},
              {"id": "2", "displayName": "Security only", "groupTypes": []},
              {"id": "3", "displayName": "Alpha triage", "groupTypes": ["Unified"]}]
    client = RecordingClient({("GET", "/me/memberOf/microsoft.graph.group"): {"value": groups}})
    assert [g["id"] for g in client.list_planner_groups()] == ["3", "1"]
    assert [g["id"] for g in client.list_planner_groups(search="TRIAGE")] == ["3"]
    assert client.calls[0][2]["params"]["$select"].startswith("id,displayName")


def test_pin_sends_only_the_one_plan_with_the_etag_and_never_retries():
    existing = {"old": {"@odata.type": "#microsoft.graph.plannerFavoritePlanReference", "orderHint": "000", "planTitle": "Old"}}
    client = RecordingClient({("PATCH", planner.BETA_PLANNER_USER): {"favoritePlanReferences": {}}})
    client.pin_planner_plan(PLAN, "Triage test", {"@odata.etag": 'W/"u1"', "favoritePlanReferences": existing})
    method, path, kwargs = client.calls[0]
    assert (method, path) == ("PATCH", planner.BETA_PLANNER_USER)
    assert set(kwargs["json_body"]["favoritePlanReferences"]) == {PLAN}
    assert kwargs["extra_headers"]["If-Match"] == 'W/"u1"' and kwargs["max_retries"] == 0


def test_pin_plan_previews_then_pins_and_skips_a_plan_that_is_already_pinned():
    user = {"@odata.etag": "e", "favoritePlanReferences": {}}
    fake = FakeClient(get_planner_plan={"id": PLAN, "title": "Triage test"}, get_planner_user=user)
    preview = planner.pin_plan(fake, PLAN)
    assert preview == {"plan": {"id": PLAN, "title": "Triage test"}, "alreadyPinned": False, "pinned": False}
    assert "pin_planner_plan" not in fake.names()
    assert planner.pin_plan(fake, PLAN, confirm=True)["pinned"] is True
    assert fake.call("pin_planner_plan")[0] == (PLAN, "Triage test", user)

    again = FakeClient(get_planner_plan={"title": "T"}, get_planner_user={"favoritePlanReferences": {PLAN: {}}})
    assert planner.pin_plan(again, PLAN, confirm=True) == {
        "plan": {"id": PLAN, "title": "T"}, "alreadyPinned": True, "pinned": True}
    assert "pin_planner_plan" not in again.names()
    with pytest.raises(ValueError):
        planner.pin_plan(again, "3")


BUCKETS = [{"id": "BucketAAAAAA", "name": "A", "orderHint": "c", "planId": PLAN},
           {"id": "BucketBBBBBB", "name": "B", "orderHint": "b", "planId": PLAN},
           {"id": "BucketCCCCCC", "name": "C", "orderHint": "a", "planId": PLAN}]  # board order: A, B, C


def _bucket_fake(**extra):
    return FakeClient(get_planner_bucket=lambda bucket_id, *a: next(b for b in BUCKETS if b["id"] == bucket_id),
                      list_planner_buckets=[BUCKETS[2], BUCKETS[0], BUCKETS[1]], **extra)


def test_board_order_is_the_reverse_of_the_service_order():
    assert [b["name"] for b in planner.board_order([BUCKETS[2], BUCKETS[0], BUCKETS[1]])] == ["A", "B", "C"]


def test_move_hints_place_a_bucket_first_in_the_middle_and_last_on_the_board():
    fake = _bucket_fake()
    assert planner.update_bucket(fake, "BucketCCCCCC", first=True)["changes"]["orderHint"] == "c !"
    assert planner.update_bucket(fake, "BucketCCCCCC", after="BucketAAAAAA")["changes"]["orderHint"] == "b c!"
    assert planner.update_bucket(fake, "BucketAAAAAA", after="BucketCCCCCC")["changes"]["orderHint"] == " a!"
    assert planner.update_bucket(fake, "BucketAAAAAA", after="BucketBBBBBB")["orderAfter"] == ["B", "A", "C"]
    assert planner.update_bucket(fake, "BucketCCCCCC", first=True)["orderAfter"] == ["C", "A", "B"]


def test_rename_and_move_preview_then_apply():
    fake = _bucket_fake(update_planner_bucket={"id": "BucketBBBBBB", "name": "Renamed"})
    preview = planner.update_bucket(fake, "BucketBBBBBB", name=" Renamed ")
    assert preview["changes"] == {"name": "Renamed"} and preview["orderAfter"] == ["A", "Renamed", "C"]
    assert preview["updated"] is False and "update_planner_bucket" not in fake.names()
    done = planner.update_bucket(fake, "BucketBBBBBB", name="Renamed", first=True, confirm=True)
    assert done["updated"] is True
    assert fake.call("update_planner_bucket")[0] == ("BucketBBBBBB", {"name": "Renamed", "orderHint": "c !"})


def test_bucket_update_rejects_bad_input():
    fake = _bucket_fake()
    for kwargs in ({}, {"name": "  "}, {"after": "BucketBBBBBB", "first": True}, {"after": "BucketAAAAAA"}, {"after": "BucketMissing"}):
        with pytest.raises(ValueError):
            planner.update_bucket(fake, "BucketAAAAAA", **kwargs)
    with pytest.raises(ValueError):
        planner.update_bucket(fake, "3", name="x")


def test_bucket_delete_refuses_a_bucket_that_holds_tasks():
    fake = _bucket_fake(list_planner_tasks=[{"title": "One"}, {"title": "Two"}])
    preview = planner.delete_bucket(fake, "BucketAAAAAA")
    assert (preview["tasksInBucket"], preview["taskTitles"], preview["blocked"]) == (2, ["One", "Two"], True)
    with pytest.raises(ValueError, match="orphaned"):
        planner.delete_bucket(fake, "BucketAAAAAA", confirm=True)
    assert "delete_planner_bucket" not in fake.names()


def test_bucket_delete_removes_an_empty_bucket_only_when_confirmed():
    fake = _bucket_fake(list_planner_tasks=[])
    assert planner.delete_bucket(fake, "BucketAAAAAA")["deleted"] is False
    assert "delete_planner_bucket" not in fake.names()
    assert planner.delete_bucket(fake, "BucketAAAAAA", confirm=True)["deleted"] is True


def test_bucket_update_and_delete_use_the_etag_and_never_retry():
    client = RecordingClient({("GET", f"/planner/buckets/{BUCKET}"): {"@odata.etag": 'W/"b1"', "name": "x"}})
    client.update_planner_bucket(BUCKET, {"name": "y"})
    client.delete_planner_bucket(BUCKET)
    for method in ("PATCH", "DELETE"):
        kwargs = next(kw for m, _, kw in client.calls if m == method)
        assert kwargs["extra_headers"]["If-Match"] == 'W/"b1"' and kwargs["max_retries"] == 0


def test_unpin_sends_only_a_null_for_that_plan_with_the_etag_and_never_retries():
    pins = {PLAN: {"orderHint": "000", "planTitle": "A"}, "other": {"orderHint": "000", "planTitle": "B"}}
    client = RecordingClient({})
    client.unpin_planner_plan(PLAN, {"@odata.etag": 'W/"u1"', "favoritePlanReferences": pins})
    method, path, kwargs = client.calls[0]
    assert (method, path) == ("PATCH", planner.BETA_PLANNER_USER)
    assert kwargs["json_body"] == {"favoritePlanReferences": {PLAN: None}}
    assert kwargs["extra_headers"]["If-Match"] == 'W/"u1"' and kwargs["max_retries"] == 0


MEMBERS = [{"id": USER, "displayName": "Ross", "mail": "ross@example.com"}]


def _member_fake(**extra):
    return FakeClient(get_planner_plan={"id": PLAN, "title": "Triage", "owner": "group-1"}, list_planner_members=MEMBERS,
                      resolve_planner_user=lambda ref: {"id": USER if "ross" in ref else OTHER, "displayName": ref}, **extra)


def test_list_members_reads_the_owning_group():
    result = planner.list_members(_member_fake(), PLAN)
    assert result["groupId"] == "group-1" and result["members"][0]["mail"] == "ross@example.com"
    with pytest.raises(ValueError):
        planner.list_members(FakeClient(get_planner_plan={"title": "Roster plan"}), PLAN)


def test_member_add_previews_skips_existing_members_and_adds_only_new_people():
    fake = _member_fake()
    preview = planner.add_members(fake, PLAN, ["ross@example.com", "ann@example.com", "ann@example.com"])
    assert [p["id"] for p in preview["toAdd"]] == [OTHER] and [p["id"] for p in preview["alreadyMembers"]] == [USER]
    assert preview["applied"] is False and "add_planner_member" not in fake.names()
    done = planner.add_members(fake, PLAN, ["ross@example.com", "ann@example.com"], confirm=True)
    assert done["applied"] is True and [p["id"] for p in done["added"]] == [OTHER]
    assert fake.call("add_planner_member")[0] == ("group-1", OTHER)
    with pytest.raises(ValueError):
        planner.add_members(fake, PLAN, [])


def test_member_add_names_who_was_added_before_a_failure():
    def add(group_id, user_id):
        if user_id == OTHER:
            raise OutlookAPIError(403, "not an owner")

    fake = _member_fake(add_planner_member=add)
    with pytest.raises(OutlookAPIError, match="Added nobody before failing on ann@example.com"):
        planner.add_members(fake, PLAN, ["ann@example.com"], confirm=True)


def test_member_add_posts_a_reference_and_never_retries():
    client = RecordingClient({})
    client.add_planner_member("group-1", USER)
    method, path, kwargs = client.calls[0]
    assert (method, path) == ("POST", "/groups/group-1/members/$ref")
    assert kwargs["json_body"]["@odata.id"].endswith("/directoryObjects/" + USER) and kwargs["max_retries"] == 0


TWO_MEMBERS = [{"id": USER, "displayName": "Ross", "mail": "ross@example.com"},
               {"id": OTHER, "displayName": "Ann", "mail": "ann@example.com"}]


def _remove_fake(members=None, owners=None, **extra):
    return FakeClient(
        get_planner_plan={"id": PLAN, "title": "Triage", "owner": "group-1"},
        list_planner_members=TWO_MEMBERS if members is None else members,
        list_planner_owners=[{"id": USER}] if owners is None else owners,
        list_planner_tasks=[{"assignments": {OTHER: {}}}, {"assignments": {OTHER: {}, USER: {}}}, {"assignments": {}}],
        resolve_planner_user=lambda ref: {"id": USER if "ross" in ref else OTHER, "displayName": ref}, **extra)


def test_member_remove_previews_counts_assigned_tasks_then_removes():
    fake = _remove_fake()
    preview = planner.remove_members(fake, PLAN, ["ann@example.com", "nobody-new@example.com"])
    assert [(p["id"], p["assignedTasks"]) for p in preview["toRemove"]] == [(OTHER, 2)]
    assert preview["applied"] is False and "remove_planner_member" not in fake.names()
    done = planner.remove_members(fake, PLAN, ["ann@example.com"], confirm=True)
    assert done["applied"] is True and [p["id"] for p in done["removed"]] == [OTHER]
    assert fake.call("remove_planner_member")[0] == ("group-1", OTHER)


def test_member_remove_lists_people_who_are_not_members_and_skips_them():
    fake = _remove_fake(members=[TWO_MEMBERS[0]], owners=[{"id": USER}])
    result = planner.remove_members(fake, PLAN, ["ann@example.com"], confirm=True)
    assert result["toRemove"] == [] and [p["id"] for p in result["notMembers"]] == [OTHER]
    assert "remove_planner_member" not in fake.names()


def test_member_remove_refuses_owners_the_last_member_and_empty_input():
    with pytest.raises(ValueError, match="owns the group"):
        planner.remove_members(_remove_fake(), PLAN, ["ross@example.com"])
    with pytest.raises(ValueError, match="every member"):
        planner.remove_members(_remove_fake(owners=[]), PLAN, ["ross@example.com", "ann@example.com"])
    with pytest.raises(ValueError):
        planner.remove_members(_remove_fake(), PLAN, [])
    with pytest.raises(ValueError):
        planner.remove_members(_remove_fake(), "3", ["ann@example.com"])


def test_member_remove_names_who_was_removed_before_a_failure():
    def remove(group_id, user_id):
        raise OutlookAPIError(403, "not an owner")

    with pytest.raises(OutlookAPIError, match="Removed nobody before failing on ann@example.com"):
        planner.remove_members(_remove_fake(remove_planner_member=remove), PLAN, ["ann@example.com"], confirm=True)


def test_member_remove_deletes_the_reference_and_never_retries():
    client = RecordingClient({})
    client.remove_planner_member("group-1", USER)
    method, path, kwargs = client.calls[0]
    assert (method, path) == ("DELETE", f"/groups/group-1/members/{USER}/$ref") and kwargs["max_retries"] == 0


def test_pin_and_unpin_send_no_other_pin_so_the_service_limit_is_not_hit():
    client = RecordingClient({})
    many = {f"p{i}": {"orderHint": "000"} for i in range(5)}
    client.pin_planner_plan(PLAN, "T", {"@odata.etag": "e", "favoritePlanReferences": many})
    client.unpin_planner_plan("p1", {"@odata.etag": "e", "favoritePlanReferences": many})
    assert [len(c[2]["json_body"]["favoritePlanReferences"]) for c in client.calls] == [1, 1]


def test_unpin_previews_then_unpins_and_works_for_a_plan_that_no_longer_exists():
    user = {"@odata.etag": "e", "favoritePlanReferences": {PLAN: {"planTitle": "Gone plan"}}}
    fake = FakeClient(get_planner_user=user)
    preview = planner.unpin_plan(fake, PLAN)
    assert preview == {"plan": {"id": PLAN, "title": "Gone plan"}, "wasPinned": True, "unpinned": False}
    assert "unpin_planner_plan" not in fake.names() and "get_planner_plan" not in fake.names()
    assert planner.unpin_plan(fake, PLAN, confirm=True)["unpinned"] is True
    assert fake.call("unpin_planner_plan")[0] == (PLAN, user)

    none = FakeClient(get_planner_user={"favoritePlanReferences": {}})
    assert planner.unpin_plan(none, PLAN, confirm=True)["wasPinned"] is False and "unpin_planner_plan" not in none.names()
    with pytest.raises(ValueError):
        planner.unpin_plan(none, "3")
