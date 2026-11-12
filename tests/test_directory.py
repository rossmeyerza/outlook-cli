"""Organisation reads use fake Graph responses and no live directory data."""
from __future__ import annotations

from unittest.mock import Mock

import httpx
import pytest
from mcp_fakes import FakeClient

from outlook_draft import config
from outlook_draft.directory import PROFILE_SELECT, organisation_view, resolve_org_user, validate_org_options
from outlook_draft.errors import OutlookAPIError
from outlook_draft.outlook_client import OutlookClient

PERSON = {"id": "person-id", "displayName": "Person", "mail": "person@example.test", "jobTitle": "Consultant",
          "department": "Department", "companyName": "Company"}
MANAGER = {**PERSON, "id": "manager-id", "displayName": "Manager", "mail": "manager@example.test"}
REPORT = {**PERSON, "id": "report-id", "displayName": "Report", "mail": "report@example.test"}
BASE = config.GRAPH_BASE_URL.rstrip("/")


def client_for(*responses):
    client = OutlookClient(Mock(), base_url=BASE)
    client._request = Mock(side_effect=responses)
    return client


def response(payload):
    return httpx.Response(200, json=payload)


def fake_view(**extra):
    values = dict(get_organisation_profile=PERSON, get_organisation_manager=MANAGER,
                  list_organisation_reports={"items": [REPORT], "complete": True, "error": None})
    values.update(extra)
    return FakeClient(**values)


@pytest.mark.parametrize("user", ["", " ", ".", "..", "1", "name\n", "\tname", None])
def test_invalid_reference(user):
    with pytest.raises(ValueError):
        validate_org_options(user, False, None)


@pytest.mark.parametrize("chain, depth", [(False, 1), (True, 0), (True, -1), (True, True), (True, 1.5), ("false", None)])
def test_invalid_chain_options(chain, depth):
    with pytest.raises(ValueError):
        validate_org_options("me", chain, depth)


def test_profile_paths_and_select_are_read_only():
    client = client_for(response(PERSON), response(PERSON))
    assert client.get_organisation_profile("me") == PERSON
    assert client.get_organisation_profile("person+tag@example.test") == PERSON
    calls = client._request.call_args_list
    assert calls[0].args == ("GET", "/me")
    assert calls[1].args == ("GET", "/users/person%2Btag%40example.test")
    assert calls[0].kwargs == {"params": {"$select": PROFILE_SELECT}}


def test_name_filters_escape_apostrophes():
    client = client_for(response({"value": []}), response({"value": []}))
    client.find_organisation_users("O'Brien")
    client.find_organisation_users("O'Brien", prefix=True)
    queries = [call.kwargs["params"] for call in client._request.call_args_list]
    assert queries[0]["$filter"] == "displayName eq 'O''Brien'"
    assert queries[1]["$filter"] == "startswith(displayName,'O''Brien')"
    assert all(query["$top"] == "2" for query in queries)


def test_unique_exact_name_and_prefix_resolution():
    exact = FakeClient(find_organisation_users={"value": [PERSON]})
    assert resolve_org_user(exact, "Person")["id"] == PERSON["id"]
    def find(query, **kwargs):
        return {"value": [PERSON] if kwargs.get("prefix") else []}
    prefix = FakeClient(find_organisation_users=find)
    assert resolve_org_user(prefix, "Per")["id"] == PERSON["id"]
    assert len(prefix.calls) == 2


@pytest.mark.parametrize("people, next_link", [([PERSON, MANAGER], None), ([PERSON], "more")])
def test_ambiguous_or_incomplete_name_never_selects_a_person(people, next_link):
    client = FakeClient(find_organisation_users={"value": people, "@odata.nextLink": next_link})
    with pytest.raises(ValueError, match="More than one"):
        resolve_org_user(client, "Person")
    assert "get_organisation_manager" not in client.names()


def test_primary_email_fallback_and_missing_person():
    def missing(*args):
        raise OutlookAPIError(404, "Missing")
    client = FakeClient(get_organisation_profile=missing, find_organisation_users={"value": [PERSON]})
    assert resolve_org_user(client, "person@example.test")["id"] == PERSON["id"]
    assert client.call("find_organisation_users")[1] == {"field": "mail"}
    client.returns["find_organisation_users"] = {"value": []}
    with pytest.raises(ValueError, match="No directory user"):
        resolve_org_user(client, "person@example.test")


@pytest.mark.parametrize("status", [401, 403, 503])
def test_profile_errors_do_not_start_an_alternative_lookup(status):
    def fail(*args):
        raise OutlookAPIError(status, "Unavailable")
    client = FakeClient(get_organisation_profile=fail)
    with pytest.raises(OutlookAPIError):
        resolve_org_user(client, "person@example.test")
    assert "find_organisation_users" not in client.names()


def test_no_manager_is_distinct_from_forbidden():
    no_manager = client_for(OutlookAPIError(404, "Not recorded"))
    assert no_manager.get_organisation_manager("person-id") is None
    forbidden = client_for(OutlookAPIError(403, "Denied"))
    with pytest.raises(OutlookAPIError):
        forbidden.get_organisation_manager("person-id")


def test_direct_reports_follow_all_pages_without_an_mcp_limit():
    next_link = BASE + "/users/person-id/directReports?$skiptoken=next"
    first = [{**REPORT, "id": f"report-{index}"} for index in range(100)]
    client = client_for(response({"value": first, "@odata.nextLink": next_link}), response({"value": [REPORT]}))
    result = client.list_organisation_reports("person-id")
    assert len(result["items"]) == 101 and result["complete"]
    assert client._request.call_args_list[1].args == ("GET", next_link)
    assert client._request.call_args_list[1].kwargs == {"params": None}


@pytest.mark.parametrize("link", ["https://other.example.test/v1.0/users/person-id/directReports",
    BASE + "/me/messages", "http://graph.microsoft.com/v1.0/users/person-id/directReports",
    BASE + "/users/person-id/directReports#fragment", "/relative/path"])
def test_untrusted_pagination_never_receives_a_token(link):
    client = client_for(response({"value": [REPORT], "@odata.nextLink": link}))
    with pytest.raises(ValueError, match="pagination"):
        client.list_organisation_reports("person-id")
    assert client._request.call_count == 1


def test_repeated_page_cannot_loop_forever():
    url = BASE + "/users/person-id/directReports"
    client = client_for(response({"value": [REPORT], "@odata.nextLink": url}))
    with pytest.raises(ValueError, match="repeated a page"):
        client.list_organisation_reports("person-id")
    assert client._request.call_count == 1


def test_duplicate_reports_do_not_inflate_the_count():
    next_link = BASE + "/users/person-id/directReports?$skiptoken=next"
    client = client_for(response({"value": [REPORT], "@odata.nextLink": next_link}), response({"value": [REPORT]}))
    result = client.list_organisation_reports("person-id")
    assert len(result["items"]) == 1 and not result["complete"]
    assert "repeated a report" in result["error"]


@pytest.mark.parametrize("status", [403, 429, 503])
def test_report_error_keeps_partial_data_and_marks_it_incomplete(status):
    next_link = BASE + "/users/person-id/directReports?$skiptoken=next"
    client = client_for(response({"value": [REPORT], "@odata.nextLink": next_link}), OutlookAPIError(status, "Unavailable"))
    result = client.list_organisation_reports("person-id")
    assert result["items"] == [REPORT] and not result["complete"]
    assert str(status) in result["error"]


def test_expired_token_aborts_instead_of_returning_an_empty_list():
    client = client_for(OutlookAPIError(401, "Expired"))
    with pytest.raises(OutlookAPIError):
        client.list_organisation_reports("person-id")


@pytest.mark.parametrize("page", [{}, {"value": None}, {"value": [{}]}, {"value": [{"id": 1}]}, []])
def test_malformed_directory_pages_are_not_complete_empty_results(page):
    client = client_for(response(page))
    with pytest.raises(ValueError):
        client.list_organisation_reports("person-id")


def test_view_has_person_manager_reports_and_no_implicit_chain():
    client = fake_view()
    result = organisation_view(client)
    assert result["person"]["company"] == "Company"
    assert result["manager"]["id"] == "manager-id"
    assert result["directReportsCount"] == 1 and result["directReportsComplete"]
    assert result["managementChain"] is None and result["managementChainComplete"] is None
    assert client.names().count("get_organisation_manager") == 1


def test_management_chain_stops_at_the_end_of_the_recorded_chain():
    client = fake_view(get_organisation_manager=lambda user: MANAGER if user == "person-id" else None)
    result = organisation_view(client, include_chain=True)
    assert [person["id"] for person in result["managementChain"]] == ["manager-id"]
    assert result["managementChainComplete"]


def test_management_cycle_is_explicit_and_finite():
    client = fake_view(get_organisation_manager=lambda user: MANAGER if user == "person-id" else PERSON)
    result = organisation_view(client, include_chain=True)
    assert not result["managementChainComplete"]
    assert any("cycle" in warning for warning in result["warnings"])
    assert client.names().count("get_organisation_manager") == 2


def test_depth_limit_is_explicit_and_does_not_fetch_further_managers():
    client = fake_view()
    result = organisation_view(client, include_chain=True, max_depth=1)
    assert len(result["managementChain"]) == 1 and not result["managementChainComplete"]
    assert client.names().count("get_organisation_manager") == 1
    assert any("depth" in warning for warning in result["warnings"])


@pytest.mark.parametrize("manager", [{"id": "manager-id"}, {**MANAGER, "@odata.type": "#microsoft.graph.orgContact"}])
def test_restricted_or_contact_manager_does_not_invent_a_chain(manager):
    client = fake_view(get_organisation_manager=manager)
    result = organisation_view(client, include_chain=True)
    assert not result["managementChainComplete"]
    assert client.names().count("get_organisation_manager") == 1


def test_forbidden_manager_is_not_reported_as_no_manager():
    def fail(*args):
        raise OutlookAPIError(403, "Denied")
    result = organisation_view(fake_view(get_organisation_manager=fail), include_chain=True)
    assert result["manager"] is None and result["managerStatus"] == "unavailable"
    assert not result["managementChainComplete"] and result["directReportsComplete"]
    assert any("403" in warning for warning in result["warnings"])


def test_missing_manager_and_no_reports_are_legitimate_directory_results():
    result = organisation_view(fake_view(get_organisation_manager=None,
        list_organisation_reports={"items": [], "complete": True, "error": None}), include_chain=True)
    assert result["managerStatus"] == "notSetOrNotVisible"
    assert result["directReportsCount"] == 0 and result["directReportsComplete"]
    assert result["managementChain"] == [] and result["managementChainComplete"]


def test_limited_reports_retain_their_ids_and_missing_fields():
    result = organisation_view(fake_view(list_organisation_reports={"items": [{"id": "limited-id"}], "complete": True, "error": None}))
    report = result["directReports"][0]
    assert report["id"] == "limited-id" and report["limitedInformation"]
    assert report["displayName"] is None and report["company"] is None
    assert result["warnings"]
