"""Read organisation relationships from the Microsoft Graph directory."""
from __future__ import annotations

import uuid
from typing import Any
from urllib.parse import quote, urlsplit

from .errors import OutlookAPIError

PROFILE_SELECT = "id,displayName,mail,userPrincipalName,jobTitle,department,companyName"
JsonDict = dict[str, Any]


def validate_org_options(user: str, include_chain: bool, max_depth: int | None) -> str:
    """Validate options before authentication or a directory request."""
    if not isinstance(user, str) or any(ord(char) < 32 or ord(char) == 127 for char in user):
        raise ValueError("The directory reference must be text without control characters.")
    user = user.strip()
    if not user or user in {".", ".."} or user.isdecimal():
        raise ValueError("Use an email, UPN, directory user ID, unique name or me, not a cached contact index.")
    if type(include_chain) is not bool:
        raise ValueError("include_chain must be a boolean")
    if max_depth is not None and (type(max_depth) is not int or max_depth < 1):
        raise ValueError("max_depth must be a positive integer")
    if max_depth is not None and not include_chain:
        raise ValueError("max_depth requires include_chain=true (--chain)")
    return user


def _entries(page: JsonDict) -> list[JsonDict]:
    if not isinstance(page, dict):
        raise ValueError("The directory returned an invalid page.")
    items = page.get("value")
    if not isinstance(items, list) or any(not isinstance(item, dict) or not isinstance(item.get("id"), str) or not item["id"] for item in items):
        raise ValueError("The directory returned an invalid collection.")
    return items


def _profile(person: JsonDict) -> JsonDict:
    if not isinstance(person, dict) or not isinstance(person.get("id"), str) or not person["id"]:
        raise ValueError("The directory returned a profile without a valid ID.")
    return person


class DirectoryMixin:
    """Graph-only read methods for user profiles and organisation relationships."""

    def get_organisation_profile(self, user: str) -> JsonDict:
        user = validate_org_options(user, False, None)
        path = "/me" if user.casefold() == "me" else f"/users/{quote(user, safe='')}"
        return _profile(self._request("GET", path, params={"$select": PROFILE_SELECT}).json())

    def find_organisation_users(self, query: str, field: str = "displayName", prefix: bool = False) -> JsonDict:
        """Read at most two candidates to detect ambiguity, not to list users."""
        if field not in {"displayName", "mail"}:
            raise ValueError("Unsupported directory search field")
        escaped = query.replace("'", "''")
        expression = f"startswith({field},'{escaped}')" if prefix else f"{field} eq '{escaped}'"
        page = self._request("GET", "/users", params={"$filter": expression, "$select": PROFILE_SELECT, "$top": "2"}).json()
        _entries(page)
        return page

    def get_organisation_manager(self, user_id: str) -> JsonDict | None:
        user_id = validate_org_options(user_id, False, None)
        try:
            response = self._request("GET", f"/users/{quote(user_id, safe='')}/manager", params={"$select": PROFILE_SELECT})
        except OutlookAPIError as exc:
            if exc.status == 404:
                return None
            raise
        return _profile(response.json())

    def _organisation_page_url(self, url: str, path: str) -> str:
        expected, actual = urlsplit(self._base_url + path), urlsplit(url)
        if (actual.scheme != "https" or actual.netloc.lower() != expected.netloc.lower()
                or actual.path != expected.path or actual.fragment or actual.username or actual.password):
            raise ValueError("Directory pagination points outside the selected Graph relationship.")
        return url

    def list_organisation_reports(self, user_id: str) -> JsonDict:
        """Read every page of direct reports. Preserve partial results after an API error."""
        user_id = validate_org_options(user_id, False, None)
        path = f"/users/{quote(user_id, safe='')}/directReports"
        url, params = self._base_url + path, {"$select": PROFILE_SELECT, "$top": "100"}
        seen_urls, seen_ids, items = set(), set(), []
        duplicate = False
        while url:
            self._organisation_page_url(url, path)
            if url in seen_urls:
                raise ValueError("The directory repeated a page URL. The report list is incomplete.")
            seen_urls.add(url)
            try:
                page = self._request("GET", url, params=params).json()
            except OutlookAPIError as exc:
                if exc.status == 401:
                    raise
                return {"items": items, "complete": False, "error": f"Direct reports unavailable or incomplete (HTTP {exc.status})."}
            for item in _entries(page):
                if item["id"] in seen_ids:
                    duplicate = True
                    continue
                seen_ids.add(item["id"])
                items.append(item)
            url, params = page.get("@odata.nextLink"), None
        warning = "The directory repeated a report across pages; the result may reflect concurrent changes." if duplicate else None
        return {"items": items, "complete": not duplicate, "error": warning}


def _is_user_id(value: str) -> bool:
    try:
        uuid.UUID(value)
    except ValueError:
        return False
    return True


def resolve_org_user(client, user: str) -> JsonDict:
    """Resolve a directory user without an arbitrary choice between name matches."""
    if user.casefold() == "me" or "@" in user or _is_user_id(user):
        try:
            return client.get_organisation_profile(user)
        except OutlookAPIError as exc:
            if exc.status != 404:
                raise
        if "@" not in user:
            raise ValueError("No directory user matches this ID.")
        page = client.find_organisation_users(user, field="mail")
    else:
        page = client.find_organisation_users(user)
        if not _entries(page) and not page.get("@odata.nextLink"):
            page = client.find_organisation_users(user, prefix=True)
    people = _entries(page)
    if not people:
        raise ValueError("No directory user matches. Use their company email address, not a personal contact ID.")
    if len(people) != 1 or page.get("@odata.nextLink"):
        candidates = "; ".join(f"{person.get('displayName') or person['id']} <{person.get('mail') or person.get('userPrincipalName') or person['id']}>" for person in people)
        raise ValueError(f"More than one directory user may match: {candidates}. Use an email or directory user ID.")
    return people[0]


def org_person(person: JsonDict) -> JsonDict:
    """Keep missing fields visible instead of guessing organisation details."""
    _profile(person)
    return {"id": person["id"], "displayName": person.get("displayName"),
            "email": person.get("mail") or person.get("userPrincipalName"),
            "userPrincipalName": person.get("userPrincipalName"), "jobTitle": person.get("jobTitle"),
            "department": person.get("department"), "company": person.get("companyName"),
            "objectType": person.get("@odata.type", "#microsoft.graph.user"),
            "limitedInformation": not bool(person.get("displayName") or person.get("mail") or person.get("userPrincipalName"))}


def _manager(client, user_id: str, warnings: list[str]) -> tuple[JsonDict | None, str]:
    try:
        manager = client.get_organisation_manager(user_id)
    except OutlookAPIError as exc:
        if exc.status == 401:
            raise
        warnings.append(f"Manager details unavailable (HTTP {exc.status}).")
        return None, "unavailable"
    return manager, "available" if manager else "notSetOrNotVisible"


def _management_chain(client, person: JsonDict, manager: JsonDict | None, status: str,
                      max_depth: int | None, warnings: list[str]) -> tuple[list[JsonDict], bool]:
    chain, seen = [], {person["id"]}
    while manager:
        if manager["id"] in seen:
            warnings.append("The directory contains a management cycle. The chain is incomplete.")
            return chain, False
        seen.add(manager["id"])
        shaped = org_person(manager)
        chain.append(shaped)
        if shaped["limitedInformation"] or shaped["objectType"] != "#microsoft.graph.user":
            warnings.append("Further managers are unavailable for a restricted profile or organisational contact.")
            return chain, False
        if max_depth is not None and len(chain) >= max_depth:
            warnings.append("The requested management depth was reached. Further managers were not checked.")
            return chain, False
        manager, status = _manager(client, manager["id"], warnings)
    return chain, status != "unavailable"


def organisation_view(client, user: str = "me", include_chain: bool = False, max_depth: int | None = None) -> JsonDict:
    """Read the available directory structure, not an inferred company structure."""
    user = validate_org_options(user, include_chain, max_depth)
    person = resolve_org_user(client, user)
    warnings: list[str] = []
    manager, manager_status = _manager(client, person["id"], warnings)
    reports = client.list_organisation_reports(person["id"])
    if reports["error"]:
        warnings.append(reports["error"])
    shaped_reports = [org_person(item) for item in reports["items"]]
    shaped_reports.sort(key=lambda item: (item["displayName"] or item["email"] or item["id"]).casefold())
    if any(item["limitedInformation"] for item in shaped_reports) or (manager and org_person(manager)["limitedInformation"]):
        warnings.append("Some directory entries have limited information. Missing fields are not proof that no organisation detail exists.")
    result = {"source": "Microsoft Graph directory", "person": org_person(person),
              "manager": org_person(manager) if manager else None, "managerStatus": manager_status,
              "directReports": shaped_reports, "directReportsCount": len(shaped_reports),
              "directReportsComplete": reports["complete"], "managementChain": None, "managementChainComplete": None,
              "warnings": warnings}
    if include_chain:
        result["managementChain"], result["managementChainComplete"] = _management_chain(client, person, manager, manager_status, max_depth, warnings)
    return result
