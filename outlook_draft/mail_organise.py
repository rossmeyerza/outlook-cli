"""Mail categories and inbox rules through Microsoft Graph v1.0."""

from __future__ import annotations

import re
from urllib.parse import quote

from .errors import OutlookAPIError
from .mail_urls import validate_mail_url

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
CATEGORY_PATH = "/me/outlook/masterCategories"
RULE_PATH = "/me/mailFolders/inbox/messageRules"
COLOURS = ("none", *(f"preset{i}" for i in range(25)))
SETTINGS_SCOPE = "MailboxSettings.ReadWrite"
MESSAGE_SCOPE = "Mail.ReadWrite"


def _text(value, label: str) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ValueError(f"{label} must be a non-empty string.")
    if any(ord(char) < 32 or ord(char) == 127 for char in value):
        raise ValueError(f"{label} must not contain control characters.")
    return value


def _strings(values, label: str, *, empty: bool = False) -> list[str]:
    if not isinstance(values, list) or (not values and not empty):
        raise ValueError(f"{label} must be a {'possibly empty' if empty else 'non-empty'} list.")
    return list(dict.fromkeys(_text(value, label) for value in values))


def _object(value, allowed: set[str], label: str) -> dict:
    if not isinstance(value, dict):
        raise ValueError(f"{label} must be an object.")
    unknown = set(value) - allowed
    if unknown:
        raise ValueError(f"Unsupported {label} keys: {', '.join(sorted(map(str, unknown)))}")
    return value


def _boolean(value, label: str) -> bool:
    if type(value) is not bool:
        raise ValueError(f"{label} must be a boolean.")
    return value


def _colour(value: str) -> str:
    value = _text(value, "Colour").lower()
    if value not in COLOURS:
        raise ValueError("Colour must be none or preset0 through preset24.")
    return value


def _recipients(value) -> list[dict]:
    if not isinstance(value, list) or not value:
        raise ValueError("fromAddresses must be a non-empty list of recipients.")
    recipients = []
    for recipient in value:
        recipient = _object(recipient, {"emailAddress"}, "recipient")
        address = _object(recipient.get("emailAddress"), {"address", "name"}, "emailAddress")
        email = _text(address.get("address"), "Sender address")
        if not re.fullmatch(r"[^\s@<>]+@[^\s@<>]+\.[^\s@<>]+", email):
            raise ValueError(f"Invalid sender address: {email!r}")
        clean = {"address": email}
        if "name" in address:
            clean["name"] = _text(address["name"], "Sender name")
        recipients.append({"emailAddress": clean})
    return recipients


def _conditions(value) -> dict:
    value = _object(value, {"fromAddresses", "subjectContains"}, "conditions")
    if not value:
        raise ValueError("At least one sender or subject condition is required.")
    result = {}
    for key, item in value.items():
        result[key] = _recipients(item) if key == "fromAddresses" else _strings(item, key)
    return result


def _actions(value) -> dict:
    value = _object(value, {"moveToFolder", "assignCategories", "markAsRead", "stopProcessingRules"}, "actions")
    if not value:
        raise ValueError("At least one supported action is required.")
    result = {}
    for key, item in value.items():
        if key == "moveToFolder":
            result[key] = _text(item, "Folder ID")
        elif key == "assignCategories":
            result[key] = _strings(item, key)
        else:
            result[key] = _boolean(item, key)
    if not any(result.values()):
        raise ValueError("At least one action must have an effect.")
    return result


def validate_rule(payload: dict, *, create: bool = False) -> dict:
    payload = _object(payload, {"displayName", "sequence", "isEnabled", "conditions", "actions"}, "rule")
    required = {"displayName", "sequence", "conditions", "actions"} if create else set()
    if required - set(payload):
        raise ValueError(f"Missing rule fields: {', '.join(sorted(required - set(payload)))}")
    if not payload:
        raise ValueError("Specify at least one rule field.")
    result = dict(payload)
    if "displayName" in result:
        result["displayName"] = _text(result["displayName"], "Rule name")
    if "sequence" in result and (type(result["sequence"]) is not int or not 0 <= result["sequence"] <= 2147483647):
        raise ValueError("Rule sequence must be a non-negative Int32 integer.")
    if "conditions" in result:
        result["conditions"] = _conditions(result["conditions"])
    if "actions" in result:
        result["actions"] = _actions(result["actions"])
    if "isEnabled" in result:
        _boolean(result["isEnabled"], "isEnabled")
    elif create:
        result["isEnabled"] = False
    return result


class MailOrganiseMixin:
    """Require a Graph client and use Graph IDs, not cached CLI references."""

    def _require_organise_graph(self) -> None:
        if self._base_url.rstrip("/") != GRAPH_ROOT:
            raise ValueError("Mail categories and rules require a Microsoft Graph v1.0 client.")

    def _organise_request(self, method: str, path: str, **kwargs):
        self._require_organise_graph()
        try:
            validate_mail_url(path, GRAPH_ROOT)
        except ValueError as exc:
            raise ValueError("Refused an invalid pagination URL for Microsoft Graph v1.0.") from exc
        try:
            return self._request(method, path, **kwargs)
        except OutlookAPIError as exc:
            if exc.status == 403:
                scope = MESSAGE_SCOPE if "/me/messages/" in path else SETTINGS_SCOPE
                if method == "GET":
                    scope = "Mail.Read" if scope == MESSAGE_SCOPE else "MailboxSettings.Read"
                raise OutlookAPIError(403, f"Microsoft Graph denied access. Required scope: {scope}. No API fallback.") from exc
            raise

    def _organise_list(self, path: str) -> list[dict]:
        result, seen = [], set()
        while path:
            if path in seen:
                raise ValueError("Microsoft Graph returned a repeated pagination URL.")
            seen.add(path)
            data = self._organise_request("GET", path).json()
            result.extend(data.get("value", []))
            path = data.get("@odata.nextLink")
            if path and not path.startswith(GRAPH_ROOT + "/"):
                raise ValueError("Refused a pagination URL outside Microsoft Graph v1.0.")
        return result

    def _organise_change(self, method: str, path: str, payload, confirm: bool) -> dict:
        self._require_organise_graph()
        _boolean(confirm, "confirm")
        preview = {"preview": True, "method": method, "path": path, "payload": payload}
        if not confirm:
            return preview
        kwargs = {"max_retries": 0}
        if payload is not None:
            kwargs["json_body"] = payload
        response = self._organise_request(method, path, **kwargs)
        return {"preview": False, "result": response.json() if response.content else None}

    def list_mail_categories(self) -> list[dict]:
        return self._organise_list(CATEGORY_PATH)

    def create_mail_category(self, name: str, colour: str = "none", *, confirm: bool = False) -> dict:
        payload = {"displayName": _text(name, "Category name"), "color": _colour(colour)}
        return self._organise_change("POST", CATEGORY_PATH, payload, confirm)

    def update_mail_category(self, category_id: str, colour: str, *, confirm: bool = False) -> dict:
        path = CATEGORY_PATH + "/" + quote(_text(category_id, "Category ID"), safe="")
        return self._organise_change("PATCH", path, {"color": _colour(colour)}, confirm)

    def delete_mail_category(self, category_id: str, *, confirm: bool = False) -> dict:
        path = CATEGORY_PATH + "/" + quote(_text(category_id, "Category ID"), safe="")
        return self._organise_change("DELETE", path, None, confirm)

    def list_mail_rules(self) -> list[dict]:
        return self._organise_list(RULE_PATH)

    def create_mail_rule(self, payload: dict, *, confirm: bool = False) -> dict:
        return self._organise_change("POST", RULE_PATH, validate_rule(payload, create=True), confirm)

    def update_mail_rule(self, rule_id: str, payload: dict, *, confirm: bool = False) -> dict:
        path = RULE_PATH + "/" + quote(_text(rule_id, "Rule ID"), safe="")
        payload = validate_rule(payload)
        if payload.get("isEnabled") is not False:
            current = self._organise_request("GET", path).json()
            if current.get("isReadOnly") or current.get("hasError"):
                raise ValueError("Cannot update a read-only or invalid rule.")
            if payload.get("isEnabled", current.get("isEnabled", False)):
                if current.get("exceptions"):
                    raise ValueError("Cannot change an active rule with unsupported exceptions.")
                _conditions(payload.get("conditions", current.get("conditions")))
                _actions(payload.get("actions", current.get("actions")))
        return self._organise_change("PATCH", path, payload, confirm)

    def delete_mail_rule(self, rule_id: str, *, confirm: bool = False) -> dict:
        path = RULE_PATH + "/" + quote(_text(rule_id, "Rule ID"), safe="")
        return self._organise_change("DELETE", path, None, confirm)

    def categorise_messages(self, message_ids: list[str], add: list[str] | None = None,
                            remove: list[str] | None = None, *, confirm: bool = False) -> dict:
        """Use full Graph IDs. A failed write stops the batch; never retry it automatically."""
        self._require_organise_graph()
        ids = _strings(message_ids, "Message IDs")
        if any(message_id.isdecimal() for message_id in ids):
            raise ValueError("Use full Graph message IDs, not cached CLI indices.")
        add = _strings([] if add is None else add, "Add categories", empty=True)
        remove = _strings([] if remove is None else remove, "Remove categories", empty=True)
        _boolean(confirm, "confirm")
        if not add and not remove:
            raise ValueError("Specify at least one category to add or remove.")
        if set(add) & set(remove):
            raise ValueError("The same category cannot be added and removed.")
        results, stopped = [], False
        for message_id in ids:
            if stopped:
                results.append({"id": message_id, "status": "unattempted"})
                continue
            result = self._categorise_one(message_id, add, remove, confirm)
            results.append(result)
            stopped = result["status"] == "write_failed"
        return {"preview": not confirm, "results": results,
                "failed": any(item["status"] in {"read_failed", "write_failed"} for item in results),
                "stopped": stopped}

    def _categorise_one(self, message_id: str, add: list[str], remove: list[str], confirm: bool) -> dict:
        path = "/me/messages/" + quote(message_id, safe="")
        result = {"id": message_id}
        try:
            message = self._organise_request("GET", path, params={"$select": "id,categories"}).json()
            before = _strings(message.get("categories", []), "Existing categories", empty=True)
        except (OutlookAPIError, ValueError) as exc:
            return {**result, "status": "read_failed", "error": str(exc)}
        after = list(dict.fromkeys([name for name in before if name not in remove] + add))
        result.update(before=before, categories=after)
        if before == after or not confirm:
            return {**result, "status": "unchanged" if before == after else "preview"}
        headers = {"If-Match": message["@odata.etag"]} if message.get("@odata.etag") else {}
        try:
            self._organise_request("PATCH", path, json_body={"categories": after},
                                   extra_headers=headers, max_retries=0)
        except OutlookAPIError as exc:
            return {**result, "status": "write_failed", "error": str(exc),
                    "uncertain": exc.status == 0 or exc.status >= 500}
        return {**result, "status": "updated"}
