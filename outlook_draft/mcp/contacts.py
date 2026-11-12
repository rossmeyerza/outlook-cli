"""Contact tools."""

from __future__ import annotations

from ..commands.contacts import _contact_id, _type_label
from ..directory import organisation_view, validate_org_options
from .clients import clamp, graph_client, outlook_client
from .shaping import JsonDict
from .validation import require_email, require_id


def contact_search(query: str, count: int = 10) -> list[JsonDict]:
    """Search the org directory and recent contacts by name or email. Read-only."""
    with outlook_client() as client:
        people = client.search_people(query, top=clamp(count))
    return [
        {
            "id": _contact_id(p),
            "displayName": p.get("DisplayName", ""),
            "emails": [
                e.get("Address", "")
                for e in p.get("ScoredEmailAddresses", [])
                if e.get("Address")
            ],
            "type": _type_label(p),
        }
        for p in people
    ]


def _contact(contact: JsonDict) -> JsonDict:
    return {"id": contact.get("Id", ""), "displayName": contact.get("DisplayName", ""),
            "emails": contact.get("EmailAddresses", []), "givenName": contact.get("GivenName", ""),
            "surname": contact.get("Surname", ""), "company": contact.get("CompanyName", ""),
            "mobilePhone": contact.get("MobilePhone1", "")}


def contact_create(display_name: str, email: str, given_name: str | None = None,
                   surname: str | None = None, company: str | None = None,
                   mobile_phone: str | None = None, confirm: bool = False) -> JsonDict:
    """Preview a new personal contact. Set confirm=true only after approval. Does not modify directory users. Never retry uncertain writes automatically."""
    if not display_name.strip():
        raise ValueError("display_name must not be empty")
    fields = dict(display_name=display_name, email=require_email(email), given_name=given_name,
                  surname=surname, company=company, mobile_phone=mobile_phone)
    result = {"operation": "create", "fields": fields, "created": False}
    if confirm is True:
        with outlook_client() as client:
            result.update(contact=_contact(client.create_contact(**fields)), created=True)
    return result


def contact_update(contact_id: str, display_name: str | None = None, email: str | None = None,
                   given_name: str | None = None, surname: str | None = None,
                   company: str | None = None, mobile_phone: str | None = None,
                   confirm: bool = False) -> JsonDict:
    """Preview changes to a personal contact by its contact ID, not a directory person ID. Omitted fields stay unchanged. An email replaces the whole email collection. Empty optional text fields clear. Set confirm=true only after approval. Never retry uncertain writes automatically."""
    require_id(contact_id, "contact_id")
    if email is not None:
        require_email(email)
    if display_name is not None and not display_name.strip():
        raise ValueError("display_name must not be empty")
    fields = {key: value for key, value in dict(display_name=display_name, email=email, given_name=given_name,
              surname=surname, company=company, mobile_phone=mobile_phone).items() if value is not None}
    if not fields:
        raise ValueError("Supply at least one contact field to change.")
    with outlook_client() as client:
        result = {"id": contact_id, "before": _contact(client.get_contact(contact_id)), "fields": fields, "updated": False}
        if confirm is True:
            result.update(contact=_contact(client.update_contact(contact_id, **fields)), updated=True)
    return result


def contact_org(user: str = "me", include_chain: bool = False, max_depth: int | None = None) -> JsonDict:
    """Read a directory user\'s company, department, title, manager and all direct reports. Accepts an email, UPN, directory user ID, unique name or me, not a personal contact ID. Optional include_chain adds the recorded management chain. max_depth requires include_chain and reports an incomplete chain at the limit. Missing or restricted data remains explicit. Read-only; does not edit contacts or the directory."""
    user = validate_org_options(user, include_chain, max_depth)
    with graph_client() as client:
        return organisation_view(client, user, include_chain, max_depth)


TOOLS = [contact_search, contact_create, contact_update, contact_org]
