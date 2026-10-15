"""Contact tools."""

from __future__ import annotations

from ..commands.contacts import _contact_id, _type_label
from .clients import clamp, outlook_client
from .shaping import JsonDict


def contact_search(query: str, count: int = 10) -> list[JsonDict]:
    """Search the org directory and recent contacts by name or email. Read-only."""
    with outlook_client() as client:
        people = client.search_people(query, top=clamp(count, high=50))
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


TOOLS = [contact_search]
