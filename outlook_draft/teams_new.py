"""Resolve recipients and start a new Teams chat with a first message."""

from __future__ import annotations

import uuid
from typing import Any

from .errors import OutlookAPIError

JsonDict = dict[str, Any]
MAX_CANDIDATES = 5


class ChatCreatedSendFailed(Exception):
    """Graph created the chat, but the first message failed."""

    def __init__(self, chat: JsonDict, error: OutlookAPIError):
        self.chat = chat
        self.error = error
        super().__init__(
            f"Chat {chat.get('id', '')} exists, but the message failed: {error}. "
            "Send it with 'teams send <chat-id>'. Do not start the chat again."
        )


def split_refs(values: list[str] | None) -> list[str]:
    """Split repeated and comma-separated recipient values."""
    refs: list[str] = []
    for value in values or []:
        refs.extend(part.strip() for part in value.split(",") if part.strip())
    return refs


def _terms_match(name: str, query: str) -> bool:
    haystack = name.casefold()
    return all(term in haystack for term in query.casefold().split())


def _person_address(person: JsonDict) -> str:
    emails = person.get("scoredEmailAddresses") or []
    return person.get("userPrincipalName") or (emails[0].get("address") if emails else "") or ""


def _is_org_user(person: JsonDict) -> bool:
    # Personal contacts cannot join a Teams chat, so only directory users count.
    subclass = (person.get("personType") or {}).get("subclass")
    return subclass in (None, "OrganizationUser")


def _resolve_name(client: Any, name: str) -> JsonDict:
    people = [
        person for person in client.search_people(name, top=10)
        if _terms_match(person.get("displayName") or "", name)
        and _person_address(person) and _is_org_user(person)
    ]
    if not people:
        raise ValueError(f"No person in your organisation matches '{name}'. Use an email address.")
    if len(people) > 1:
        listed = "; ".join(
            f"{person.get('displayName')} <{_person_address(person)}>"
            for person in people[:MAX_CANDIDATES]
        )
        raise ValueError(f"'{name}' matches more than one person: {listed}. Use an email address.")
    user = client.get_directory_user(_person_address(people[0]))
    if not user:
        raise ValueError(f"'{name}' is not a user in your organisation directory.")
    return user


def _is_object_id(ref: str) -> bool:
    try:
        uuid.UUID(ref)
    except ValueError:
        return False
    return True


def resolve_recipient(client: Any, ref: str) -> JsonDict:
    """Resolve an email address, UPN, user ID or a unique name to a directory user."""
    if "@" not in ref and not _is_object_id(ref):
        return _resolve_name(client, ref)
    user = client.get_directory_user(ref)
    if not user:
        raise ValueError(f"'{ref}' is not a user in your organisation directory. Guests are not supported.")
    return user


def resolve_recipients(client: Any, refs: list[str]) -> tuple[JsonDict, list[JsonDict]]:
    """Return the current user and the other members, without duplicates."""
    if not refs:
        raise ValueError("At least one recipient is required.")
    me = client.get_current_user()
    users: list[JsonDict] = []
    seen = {me["id"]}
    for ref in refs:
        user = resolve_recipient(client, ref)
        if user["id"] not in seen:
            seen.add(user["id"])
            users.append(user)
    if not users:
        raise ValueError("The recipients are only you. Use 'teams self' for your self-chat.")
    return me, users


def chat_type_for(users: list[JsonDict]) -> str:
    return "oneOnOne" if len(users) == 1 else "group"


def validate_new_chat(users: list[JsonDict], message: str, topic: str | None) -> None:
    if not message.strip():
        raise ValueError("The message must not be empty.")
    if topic and chat_type_for(users) == "oneOnOne":
        raise ValueError("A 1:1 chat cannot have a topic. Remove --topic or add more recipients.")


def start_chat(
    client: Any,
    me: JsonDict,
    users: list[JsonDict],
    message: str,
    *,
    content_type: str = "text",
    topic: str | None = None,
) -> tuple[JsonDict, JsonDict]:
    """Create the chat and send the first message. Neither call retries."""
    validate_new_chat(users, message, topic)
    chat = client.create_teams_chat(me["id"], [user["id"] for user in users], topic=topic or None)
    try:
        sent = client.send_teams_message(chat["id"], message, content_type=content_type)
    except OutlookAPIError as exc:
        raise ChatCreatedSendFailed(chat, exc) from exc
    return chat, sent


def user_label(user: JsonDict) -> str:
    address = user.get("mail") or user.get("userPrincipalName") or user.get("id", "")
    return f"{user.get('displayName') or address} <{address}>"
