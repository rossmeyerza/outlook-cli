"""Microsoft Graph category and inbox rule tools with explicit confirmation."""

from __future__ import annotations

from .clients import graph_client


def mail_categories() -> list[dict]:
    """List master categories through Graph. Requires MailboxSettings.Read."""
    with graph_client() as client:
        return client.list_mail_categories()


def mail_category_create(name: str, colour: str = "none", confirm: bool = False) -> dict:
    """Preview a category creation. Set confirm=true to create it. Requires MailboxSettings.ReadWrite. Colour: none or preset0..preset24."""
    with graph_client() as client:
        return client.create_mail_category(name, colour, confirm=confirm)


def mail_category_update(category_id: str, colour: str, confirm: bool = False) -> dict:
    """Preview a category colour update. Set confirm=true to apply it. Requires MailboxSettings.ReadWrite. A category name cannot change."""
    with graph_client() as client:
        return client.update_mail_category(category_id, colour, confirm=confirm)


def mail_category_delete(category_id: str, confirm: bool = False) -> dict:
    """Preview a master category deletion. Set confirm=true to delete it. Requires MailboxSettings.ReadWrite."""
    with graph_client() as client:
        return client.delete_mail_category(category_id, confirm=confirm)


def mail_categorise(message_ids: list[str], add: list[str] | None = None,
                    remove: list[str] | None = None, confirm: bool = False) -> dict:
    """Preview category changes on full Graph message IDs, not cached indices. Set confirm=true to apply. Requires Mail.ReadWrite. REST v2 IDs may differ from Graph IDs. Returns per-message results; a failed write stops the batch. Never retry an uncertain result automatically."""
    with graph_client() as client:
        return client.categorise_messages(message_ids, add, remove, confirm=confirm)


def mail_rules() -> list[dict]:
    """List inbox rules through Graph. Requires MailboxSettings.Read."""
    with graph_client() as client:
        return client.list_mail_rules()


def mail_rule_create(name: str, conditions: dict, actions: dict, sequence: int = 0,
                     enabled: bool = False, confirm: bool = False) -> dict:
    """Preview an inbox rule. Set confirm=true to create it. Rules start disabled unless enabled=true. Requires MailboxSettings.ReadWrite. Use Graph lowerCamelCase JSON: conditions fromAddresses (recipient objects), subjectContains; actions moveToFolder (Graph folder ID), assignCategories, markAsRead, stopProcessingRules. Other fields and dangerous actions are rejected."""
    payload = {"displayName": name, "conditions": conditions, "actions": actions,
               "sequence": sequence, "isEnabled": enabled}
    with graph_client() as client:
        return client.create_mail_rule(payload, confirm=confirm)


def mail_rule_update(rule_id: str, changes: dict, confirm: bool = False) -> dict:
    """Preview a rule update. Set confirm=true to apply it. Requires MailboxSettings.ReadWrite. Changes accept displayName, sequence, isEnabled, conditions and actions in Graph lowerCamelCase. Unknown keys and dangerous actions fail validation. An enable request also checks the current rule."""
    with graph_client() as client:
        return client.update_mail_rule(rule_id, changes, confirm=confirm)


def mail_rule_delete(rule_id: str, confirm: bool = False) -> dict:
    """Preview an inbox rule deletion. Set confirm=true to delete it. Requires MailboxSettings.ReadWrite."""
    with graph_client() as client:
        return client.delete_mail_rule(rule_id, confirm=confirm)


TOOLS = [mail_categories, mail_category_create, mail_category_update, mail_category_delete,
         mail_categorise, mail_rules, mail_rule_create, mail_rule_update, mail_rule_delete]
