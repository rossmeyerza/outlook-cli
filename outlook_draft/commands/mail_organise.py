"""CLI commands for Graph categories and inbox rules."""

from __future__ import annotations

import json

from .. import runtime
from ..errors import CliError, OutlookAPIError
from ..parser import add_output_args


def _json_object(value: str) -> dict:
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError as exc:
        raise ValueError(f"Invalid JSON: {exc.msg}") from exc
    if not isinstance(parsed, dict):
        raise ValueError("The JSON value must be an object.")
    return parsed


def _rule_payload(args) -> dict:
    payload = {}
    for field, attribute in (("displayName", "name"), ("sequence", "sequence"), ("isEnabled", "enabled")):
        value = getattr(args, attribute, None)
        if value is not None:
            payload[field] = value
    for field in ("conditions", "actions"):
        value = getattr(args, field, None)
        if value is not None:
            payload[field] = _json_object(value)
    return payload


def _invoke(client, args):
    operation = args.organise_operation
    confirm = getattr(args, "yes", False)
    if operation == "categorise":
        return client.categorise_messages(args.message_ids, args.add, args.remove, confirm=confirm)
    if operation.endswith("_list"):
        method = client.list_mail_categories if operation == "categories_list" else client.list_mail_rules
        return method()
    if operation == "categories_create":
        return client.create_mail_category(args.name, args.colour, confirm=confirm)
    if operation == "categories_update":
        return client.update_mail_category(args.id, args.colour, confirm=confirm)
    if operation == "categories_delete":
        return client.delete_mail_category(args.id, confirm=confirm)
    if operation == "rules_create":
        return client.create_mail_rule(_rule_payload(args), confirm=confirm)
    if operation == "rules_update":
        return client.update_mail_rule(args.id, _rule_payload(args), confirm=confirm)
    return client.delete_mail_rule(args.id, confirm=confirm)


def cmd_mail_organise(args) -> None:
    client = runtime.get_graph_client()
    try:
        result = _invoke(client, args)
    except (ValueError, OutlookAPIError) as exc:
        raise CliError(str(exc)) from exc
    finally:
        client.close()
    runtime.print_json(result)
    if isinstance(result, dict) and result.get("failed"):
        raise CliError("Some messages failed. Check the per-message results before any retry.")


def _command(subparsers, name: str, operation: str):
    parser = subparsers.add_parser(name)
    add_output_args(parser)
    parser.set_defaults(func=cmd_mail_organise, organise_operation=operation)
    if not operation.endswith("_list"):
        parser.add_argument("--yes", action="store_true", help="Apply the previewed mailbox change")
    return parser


def _register_categories(mail_sub) -> None:
    categories = mail_sub.add_parser("categories", help="Manage Graph master categories")
    sub = categories.add_subparsers(dest="category_command", required=True)
    _command(sub, "list", "categories_list")
    create = _command(sub, "create", "categories_create")
    create.add_argument("name")
    create.add_argument("--colour", "--color", default="none", help="none or preset0 through preset24")
    update = _command(sub, "update", "categories_update")
    update.add_argument("id", help="Full Graph category ID; names cannot be changed")
    update.add_argument("--colour", "--color", required=True)
    delete = _command(sub, "delete", "categories_delete")
    delete.add_argument("id", help="Full Graph category ID")


def _rule_arguments(parser, *, create: bool) -> None:
    if create:
        parser.add_argument("name")
    else:
        parser.add_argument("id", help="Full Graph rule ID")
        parser.add_argument("--name")
    parser.add_argument("--sequence", type=int, default=0 if create else None)
    parser.add_argument("--conditions", required=create, help="Graph lowerCamelCase conditions JSON")
    parser.add_argument("--actions", required=create, help="Graph lowerCamelCase actions JSON")
    enabled = parser.add_mutually_exclusive_group()
    enabled.add_argument("--enabled", action="store_true", dest="enabled")
    enabled.add_argument("--disabled", action="store_false", dest="enabled")
    parser.set_defaults(enabled=False if create else None)


def register(mail_sub) -> None:
    _register_categories(mail_sub)
    categorise = _command(mail_sub, "categorise", "categorise")
    categorise.add_argument("message_ids", nargs="+", help="Full Graph message IDs only; no cached indices. REST IDs can differ.")
    categorise.add_argument("--add", action="append", help="Category name to add; repeatable")
    categorise.add_argument("--remove", action="append", help="Category name to remove; repeatable")
    rules = mail_sub.add_parser("rules", help="Manage Graph inbox rules; new rules start disabled")
    sub = rules.add_subparsers(dest="rule_command", required=True)
    _command(sub, "list", "rules_list")
    _rule_arguments(_command(sub, "create", "rules_create"), create=True)
    _rule_arguments(_command(sub, "update", "rules_update"), create=False)
    delete = _command(sub, "delete", "rules_delete")
    delete.add_argument("id", help="Full Graph rule ID")
