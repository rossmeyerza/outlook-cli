"""Mail queries, conversations and bulk command handlers."""
from __future__ import annotations

from .. import runtime
from ..cache import MAIL_CACHE, save_cache
from ..errors import CliError, OutlookAPIError
from ..mail_query import BULK_ACTIONS
from ..parser import add_output_args
from ..references import _resolve_message_id


def query_options(parser):
    parser.add_argument("--from", dest="sender", help="Exact sender email address")
    parser.add_argument("--to", dest="recipient", help="Exact To or CC email address")
    parser.add_argument("--after", help="Inclusive ISO date or timestamp; UTC if no zone")
    parser.add_argument("--before", help="Exclusive ISO date or timestamp; UTC if no zone")
    parser.add_argument("--has-attachments", action="store_true", default=None)
    parser.add_argument("--unread", action="store_true", default=None)
    parser.add_argument("--folder", help="Folder name, cached index or full ID")
    parser.add_argument("--all", dest="all_pages", action="store_true", help="Fetch all available pages")
    parser.add_argument("--page", help="Continuation URL from a previous result; keep query and filters unchanged")


def enhanced_search(args):
    return any(getattr(args, field, None) is not None for field in
               ("sender", "recipient", "after", "before", "has_attachments", "unread", "folder", "page")) or args.all_pages or not args.query


def cmd_query(args):
    from .mail import _message_summary, _resolve_folder_id

    filters = {key: getattr(args, key) for key in
               ("sender", "recipient", "after", "before", "has_attachments", "unread") if getattr(args, key) is not None}
    client = runtime.get_client()
    try:
        folder = _resolve_folder_id(args, args.folder) if args.folder else None
        result = client.query_mail(args.query or "", filters=filters, folder=folder, page=args.page,
                                   all_pages=args.all_pages, page_size=args.count)
    except (ValueError, OutlookAPIError) as exc:
        raise CliError(str(exc)) from exc
    finally:
        client.close()
    save_cache(MAIL_CACHE, result["messages"])
    result["messages"] = [_message_summary(args, message, i) for i, message in enumerate(result["messages"], 1)]
    runtime.print_json(result)


def cmd_conversation(args):
    client = runtime.get_client()
    try:
        mid = _resolve_message_id(client, args.message_id)
        result = client.mail_conversation(mid)
    except (ValueError, OutlookAPIError) as exc:
        raise CliError(str(exc)) from exc
    finally:
        client.close()
    runtime.print_json(result)


def cmd_bulk(args):
    from ..mail_query import MailQueryMixin
    from .mail import _resolve_folder_id

    ids = [_resolve_message_id(None, ref) for ref in args.message_ids]
    folder = _resolve_folder_id(args, args.folder) if args.folder else None
    action = "trash" if args.command == "trash" else args.action
    if not args.yes:
        try:
            runtime.print_json(MailQueryMixin.bulk_mail(None, ids, action, folder=folder))
        except ValueError as exc:
            raise CliError(str(exc)) from exc
        return
    client = runtime.get_client()
    try:
        result = client.bulk_mail(ids, action, folder=folder, confirm=True)
    except (ValueError, OutlookAPIError) as exc:
        raise CliError(str(exc)) from exc
    finally:
        client.close()
    runtime.print_json(result)
    if any(item["status"] != "applied" for item in result["results"]):
        raise CliError("Bulk operation stopped. Check the reported results before any retry.")


def register(mail_sub):
    conversation = mail_sub.add_parser("conversation", help="Read a complete conversation and attachment metadata")
    conversation.add_argument("message_id", help="Any message ID or cached reference in the conversation")
    add_output_args(conversation)
    conversation.set_defaults(func=cmd_conversation)
    for name in ("bulk", "trash"):
        parser = mail_sub.add_parser(name, help="Preview changes to multiple emails; use --yes to apply")
        if name == "bulk":
            parser.add_argument("action", choices=sorted(BULK_ACTIONS))
        parser.add_argument("message_ids", nargs="+", help="Explicit message IDs or cached references")
        parser.add_argument("--folder", help="Destination for the move action")
        parser.add_argument("--yes", action="store_true", help="Apply the previewed operation")
        add_output_args(parser)
        parser.set_defaults(func=cmd_bulk)
