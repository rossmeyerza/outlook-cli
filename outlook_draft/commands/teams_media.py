"""CLI handlers for Teams media sends."""
from __future__ import annotations

from .. import runtime, teams_new
from ..errors import CliError, OutlookAPIError, OutlookDraftError
from ..teams_media import (
    DEFAULT_UPLOAD_FOLDER,
    media_plan,
    prepare_media,
    resolve_upload_folder,
    send_media,
    validate_content,
)


def _prepared(args):
    try:
        media = prepare_media(getattr(args, "attach", None), getattr(args, "image", None))
        validate_content(args.message, media, is_html=args.html)
        return media
    except ValueError as exc:
        raise CliError(str(exc)) from exc


def cmd_send(args):
    from .teams import _resolve_teams_chat_id

    media = _prepared(args)
    client = runtime.get_graph_client() if media.files or args.yes else None
    try:
        chat_id = _resolve_teams_chat_id(args, args.chat_id)
        if args.yes:
            result = send_media(client, chat_id, args.message, media, is_html=args.html, upload_folder=args.upload_folder)
        else:
            result = media_plan(client, chat_id, args.message, media, args.upload_folder, is_html=args.html)
        runtime.print_json(result)
        if not args.yes:
            raise CliError("", 2)
    except (ValueError, OutlookAPIError, OutlookDraftError) as exc:
        if isinstance(exc, CliError):
            raise
        raise CliError(str(exc)) from exc
    finally:
        if client is not None:
            client.close()


def cmd_new(args):
    media = _prepared(args)
    client = runtime.get_graph_client()
    try:
        me, users = teams_new.resolve_recipients(client, teams_new.split_refs(args.to))
        teams_new.validate_new_chat(users, args.message, args.topic, has_media=bool(media))
        result = {"sent": False, "chatType": teams_new.chat_type_for(users), "members": users,
                  "content": args.message, "topic": args.topic, **media.summary()}
        if media.files:
            folder = resolve_upload_folder(client, args.upload_folder)
            result.update(uploadFolder={"id": folder["id"], "path": args.upload_folder},
                          accessUserIds=[user["id"] for user in users])
        if args.yes:
            chat, sent = teams_new.start_chat(client, me, users, args.message,
                                              content_type="html" if args.html else "text", topic=args.topic,
                                              media=media, upload_folder=args.upload_folder)
            result.update(sent=True, chat_id=chat["id"], message=sent)
        runtime.print_json(result)
        if not args.yes:
            raise CliError("", 2)
    except (ValueError, OutlookDraftError, teams_new.ChatCreatedSendFailed) as exc:
        if isinstance(exc, CliError):
            raise
        raise CliError(str(exc)) from exc
    finally:
        client.close()


def add_media_args(parser):
    parser.add_argument("--attach", action="append", help="Local file attachment; repeatable, at most 25 MiB per file")
    parser.add_argument("--image", action="append", help="Inline PNG, JPEG or GIF; repeatable, at most 1 MiB per image")
    parser.add_argument("--upload-folder", default=DEFAULT_UPLOAD_FOLDER,
                        help="Existing private OneDrive folder for files; no folders are created")
