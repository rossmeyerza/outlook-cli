"""Ask Microsoft 365 Copilot questions, with OneDrive or SharePoint files as context."""
from __future__ import annotations

import argparse
import sys

from .. import cache
from ..copilot import CopilotAnswer, CopilotClient, CopilotFile, copilot_file_from_drive_item
from ..errors import CliError, CopilotError, OutlookAPIError, TokenExpiredError, TokenNotFoundError
from ..outlook_client import OutlookClient
from ..parser import add_output_args
from ..progress import spinner
from ..runtime import console, get_graph_client, print_json
from ..token_manager import CopilotTokenManager
from .files import _find_sp_item_by_path, _looks_like_drive_item_id


def resolve_copilot_file(
    gc: OutlookClient,
    ref: str,
    *,
    site: str | None = None,
    library: str | None = None,
) -> CopilotFile:
    """Resolve a file URL, a drive item ID or a path to a CopilotFile.

    A URL can be a SharePoint or OneDrive file link, such as the webUrl from
    `files search --json`. Without a site, IDs and paths are in the user's OneDrive.
    """
    ref = ref.strip()
    if ref.startswith("https://"):
        item = gc.get_share_drive_item(ref)
    elif site:
        _, item = _find_sp_item_by_path(gc, site, ref, library)
    elif _looks_like_drive_item_id(ref):
        item = gc.drive_item_with_sharepoint_ids(ref)
    else:
        item = gc.od_item_by_path(ref)
    if "folder" in item:
        raise CopilotError(f"{item.get('name', ref)} is a folder. Give Copilot a file.")
    drive_id = (item.get("parentReference") or {}).get("driveId")
    return copilot_file_from_drive_item(gc.drive_item_with_sharepoint_ids(item["id"], drive_id))


def _last_conversation_id() -> str | None:
    items = cache.load_cache(cache.COPILOT_CACHE)
    return items[0]["Id"] if items else None


def _conversation_id(args: argparse.Namespace) -> str | None:
    if args.conversation:
        return args.conversation
    if args.continue_last:
        conversation_id = _last_conversation_id()
        if not conversation_id:
            raise CliError("[red]No earlier Copilot conversation. Run 'outlook-cli copilot ask' without --continue first.[/]")
        return conversation_id
    return None


def _resolve_files(args: argparse.Namespace) -> list[CopilotFile]:
    if not args.file:
        return []
    gc = get_graph_client()
    try:
        with spinner(args, "Finding files..."):
            return [resolve_copilot_file(gc, ref, site=args.site, library=args.library) for ref in args.file]
    except (OutlookAPIError, ValueError) as e:
        raise CliError(f"[red]Could not find the file: {e}[/]") from e
    finally:
        gc.close()


def _print_answer_footer(answer: CopilotAnswer, files: list[CopilotFile]) -> None:
    if answer.sources:
        console.print("\n[bold]Sources[/]")
        for number, source in enumerate(answer.sources, start=1):
            console.print(f"  {number}. {source['title']}\n     [dim]{source['url']}[/]", soft_wrap=True)
    elif files:
        console.print(f"\n[dim]Files: {', '.join(f.name for f in files)}[/]")
    if not answer.complete:
        console.print("[yellow]The answer may be incomplete: Copilot did not finish in time.[/]")
    console.print(f"[dim]Conversation: {answer.conversation_id} (continue with --continue)[/]")


def cmd_copilot_ask(args: argparse.Namespace) -> None:
    files = _resolve_files(args)
    conversation_id = _conversation_id(args)
    manager = CopilotTokenManager()
    try:
        manager.get_token(auto_reauth=True)
    except (TokenNotFoundError, TokenExpiredError):
        raise CliError("[red]Microsoft 365 Copilot authentication failed. Run 'outlook-cli auth' to retry.[/]") from None
    client = CopilotClient(manager)

    stream = not args.json and not args.no_stream and sys.stdout.isatty()
    streamed: list[str] = []

    def on_text(delta: str) -> None:
        streamed.append(delta)
        sys.stdout.write(delta)
        sys.stdout.flush()

    try:
        if stream:
            answer = client.ask(args.prompt, files=files, conversation_id=conversation_id,
                                timeout=args.timeout, on_text=on_text)
        else:
            with spinner(args, "Waiting for Copilot..."):
                answer = client.ask(args.prompt, files=files, conversation_id=conversation_id, timeout=args.timeout)
    except CopilotError as e:
        if streamed:
            print()
        raise CliError(f"[red]Copilot: {e}[/]") from e

    cache.save_cache(cache.COPILOT_CACHE, [{"Id": answer.conversation_id}])

    if args.json:
        result = answer.to_dict()
        result["files"] = [{"name": f.name, "id": f.item_id, "driveId": f.drive_id, "webUrl": f.web_url} for f in files]
        print_json(result)
        return
    if stream:
        # The streamed text keeps the layout of the stream; finish the line.
        print()
    else:
        print(answer.text)
    _print_answer_footer(answer, files)


def register(sub) -> None:
    p_copilot = sub.add_parser("copilot", help="Ask Microsoft 365 Copilot (internal API, can change)")
    copilot_sub = p_copilot.add_subparsers(dest="copilot_command", required=True)

    cmd_ask = copilot_sub.add_parser(
        "ask",
        help="Ask Copilot a question, optionally about OneDrive or SharePoint files",
        description=(
            "Send a prompt to Microsoft 365 Copilot. --file takes a SharePoint or OneDrive URL, "
            "a drive item ID, or a path. Paths and IDs are in OneDrive unless --site is given."
        ),
    )
    add_output_args(cmd_ask)
    cmd_ask.add_argument("prompt", help="The question or instruction for Copilot")
    cmd_ask.add_argument("--file", "-f", action="append", help="File URL, drive item ID or path (repeatable)")
    cmd_ask.add_argument("--site", help="SharePoint site name for --file paths and IDs (partial match)")
    cmd_ask.add_argument("--library", help="SharePoint document library for --file paths and IDs")
    conversation = cmd_ask.add_mutually_exclusive_group()
    conversation.add_argument("--conversation", help="Continue the conversation with this ID")
    conversation.add_argument("--continue", dest="continue_last", action="store_true",
                              help="Continue the last conversation from this account")
    cmd_ask.add_argument("--no-stream", action="store_true", help="Wait for the full answer, then print it")
    cmd_ask.add_argument("--timeout", type=float, default=240, help="Seconds to wait for the answer (default 240)")
    cmd_ask.set_defaults(func=cmd_copilot_ask)

