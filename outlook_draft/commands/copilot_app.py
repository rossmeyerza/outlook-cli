"""CLI commands for Copilot in Office web apps and the Chrome native host."""
from __future__ import annotations

import argparse
import json
import sys

from ..copilot_bridge.app import ask_app, resolve_app_file
from ..copilot_bridge.client import BridgeClient
from ..copilot_bridge.install import install
from ..errors import CliError, CopilotError, OutlookAPIError
from ..parser import add_output_args
from ..runtime import get_graph_client, print_json


def cmd_status(args: argparse.Namespace) -> None:
    try:
        result = BridgeClient().status()
    except CopilotError as exc:
        raise CliError(str(exc)) from exc
    if args.json:
        print_json(result)
    else:
        print(f"Host: protocol {result['host']}. Extension: {result.get('extension', 'unknown')}.")
        for tab in result.get("tabs", []):
            print(f"  {tab['tabId']}: {tab['app']} | {tab['title']} | pane open: {tab['paneOpen']} | {_busy_text(tab)}")
        if not result.get("tabs"):
            print("No Office editor tabs found.")


def _busy_text(tab: dict) -> str:
    busy = "busy: True"
    if tab.get("busy") and tab.get("busyAgeSeconds") is not None:
        busy += f" ({tab.get('phase')}, {tab['busyAgeSeconds']} s)"
    elif not tab.get("busy"):
        busy = "busy: False"
    if tab.get("uncertain"):
        age = tab.get("uncertainAgeSeconds")
        detail = ", ".join(str(part) for part in (tab.get("uncertainReason"), f"{age} s ago" if age is not None else None) if part)
        busy += f" | uncertain: True ({detail})" if detail else " | uncertain: True"
    return busy


def cmd_reset(args: argparse.Namespace) -> None:
    try:
        result = BridgeClient().reset(args.tab_id)
    except CopilotError as exc:
        raise CliError(str(exc)) from exc
    if args.json:
        print_json(result)
        return
    cleared = result.get("cleared") or {}
    if not any(cleared.values()):
        print(f"Tab {args.tab_id} had no busy or uncertain state.")
        return
    print(f"Tab {args.tab_id}: cleared {', '.join(k for k, v in cleared.items() if v)} state.")
    print("Copilot can still have changed the document. Check it before you send another prompt.", file=sys.stderr)


def cmd_inspect(args: argparse.Namespace) -> None:
    try:
        result = BridgeClient().inspect(args.tab_id)
    except CopilotError as exc:
        raise CliError(str(exc)) from exc
    print_json(result)


def cmd_install(args: argparse.Namespace) -> None:
    try:
        result = install()
    except (CopilotError, OSError) as exc:
        raise CliError(str(exc)) from exc
    print(f"Native host manifest: {result['manifest']}")
    print("Open chrome://extensions. Turn on Developer mode. Select Load unpacked.")
    print(f"Select this folder: {result['extension']}")
    print(f"Extension ID: {result['extensionId']}")
    print("Then run: outlook-cli copilot app status")


def _event_printer(event: dict) -> None:
    kind = event.get("type")
    if kind == "text":
        sys.stdout.write(event["text"])
        sys.stdout.flush()
    elif kind in {"action", "action_result", "approval_required"}:
        print(json.dumps(event, ensure_ascii=False), file=sys.stderr)
    elif kind == "progress":
        print(event["text"], file=sys.stderr)
    elif kind == "phase":
        print(f"[{event.get('phase')}]", file=sys.stderr)


def _ask(args: argparse.Namespace) -> dict:
    gc = get_graph_client()
    try:
        file = resolve_app_file(gc, args.file, site=args.site, library=args.library, app=args.app)
        if file["shared"]:
            print("Caution: this file is outside your own OneDrive. Other people can see edits at once.", file=sys.stderr)
        print("Copilot can edit the live document. The bridge never approves a request.", file=sys.stderr)
        return ask_app(gc, args.prompt, file, new_chat=args.new_chat, open_tab=args.open_tab,
                       wait_saved=args.wait_saved, timeout=args.timeout,
                       on_event=None if args.json else _event_printer)
    finally:
        gc.close()


def cmd_ask(args: argparse.Namespace) -> None:
    try:
        result = _ask(args)
    except (CopilotError, OutlookAPIError, ValueError) as exc:
        if args.json:
            print_json({"status": "error", "message": str(exc), "applied": None, "saved": None})
            raise CliError("") from exc
        raise CliError(str(exc)) from exc
    if args.json:
        print_json(result)
    else:
        print()
        if result.get("message"):
            print(result["message"], file=sys.stderr)
        print(f"Status: {result.get('status')}. Applied: {result.get('applied')}. Saved: {result.get('saved')}.", file=sys.stderr)
        if result.get("completion") == "settled":
            print("Completion: the pane went idle without a final answer event.", file=sys.stderr)
        if result.get("status") != "completed" and result.get("phase"):
            print(f"Last phase: {result['phase']}.", file=sys.stderr)
    if result.get("status") == "approval_required":
        print("Copilot asked for approval. Nothing was approved by the bridge.", file=sys.stderr)
        raise CliError("", exit_code=3)
    if result.get("status") != "completed":
        if result.get("sent"):
            print("The prompt was sent. Copilot can already have changed the document. Check it first.", file=sys.stderr)
        print("The request did not complete. Do not retry automatically.", file=sys.stderr)
        raise CliError("")


def register(sub) -> None:
    app = sub.add_parser("app", help="Copilot in a live Office web editor (edits the document)")
    commands = app.add_subparsers(dest="copilot_app_command", required=True)
    status = commands.add_parser("status", help="Show the native host, extension and Office tabs")
    add_output_args(status)
    status.set_defaults(func=cmd_status)
    inspect = commands.add_parser("inspect", help="Read Office control states and protocol schema names, not document content")
    add_output_args(inspect)
    inspect.add_argument("--tab-id", type=int, required=True)
    inspect.set_defaults(func=cmd_inspect)
    reset = commands.add_parser("reset", help="Clear a stale busy or uncertain tab state after you check the document")
    add_output_args(reset)
    reset.add_argument("--tab-id", type=int, required=True)
    reset.set_defaults(func=cmd_reset)
    ask = commands.add_parser("ask", help="Send a prompt to Excel or PowerPoint Copilot in Chrome")
    add_output_args(ask)
    ask.add_argument("prompt")
    ask.add_argument("--file", "-f", required=True, help="OneDrive or SharePoint URL, ID or path")
    ask.add_argument("--site")
    ask.add_argument("--library")
    ask.add_argument("--app", choices=("excel", "powerpoint", "word"))
    ask.add_argument("--new-chat", action="store_true")
    ask.add_argument("--open", dest="open_tab", action="store_true")
    ask.add_argument("--wait-saved", action="store_true")
    ask.add_argument("--timeout", type=float, default=300)
    ask.set_defaults(func=cmd_ask)
    extension = sub.add_parser("extension", help="Set up the Copilot Chrome bridge")
    ext_commands = extension.add_subparsers(dest="copilot_extension_command", required=True)
    ext_install = ext_commands.add_parser("install", help="Install the native host and copy the unpacked extension")
    ext_install.set_defaults(func=cmd_install)
