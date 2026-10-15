"""CLI registration for gateway_parser."""
from __future__ import annotations

from .gateway import cmd_gateway_start, cmd_gateway_status, cmd_gateway_stop


def register(sub) -> None:
    p_gateway = sub.add_parser("gateway", help="Teams-to-pi gateway daemon")
    gateway_sub = p_gateway.add_subparsers(dest="command", required=True)

    gw_start = gateway_sub.add_parser("start", help="Start the gateway daemon")
    gw_start.add_argument("--chat-id", metavar="ID", help="Teams chat ID to monitor")
    gw_start.add_argument("--self-chat", action="store_true", help="Monitor your Teams self-chat")
    gw_start.add_argument("--trigger", metavar="TEXT", help="Trigger string (default: @Marlow)")
    gw_start.add_argument("--poll", type=int, metavar="SECONDS", help="Poll interval seconds (default: 30)")
    gw_start.add_argument("--provider", metavar="NAME", help="Pi provider name")
    gw_start.add_argument("--model", metavar="MODEL", help="Pi model pattern or ID")
    gw_start.add_argument("--thinking", metavar="LEVEL", help="Pi thinking level")
    gw_start.add_argument("--models", metavar="PATTERNS", help="Comma-separated Pi model cycle patterns")
    gw_start.set_defaults(func=cmd_gateway_start)

    gw_stop = gateway_sub.add_parser("stop", help="Stop the gateway daemon")
    gw_stop.set_defaults(func=cmd_gateway_stop)

    gw_status = gateway_sub.add_parser("status", help="Show gateway status and recent log")
    gw_status.set_defaults(func=cmd_gateway_status)
