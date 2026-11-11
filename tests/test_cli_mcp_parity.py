"""Keep the complete comparison table consistent with the parser and MCP registry."""
from __future__ import annotations

import argparse
import asyncio
import re
from pathlib import Path

from outlook_draft.cli import build_parser
from outlook_draft.mcp.server import build_server

DOCUMENT = Path(__file__).parents[1] / "docs" / "cli-mcp-parity.md"
STATUSES = {"Full", "Partial", "Missing", "Local only"}


def parser_leaves(parser, prefix=()):
    subparsers = [action for action in parser._actions if isinstance(action, argparse._SubParsersAction)]
    if not subparsers:
        yield prefix, parser
    for action in subparsers:
        for name, child in action.choices.items():
            yield from parser_leaves(child, (*prefix, name))


def comparison_rows():
    for line in DOCUMENT.read_text().splitlines():
        cells = [cell.strip() for cell in line.split("|")]
        if len(cells) == 6 and cells[2] in STATUSES:
            yield cells[1].strip("`"), cells[2], set(re.findall(r"`([^`]+)`", cells[3]))


def test_every_executable_cli_operation_has_exactly_one_comparison_row():
    leaves = list(parser_leaves(build_parser()))
    expected = set()
    for path, parser in leaves:
        if path == ("auth",):
            action = next(action for action in parser._actions if not action.option_strings and action.choices)
            expected.update(f"auth {choice}" for choice in action.choices)
        else:
            expected.add(" ".join(path))
    rows = list(comparison_rows())
    actual = [command for command, _, _ in rows]
    assert len(actual) == len(set(actual))
    assert set(actual) == expected
    assert f"**{len(leaves)} executable parser paths**" in DOCUMENT.read_text()
    assert f"**{len(expected)} distinct operations**" in DOCUMENT.read_text()


def test_comparison_names_every_registered_mcp_tool_and_no_unknown_tools():
    names = {tool.name for tool in asyncio.run(build_server().list_tools())}
    documented = set().union(*(tools for _, _, tools in comparison_rows()))
    assert documented == names
    assert f"**{len(names)} tools**" in DOCUMENT.read_text()


def test_all_requested_domains_have_full_mcp_coverage():
    domains = {"mail", "draft", "cal", "contact", "task", "teams", "files"}
    scoped = [(command, status) for command, status, _ in comparison_rows() if command.split()[0] in domains]
    assert all(status == "Full" for _, status in scoped)
    assert f"All **{len(scoped)} CLI operations**" in DOCUMENT.read_text()
