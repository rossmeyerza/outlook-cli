# Refactor report

Baseline: `3f9e352`. The tree was clean before work. Phases ran in order. No commits were pushed.

## Commits and tests

Each phase ended with the full test suite passing before its commit.

| Phase | Commit | Message | Tests |
|---|---|---|---|
| Baseline | `3f9e352` | docs: add refactor plan | 153 passed, 1 skipped |
| 0 | `31c28ba` | test: record CLI help output and safety contracts | 164 passed, 1 skipped |
| 1 | `7249785` | refactor: share CLI runtime and reference resolution | 164 passed, 1 skipped |
| 2 | `ff55037` | refactor: register CLI parsers in command modules | 164 passed, 1 skipped |
| 3 | `51e6121` | refactor: consolidate Graph requests on the httpx client | 174 passed, 1 skipped |
| 4 | `9a910d0` | refactor: centralize CLI failures and log best-effort errors | 178 passed, 1 skipped |
| 5 | `5e026e0` | perf: fetch Teams chat details with bounded ordered workers | 185 passed, 1 skipped |
| 6 | Commit containing this report | chore: lint and verify the refactored CLI | 196 passed, 1 skipped |

The final response supplies the phase 6 hash. The skipped test requires explicit live calendar mutation permission. It was not enabled.

## Line counts

Counts use `wc -l`. New command modules include their parser registration. A dash means the file did not exist at baseline.

| File | Before | After |
|---|---:|---:|
| `outlook_draft/cli.py` | 1626 | 127 |
| `commands/__init__.py` | 0 | 0 |
| `commands/account.py` | - | 489 |
| `commands/calendar.py` | 504 | 620 |
| `commands/completion.py` | - | 19 |
| `commands/contacts.py` | 156 | 159 |
| `commands/drafts.py` | - | 366 |
| `commands/files.py` | 903 | 735 |
| `commands/gateway.py` | 1426 | 1426 |
| `commands/gateway_parser.py` | - | 26 |
| `commands/mail.py` | 536 | 612 |
| `commands/mcp.py` | - | 16 |
| `commands/pi_session.py` | 246 | 251 |
| `commands/signature.py` | 189 | 213 |
| `commands/tasks.py` | 180 | 179 |
| `commands/teams.py` | 713 | 789 |

`commands/gateway.py` is byte-for-byte unchanged from baseline.

## Dependencies and environment

- Removed `requests>=2.31` from runtime dependencies. Auth probes and signature metadata reads now use the existing `httpx` dependency.
- No runtime dependencies were added.
- Updated editable installation metadata with `uv pip install --python .venv/bin/python --no-deps -e .`.
- Confirmed that installed `outlook-draft-cli` metadata no longer declares `requests`. Existing installed packages were not separately removed.
- Never ran `uv run`, `uv sync` or `uv lock`. There is no `uv.lock`.
- Tests ran on Python 3.13.12. Ruff targets Python 3.12. A separate Python 3.12 test run was not performed.

## Teams timings

Single wall-time measurements against the live account. Both commands used `--no-spinner`. Before measurements ran after phase 4, before the concurrency changes. After measurements ran in phase 5. All four runs exited 0.

| Command | Before | After |
|---|---:|---:|
| `outlook-cli teams list -n 20` | 2.370 s | 1.709 s |
| `outlook-cli teams search refactor-no-match-20261015 --scan 100` | 23.180 s | 14.450 s |

The search scanned all 100 chats. Its output matched exactly. Live list output did not match exactly: a chat timestamp changed, and Graph returned some participant lists in a different order. The code does not sort participant lists. Fixed-data tests compare sequential and concurrent CLI output for list, received-message sorting and search. Separate tests cover reverse worker completion, stable MCP order, member failures, bounded early exit and shared HTTP client initialization.

Search work is submitted in batches of eight. An early match can leave at most seven extra reads in the current batch. Those extra reads do not increase the displayed scan or error counts. Request retry and `Retry-After` behaviour remain covered by offline tests. The HTTP client is shared across workers. Locks protect lazy client creation and token operations.

These are network measurements, not a repeated benchmark. They do not establish a fixed speedup for other accounts or queries.

## Verification

- `.venv/bin/python -m pytest -q`: **196 passed, 1 skipped**.
- `.venv/bin/python -m ruff check outlook_draft tests`: **All checks passed**.
- `git diff --check`: clean.
- All recorded phase 0 help and command outputs still match.
- Tests verify that all registered handlers are callable.
- Tests verify all three send previews exit 2 before creating a client. Confirmed sends are tested with fake clients only.
- All 36 MCP tool definitions, including input/output schemas and descriptions, match definitions extracted from `3f9e352`. `tests/mcp_tool_contracts.json` stores that baseline.
- MCP tests retain stdout, error handling and no-auto-reauth checks.
- `pi mcp list` reported **outlook: connected, 36 tools** at the configured executable path. The overall command returned 1. An unrelated `meta-devtools` server reported that it needs sign-in.

### Read-only live checks

All checks below exited 0. The global executable was used after its editable metadata refresh.

| Command | Result |
|---|---|
| `outlook-cli --help` | Help displayed |
| `outlook-cli auth status --json` | Both tokens present and valid |
| `outlook-cli mail unread -n 3 --json` | 3 messages |
| `outlook-cli cal agenda -d 1 --json` | 10 events |
| `outlook-cli teams list -n 5` | Chat table displayed |
| `outlook-cli files sites --json` | 28 sites |

The Teams timing commands were also read-only. No mail or Teams messages were sent. No drafts, events, tasks, contacts or other live objects were created. Live checks can update local CLI caches. Browser login and signature capture were not run live.

## Deviations and remaining risks

1. **Teams JSON does not exist at baseline.** `teams list --json` is rejected. Global `--json teams list` is accepted but prints a table. Phase 0 records that existing table output rather than adding JSON behaviour.
2. **Gateway registration has its own module.** `commands/gateway_parser.py` registers the gateway parser. This avoids changing the protected daemon module.
3. **Legacy error behaviour is retained.** File handlers can print an error and return exit 0. Signature lookup can also report no result without failing. These were not changed to exit 1. Some multi-line errors still print context before raising an empty or final-message `CliError`, so original output and colours stay intact. Send previews still use exit 2 directly.
4. **Exception review excludes gateway code.** Outside that module, broad best-effort catches now record debug tracebacks instead of silently swallowing failures. Broad catches were retained where narrowing could change fallback behaviour.
5. **Lint has two exclusions.** Global `E501` is ignored for existing help, API and layout literals. All rules are ignored for `commands/gateway.py` because edits there were prohibited. These exclusions are documented in `pyproject.toml`.
6. **Complexity work is limited to low-risk changes.** Auth action dispatch was split, long parser registrations were split by command, and Teams rendering was extracted. Browser signature capture and Pi process-control code were not decomposed further. Existing long functions remain, notably signature capture (137 lines), Pi prompt processing (69), calendar detail display (62), contact display (57), and task display (57). Thus a strict 50-line sweep was not completed.
7. **C901 is not enabled as a permanent rule.** A separate check still reports signature capture (20), Pi prompt processing (16), calendar detail display (13), Pi progress summarization (13), browser token capture (12), HTTP retry processing (12), and SharePoint path resolution (11). Further changes need focused tests, especially for browser and process-control paths. The gateway complexity audit was left out of lint enforcement.
8. **An intermediate auth import defect was found by lint.** Phase 4 added a logger without a module-level `logging` import in `auth.py`. The then-current tests did not import that module. Phase 6 fixed it and added offline auth HTTP tests. Final tests and lint pass.
9. **Live performance and data checks are limited.** Only one before/after timing pair was taken per command. Live list data changed between runs. No throttling or expired-auth failure was deliberately induced on the live account. Those paths use offline tests.
10. **MCP schema capture was added in phase 6.** It was extracted from the original commit, not from the refactored code. It provides an additional unchanged-contract check beyond phase 0.

No phase was abandoned. The risky browser/process-control complexity work was left unchanged rather than forcing additional restructuring.
