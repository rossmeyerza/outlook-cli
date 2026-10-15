# outlook-cli refactor plan

## Goal

Reduce structural debt and speed up the slow Teams paths. No user-visible changes: CLI commands, flags, output formats, exit codes and MCP tool names/schemas stay the same.

## Ground rules

- Work in phases. Each phase ends with `.venv/bin/python -m pytest -q` green and one Conventional Commit. Do not start a phase while the previous one is red.
- **Never run `uv run`, `uv sync` or `uv lock`** in this repo. They replace `.venv` and break the global `outlook-cli`. Install with `uv pip install --python .venv/bin/python <pkg>`. Do not create `uv.lock`.
- Do not send email or Teams messages. Do not create calendar events, contacts or tasks against the live account. Live checks must be read-only (`auth status`, `mail unread -n 3`, `cal agenda -d 1`, `teams list -n 5`, `files sites`).
- Do not touch `outlook_draft/commands/gateway.py` except where imports must change.
- Keep functions under about 50 lines. Match the existing style. No new runtime dependencies.
- The `--yes` send guard on `draft send`, `mail send`, `teams send` must keep working (preview + exit 2 without `--yes`).
- The MCP server must never print to stdout, never `sys.exit`, never auto-reauth. Existing MCP tests guard this; keep them passing.
- Do not push.

## Phase 0: CLI safety net (test commit)

Add `tests/test_cli_smoke.py`. Build the parser and run handlers with faked clients (monkeypatch the client factories). Cover at least: `--help` for every top-level command and subcommand parses; `mail unread --json`, `mail search --json`, `cal agenda --json`, `task list --json`, `teams list --json`, `files sites --json` produce the same JSON shape as today; the three send commands exit 2 without `--yes`; an API error exits 1 with a red message. Capture current outputs first so later phases compare against them.

## Phase 1: shared runtime (refactor)

The command modules each define `_ctx`, `_console`, `_get_client`, `_print_json`, `_html_to_text`, `_format_datetime` and receive helpers through `args._mail_ctx` / `_calendar_ctx` / `_teams_ctx` / `_tasks_ctx` / `_contacts_ctx` dicts built in `cli.py`. Replace this with one module, for example `outlook_draft/runtime.py`, exposing `console`, `get_client()`, `get_graph_client()`, `print_json()`. Pure helpers stay in `formatting.py`. Remove the `build_ctx` functions and `_xxx_ctx` plumbing. Keep tests able to monkeypatch the client factories.

## Phase 2: split parser registration (refactor)

`cli.py` is 1,626 lines; `main()` is ~660 lines of argparse. Give each command area a `register(subparsers)` function in its module (mail, draft, calendar, tasks, contacts, teams, files, account/auth/config/mailbox, completion, gateway, mcp, signature). Move the draft and account/auth/mailbox handlers out of `cli.py` into `commands/drafts.py` and `commands/account.py`. Target: `cli.py` under ~300 lines. `argcomplete` and `completion.py` (including Nushell externs) must still work; the existing completion tests must pass.

## Phase 3: one HTTP stack (refactor)

- `commands/files.py` has its own `_GraphClient` duplicating `OutlookClient` request/retry/auth logic. Fold it into `OutlookClient` (Graph base URL) and keep a thin alias only if needed by the MCP layer. Update `mcp/clients.py` accordingly; the `NoReauthTokenManager` behaviour must not change.
- Replace `requests` usage in `auth.py` and `commands/signature.py` with `httpx`, and remove `requests` from `pyproject.toml` dependencies.
- Keep the existing 401/429/5xx retry semantics, including `Retry-After`.

## Phase 4: error handling (refactor)

- Add `CliError(message, exit_code=1)` in `errors.py`. Command handlers raise it (or let `OutlookAPIError` propagate) instead of calling `console.print(...); sys.exit(1)`. `main()` catches `CliError` and `OutlookDraftError` once, prints the red message to the console, and exits with the right code. The send-guard exit code 2 must stay.
- Review the 44 blind `except Exception` and 11 `try/except: pass` blocks (`ruff --select BLE001,S110,S112`). Narrow them to specific exceptions where the failure mode is known, or add a `log.debug(..., exc_info=True)`. Do not change behaviour of best-effort paths (for example browser-storage token extraction).

## Phase 5: Teams performance (perf)

`teams list`, `teams search` and the MCP `teams_list`/`teams_search` fetch members and message metadata one chat at a time (N+1 sequential Graph calls; `--scan 200` means hundreds of calls). Fetch per-chat members/metadata concurrently with `concurrent.futures.ThreadPoolExecutor(max_workers=8)`. Preserve ordering and sorting exactly. Make sure the shared `httpx.Client` is safe for this use (it is thread-safe for requests) and that 429 handling still backs off. Add a test that results keep their order. Report before/after wall time for `outlook-cli teams list -n 20` and `outlook-cli teams search <something> --scan 100` against the live account (read-only).

## Phase 6: lint (chore)

Add a `[tool.ruff]` section to `pyproject.toml` (target py312, line length matching the code, rules: E, F, I, UP, B, SIM, BLE). Apply safe auto-fixes. Reduce the worst `C901` offenders that remain after phases 1 to 4 only where it is low risk. `ruff check outlook_draft tests` should end clean or with a short, justified ignore list.

## Docs

Update `AGENTS.md` file layout and implementation notes to match the new module structure. Do not change user-facing command docs unless something actually changed.

## Definition of done

- All tests pass, including the new smoke tests. Report exact counts.
- `outlook-cli --help`, the read-only live checks and `pi mcp list` (outlook connected, 36 tools) all work.
- Phase 0 recorded outputs still match.
- One commit per phase, not pushed.
- Final report: commits, line counts before/after for `cli.py` and the command modules, removed dependencies, Teams timings, and any deviations or risks.
