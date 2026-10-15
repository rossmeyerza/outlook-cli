# outlook-cli MCP server: phase 1 plan

## Goal

Add a local stdio MCP server to this repo so Pi codemode scripts can call Outlook, Teams, calendar, contacts and files operations as typed tools that return structured JSON. The CLI stays the primary interface. The MCP server is a thin layer over `OutlookClient` and existing helpers.

Phase 1 is read-only plus drafts. Sending is out of scope (see "Out of scope").

## Entry point

- New subcommand: `outlook-cli mcp` (stdio transport). Honour the global `--account <name>` flag, so `outlook-cli --account wpp mcp` works.
- New package: `outlook_draft/mcp/` with `server.py` (FastMCP app, tool registration, `run()`), and one module per area: `mail.py`, `drafts.py`, `calendar.py`, `contacts.py`, `teams.py`, `files.py`, `tasks.py`, `auth.py`.
- Dependency: add `mcp>=1.10` (official Python SDK, `mcp.server.fastmcp.FastMCP`) to `[project].dependencies` in `pyproject.toml`.

## Hard constraints

1. **stdout is the protocol channel.** Nothing may print to stdout while the server runs. Rich `Console()` objects in `auth.py`, `cli.py` and command modules default to stdout. The MCP code path must not call any function that prints. Configure `logging` to stderr. Add a test that runs a tool call and asserts nothing was written to stdout except protocol frames (or at least that tool functions write nothing to stdout via `capsys`).
2. **No `sys.exit`.** Command modules call `sys.exit` on errors (mail 14, calendar 13, teams 12, tasks 7, contacts 4). MCP tools must never call them. Raise exceptions; FastMCP turns them into `isError` results with the message.
3. **No interactive or browser reauth.** Build clients with `TokenManager.get_token(auto_reauth=False)`. On `TokenNotFoundError` / `TokenExpiredError`, raise a clear error: `"<label> token expired or missing. Run 'outlook-cli auth' in a terminal, then retry."` Headless reauth prints an MFA number the user cannot see from a stdio server.
4. **Re-read tokens per call.** The server is long-lived and tokens rotate after `outlook-cli auth`. `TokenManager` already reloads on file mtime change; make sure clients are created per tool call (or that reload is triggered) so a fresh `outlook-cli auth` takes effect without restarting the server.
5. **CLI behaviour must not change.** Existing tests must pass. Do not change CLI output formats.
6. **Return real IDs.** Tools take and return full Outlook/Graph IDs. Do not use the numeric index caches (`cache.py`) from MCP tools, and do not write to them.
7. **Return structured data.** Tool functions return typed dicts/lists (TypedDict or pydantic models) so FastMCP emits `structuredContent`. Keep payloads lean: summaries for list/search tools, full content only for read/get tools. Convert HTML bodies to text by default, with an `include_html: bool = False` option on read tools.

## Refactor needed (minimal)

The command modules mix data shaping with Rich printing and get helpers through an `args._xxx_ctx` dict. Extract only the pure helpers the MCP layer needs into importable functions without `args`:

- `cli._html_to_text`, `cli._format_datetime`, `cli._compose_email_html`, `cli._load_signature`, `cli._text_to_html`, `cli._parse_recipients` → move to a new `outlook_draft/formatting.py` (or similar) and re-import them in `cli.py` under the same names so nothing else breaks.
- `commands/mail.py`: `_message_summary`, `_collect_message_links` → make pure (take `html_to_text`/`format_datetime` as plain functions, not `args`).
- `commands/calendar.py`: `_event_summary`, `_event_when`, `_event_recurrence`, `_availability_summary`, `_parse_event_datetimes` (if needed for create) → pure.
- `commands/teams.py`: `_chat_title`, `_teams_sender`, `_chat_match_reasons`, `_latest_received_message_time`, `_is_received_user_message`, `find_self_chat`, `_collect_message_links_and_attachments` → pure where they take `args`.
- `commands/files.py`: `_GraphClient`, `_resolve_site_drives`, `_match_library`, `_resolve_site`, `_find_sp_item_by_path`, `_item_summary`, `_drive_summary`, `_site_summary` are mostly pure already. Check they do not print or exit; if they do, raise instead and keep the CLI handlers catching and printing.

Keep the existing CLI handlers as thin wrappers that call the extracted helpers. Do not rewrite handlers that the MCP layer does not need.

## Phase 1 tools

Names are snake_case. All IDs are full IDs. Sensible defaults for counts; cap them.

Mail
- `mail_search(query: str, count: int = 20)` → list of message summaries (id, subject, from, received, isRead, hasAttachments, preview, webLink if available)
- `mail_unread(count: int = 20)` → same shape
- `mail_read(message_id: str, include_html: bool = False, max_chars: int = 20000)` → full message (headers, recipients, body text, attachments list, links)
- `mail_attachments(message_id: str)` → list (id, name, contentType, size, isInline)
- `mail_download_attachments(message_id: str, dest_dir: str, attachment_ids: list[str] | None = None, overwrite: bool = False)` → list of written paths. `dest_dir` must be an absolute path; create it if missing.
- `mail_links(message_id: str)` → SharePoint/OneDrive links found in the body
- `mail_folders()` → list
- `mail_mark_read(message_id: str, is_read: bool = True)`
- `mail_move(message_id: str, folder: str)` (folder name or id) and `mail_archive(message_id: str)`

Drafts
- `draft_create(to: list[str], subject: str, body: str, cc=None, bcc=None, html: bool = False, importance: str = "Normal")` → draft id + webLink. Must apply the same Aptos HTML wrapping and new-message signature as `outlook-cli draft create`.
- `draft_reply(message_id: str, body: str, reply_all: bool = False, html: bool = False)` → same, using the reply signature, matching `outlook-cli draft reply`.
- `draft_list(count: int = 20)`, `draft_get(draft_id: str)`, `draft_delete(draft_id: str)`

Calendar
- `cal_agenda(days: int = 7, count: int = 50)` → event summaries
- `cal_get(event_id: str)` → full event incl. recurrence
- `cal_availability(attendees: list[str], start: str, end: str, ...)` and `cal_find_time(attendees: list[str], start: str, end: str, duration_minutes: int = 30, ...)` mirroring the CLI options
- `cal_create(subject, start, end, attendees=None, location=None, body=None, online_meeting=False, ...)`: mirror `cal create` options. Note: creating an event with attendees sends invitations. Include that in the tool description so the model knows.
- `cal_respond(event_id: str, response: Literal["accept","tentative","decline"], comment: str = "", send_response: bool = True)`: also emails the organizer when `send_response` is true. Say so in the description.

Contacts
- `contact_search(query: str, count: int = 10)`

Tasks
- `task_list(count: int = 50)`, `task_create(subject: str)`, `task_complete(task_id: str)`

Teams (read-only)
- `teams_list(count: int = 20)` → chats sorted the same way as `teams list`
- `teams_search(query: str, scan: int = 200)`
- `teams_messages(chat_id: str, count: int = 20)` → messages with sender, time, text body
- `teams_attachments(chat_id: str, count: int = 50)` and `teams_download_attachments(...)` mirroring the CLI

Files
- `files_sites()`, `files_libraries(site: str)`
- `files_list(path: str = "", site: str | None = None, library: str | None = None)`
- `files_search(query: str, site: str | None = None, library: str | None = None, count: int = 20)`
- `files_download(item: str, dest: str, site=None, library=None, overwrite=False)` where `item` is a path or drive item id, matching `files download`

Auth
- `auth_status()` → token label, domain, valid/expired, minutes left. Read-only. Never triggers reauth.

## Out of scope for phase 1

- Sending: `draft send`, `mail send`, `teams send`. These stay CLI-only behind `--yes`. Do not register them as MCP tools.
- Destructive or account-level operations: `cal delete/cancel/update`, `files upload/mkdir/rename/move`, `contact create/update`, `mailbox update`, account switching, `auth` (interactive), signature fetch, gateway.

## Server metadata

- Server name: `outlook`. Server instructions (first line used as the description in Pi): "Outlook mail, drafts, calendar, contacts, tasks, Teams (read-only) and OneDrive/SharePoint files for Ross's Microsoft 365 account. Sending is not available here; create drafts."
- Tool descriptions: one or two sentences, state side effects plainly.

## Registration

Add to `~/.pi/agent/mcp.json` under `mcpServers` (keep existing entries untouched):

```json
"outlook": {
  "command": "/home/ross/.local/lib/outlook-draft-cli/.venv/bin/outlook-cli",
  "args": ["mcp"],
  "description": "Outlook mail, drafts, calendar, contacts, tasks, Teams (read-only) and OneDrive/SharePoint files",
  "timeout": 120
}
```

Default exposure is `codemode`, which is what we want. Verify with `pi mcp list`.

## Tests

- Unit tests per tool module using a recording/fake client (follow the `RecordingClient` pattern in `tests/test_outlook_client_mail.py`). Cover: argument mapping to client calls, summary shaping, HTML-to-text, draft signature application, errors raised (not `sys.exit`) on API and token errors.
- A test that the expired-token path raises the "Run 'outlook-cli auth'" error and does not call `run_reauth`.
- A test that the send tools are not registered.
- A stdout-cleanliness test for at least one tool per module.
- An in-process MCP smoke test: use the SDK's in-memory client/session (`mcp.shared.memory` / `create_connected_server_and_client_session` or equivalent in the installed version) to `list_tools` and call one tool with a faked client.

## Environment rules for the implementer

- The installed CLI runs from `/home/ross/.local/lib/outlook-draft-cli/.venv` (symlinked from `~/.local/bin/outlook-cli`).
- **Do not use `uv run`, `uv sync` or `uv lock` in this repo.** They replace `.venv` and break the global `outlook-cli`. Install packages with `uv pip install --python .venv/bin/python <pkg>` and run tests with `.venv/bin/python -m pytest -q`.
- Do not create or commit `uv.lock`.
- Do not send any email or Teams message, and do not create calendar events with attendees, during testing. Live read-only checks (for example `mail_unread`, `cal_agenda`, `auth_status`) are fine. Draft creation live tests must go to Ross's own address only and must be deleted afterwards.

## Docs

- `README.md` and `AGENTS.md`: short "MCP server" section with the run command, registration snippet, tool list, and the no-send rule.
- `skills/outlook-cli/SKILL.md`: short note that in codemode the `mcp__outlook` tools are preferred for reading and drafting, and sending stays CLI-only with `--yes`.

## Definition of done

- `.venv/bin/python -m pytest -q` passes, including new tests.
- `outlook-cli --help` and existing commands behave as before.
- `pi mcp list` shows `outlook` connected with the phase 1 tools.
- A live codemode smoke check works: `auth_status`, `mail_unread(count=3)`, `cal_agenda(days=1)`.
- Atomic Conventional Commits (for example: refactor helpers, feat mcp server, test, docs). Do not push.
