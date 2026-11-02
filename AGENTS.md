# outlook-cli

## Overview

CLI for managing Outlook email, drafts, calendar, tasks, contacts, Teams browsing, signatures, and Microsoft 365 Copilot questions via the Outlook REST API v2.0, Microsoft Graph, and the internal Copilot StreamHub protocol. Globally available as `outlook-cli`. Repo: `rossmeyerza/outlook-cli`.

## Quick Reference

```bash
# Search and read emails
outlook-cli mail unread                    # list unread emails
outlook-cli mail unread --json             # JSON unread emails
outlook-cli mail search <query>            # search emails by keyword
outlook-cli mail search colgate -n 10      # with result limit
outlook-cli mail search <query> --json     # JSON search results
outlook-cli mail read <n>                  # read email by index from last unread/search
outlook-cli mail read <n> --json           # full message JSON
outlook-cli mail read <full-id>            # read by Outlook message ID
outlook-cli mail mark-read <n>             # mark email as read
outlook-cli mail mark-unread <n>           # mark email as unread
outlook-cli mail archive <n>               # archive email
outlook-cli mail folders                   # list mail folders
outlook-cli mail folders --json            # JSON folders
outlook-cli mail move <n> --folder <folder> # move email
outlook-cli mail attachments <n>           # list attachments
outlook-cli mail attachments <n> --json    # JSON attachments
outlook-cli mail download-attachments <n>  # download attachments

# Create drafts
outlook-cli draft create --to <email> -s "Subject" -b "Body text"
outlook-cli draft create --to <email> -s "Subject" -f ./email.html --html
outlook-cli draft reply <message-ref> -b "Thanks, I'll come back to you"
outlook-cli draft reply <message-id> --reply-all -f ./reply.html --html

# Manage drafts
outlook-cli draft list                     # list all drafts
outlook-cli draft show <n>                 # show draft by index (1, 2, ...)
outlook-cli draft delete <n>               # delete draft by index

# Manage calendar
outlook-cli cal agenda                     # list upcoming events (next 7 days)
outlook-cli cal agenda -d 14               # list next 14 days
outlook-cli cal agenda --plain             # plain text agenda
outlook-cli cal agenda --json              # JSON agenda
outlook-cli cal agenda --table             # explicit table agenda
outlook-cli cal show <n>                   # show event details by index
outlook-cli cal show <n> --json            # JSON event details with recurrence metadata
outlook-cli cal create "Subj" "Start" "End" # create event (e.g. "2026-04-10 14:00")
outlook-cli cal rooms                      # find rooms if supported by tenant/token
outlook-cli cal availability --attendee <email> --start "..." --end "..."
outlook-cli cal find-time --attendee <email> --start "..." --end "..."
outlook-cli cal update <n> --start "..." --end "..." # update event fields
outlook-cli cal delete <n>                 # delete event by cached index or full ID
outlook-cli cal accept <n>                 # accept event invitation
outlook-cli cal tentative <n>              # tentatively accept event invitation
outlook-cli cal decline <n>                # decline event invitation
outlook-cli cal cancel <n>                 # cancel event you organize

# Manage tasks
outlook-cli task create "<name>"           # create a new task
outlook-cli task list                      # list active tasks
outlook-cli task list --json               # JSON active tasks
outlook-cli task update <n> --due "..." --importance High
outlook-cli task complete <n>              # mark task as done
outlook-cli task delete <n>                # delete task

# Search contacts
outlook-cli contact search <query>         # search by name or email
outlook-cli contact search ross -n 20      # with result limit
outlook-cli contact create --name <name> --email <email>
outlook-cli contact update <n> --company <company>

# Browse Teams chats (read-only)
outlook-cli teams list -n 20               # list chats
outlook-cli teams search <query> --scan 200 # search chat names/participants
outlook-cli teams self                     # check access to Teams self-chat
outlook-cli teams show <chat-ref>          # show chat details
outlook-cli teams messages <chat-ref> -n 20 # read messages

# Teams gateway to Pi/Marlow
outlook-cli gateway start --self-chat      # monitor Teams self-chat via 48:notes
outlook-cli gateway start --chat-id "19:..." --trigger "@Marlow" --poll 30
outlook-cli gateway start --self-chat --model sonnet:high
outlook-cli gateway status                 # show chat, state, workspace, sessions, model
outlook-cli gateway stop

# Files (OneDrive and SharePoint)
outlook-cli files sites                    # list SharePoint sites you're a member of
outlook-cli files libraries --site "Tesco" # list SharePoint document libraries
outlook-cli files list                     # list OneDrive root
outlook-cli files list "Documents/Reports" # list OneDrive subfolder
outlook-cli files list --site "Tesco"      # list SharePoint document libraries
outlook-cli files list --site "Tesco" --library "Documents" "General" # list SharePoint folder
outlook-cli files list --site "Tesco" --library "Documents" "General" --links # include browser links
outlook-cli files list --site "Tesco" --library "Documents" "General" --json  # JSON paths + webUrl
outlook-cli files search budget --site "Tesco" -n 20 # search SharePoint libraries
outlook-cli files search budget --site "Tesco" --library "Documents" --links # search one library with links
outlook-cli files search budget --site "Tesco" --json # machine-readable paths + webUrl
outlook-cli files upload ./file.pdf "Documents"            # upload to OneDrive
outlook-cli files upload ./file.pdf "General" --site "Tesco" --library "Documents" # upload to SharePoint
outlook-cli files download <drive-item-id> --site "Tesco" --library "Documents" # download by search/list ID
outlook-cli files mkdir "Documents/Q2"     # create OneDrive folder
outlook-cli files mkdir "Q2" --site "Tesco" --library "Documents" # create SharePoint folder
outlook-cli files rename "Documents/old.pdf" "new.pdf"    # rename
outlook-cli files move "Documents/file.pdf" "Archive"     # move

# Microsoft 365 Copilot (internal protocol, can change)
outlook-cli copilot ask "<prompt>"         # ask Copilot (work chat)
outlook-cli copilot ask "Summarise" --file <sharepoint-url> # file as context (repeatable)
outlook-cli copilot ask "Summarise" --file "General/Plan.docx" --site "Tesco" --library "Documents"
outlook-cli copilot ask "And the budget?" --continue      # follow-up in the last conversation
outlook-cli copilot ask "..." --conversation <id> --json  # JSON: text, sources, conversationId

# Copilot in the live Office web editor (experimental, edits the real file)
outlook-cli copilot extension install      # native host + unpacked extension folder
outlook-cli copilot app status --json      # host, extension version, Office tabs
outlook-cli copilot app inspect --tab-id <id> --json  # read-only UI/frame diagnostics
outlook-cli copilot app ask "<prompt>" --file <url|id|path> --json [--open] [--new-chat] [--wait-saved] [--timeout 300]

# MCP server (stdio, confirmed sending)
outlook-cli mcp
outlook-cli --account wpp mcp

# Shell completion
outlook-cli completion bash                # print bash completion script
outlook-cli completion zsh                 # print zsh completion script
outlook-cli completion fish                # print fish completion script
outlook-cli completion powershell          # print PowerShell completion script
outlook-cli completion nushell             # print Nushell extern definitions

# Signatures
outlook-cli signature fetch                # fetch new + reply signatures from OWA (no MFA after first auth)
outlook-cli signature fetch --headed       # run with visible browser if headless fails

# Auth and config
outlook-cli account add wpp --email <email> --password <password> --switch # create/switch account profile
outlook-cli account add ikea --email <email> --password <password>        # create another profile
outlook-cli account list                   # list account profiles
outlook-cli account current                # show active account/profile paths
outlook-cli account switch ikea            # switch default active account
outlook-cli account switch default         # return to legacy/global .env
outlook-cli --account wpp mail unread      # one-command account override
outlook-cli auth                           # force headless re-authentication
outlook-cli auth status                    # show token status
outlook-cli auth clear                     # delete local token cache
outlook-cli auth scopes                    # list safe token metadata/scopes
outlook-cli mailbox show                   # show mailbox settings
outlook-cli mailbox update --timezone "GMT Standard Time"
outlook-cli config check                   # validate local config without printing secrets
```

## Key Details

- **Progress spinners**: Interactive commands render Rich spinners on stderr while waiting on Microsoft APIs. Use global `--no-spinner` or `OUTLOOK_CLI_NO_SPINNER=1` to disable them; JSON stdout remains parseable.
- **Account profiles**: Multiple credentials are supported with `account add/list/current/switch/remove` and the global `--account <name>` override. Profile env files live at `~/.config/outlook-cli/accounts/<name>.env`; profile data lives at `~/.local/share/outlook-cli/accounts/<name>/`. Each profile has isolated tokens, browser state, caches, signatures, and gateway state. `account switch default` returns to the legacy/global `.env` and `~/.local/share/outlook-cli/session_state/` layout, which appears in `account list` as `(default)`.
- **Email search**: `mail unread` and `mail search` results are cached to disk so `mail read <n>` works in a separate invocation. Use `--json` for agent-readable summaries, `--table` for explicit human tables.
- **Email reading**: `mail read` strips HTML to plain text for console display. Truncates at 3000 chars.
- **Calendar**: `cal agenda` results are cached to disk so `cal show <n>`, `cal update <n>`, `cal delete <n>`, `cal accept <n>`, `cal tentative <n>`, `cal decline <n>`, and `cal cancel <n>` work in a separate invocation. Supports table, plain, and JSON agenda output. `cal show --json` includes recurrence metadata. Calendar creation/update uses `LOCAL_TIMEZONE` and `OUTLOOK_TIMEZONE` from `.env`. Room discovery is tenant/token dependent; availability and find-time work with known room/user email addresses.
- **Tasks**: `task list` caches to disk so `task complete <n>` works in a separate invocation.
- **Recipients**: `--to`, `--cc`, `--bcc` are repeatable and accept comma-separated addresses.
- **Body source**: `--body` for inline text, `--body-file` for file. Add `--html` when the provided body is already HTML.
- **Reply drafts**: `draft reply` creates a draft tied to an existing message. Use a cached mail index from `mail unread` / `mail search`, a cached ID suffix, or a full message ID.
- **Draft formatting**: New and reply drafts are always saved as HTML, use Aptos for the message body, and append the saved signature HTML from `SIGNATURE_NEW_FILE` or `SIGNATURE_REPLY_FILE`. Defaults are `signature-new.html` and `signature-reply.html`.
- **Importance**: `--importance Low|Normal|High` (default: Normal).
- **Contacts**: `contact search` searches org directory and recent email contacts. Returns name, email, and type.
- **Teams**: `teams list`, `teams search`, `teams show`, and `teams messages` browse Teams chats and messages via Microsoft Graph. `list_teams_messages` asks Graph for `createdDateTime desc` (without it, long chats returned only their oldest 50 messages) and follows `@odata.nextLink`, so `--count` above 50 works; callers sort for display. `teams search <query>` scans a bounded set of chats, matches chat topics and participants, and caches the matching chats so numeric refs work in a separate invocation. `teams list` sorts by latest received user message, ignoring system events and self messages where identifiable. Sending is enabled but guarded: `draft send <n>`, `mail send --to ... -s ... -b ...` and `teams send <chat> <msg>` print a preview and send nothing unless `--yes` is passed. Prefer drafts for email and only pass `--yes` when Ross has explicitly asked to send.
- **Teams self-chat**: `teams self` checks access to the Microsoft Graph special self-chat thread, `48:notes`. This is the real Teams "chat with yourself" stream and is different from normal one-person `/me/chats` entries.
- **Teams gateway**: `gateway start --self-chat` watches `48:notes` for `@Marlow`, sends prompts to `pi --mode rpc`, and posts responses back into Teams. The gateway posts a short `...` receipt, soft-deletes it before the final response where Graph allows it, and can surface compact Pi tool progress. If the Graph token expires and headless reauth fails, the gateway records the auth error and exits instead of retrying every poll interval.
- **Gateway state**: Runtime state is outside git under the active account's `session_state/` directory. For legacy/global config this is `~/.local/share/outlook-cli/session_state/`; for profiles this is `~/.local/share/outlook-cli/accounts/<name>/session_state/`. Pi sessions live under `session_state/gateway_sessions/<chat-hash>/`; per-chat workspaces live under the active account's `gateway_workspaces/<chat-hash>/`. These are local data/state, not tracked or pushed.
- **Gateway model controls**: Start with `outlook-cli gateway start --self-chat --model <model> [--provider <provider>] [--thinking <level>]`. Inside Teams use `@Marlow !model`, `@Marlow !model help`, `@Marlow !model list [search]`, `@Marlow !model <model>`, `@Marlow !model --provider <provider> --model <model> --thinking <level>`, or `@Marlow !model reset`. Short names such as `sonnet` rely on Pi fuzzy/pattern matching; use `!model list sonnet` and an exact model ID when repeatability matters.
- **Gateway file publishing**: Pi must not send Teams messages directly. To send generated files back to Teams, create files under the per-chat workspace and write `.marlow-export.json` with JSON like `{"files":["report.html"],"message":"Created the report."}`. The gateway validates relative paths, uploads supported file types to OneDrive under `Outlook CLI/Gateway/<chat-hash>/`, creates organization view links, posts them, and removes the manifest.
- **Gateway commands in Teams**: Use `@Marlow !help`, `@Marlow !commands`, or `@Marlow !help <command>` to list commands. Current commands include `!status`, `!new`, `!reset`, `!model`, `!pause`, `!resume`, `!tools`, `!files`, `!send`, and `!logs`.
- **Draft references**: `draft show` and `draft delete` accept a numeric index from `draft list`, a partial ID suffix, or a full ID.
- **Files**: `files sites` lists SharePoint sites via M365 group membership. `files libraries --site <name>` lists SharePoint document libraries. `files list --site <name>` shows libraries; use `--library <name>` to browse or mutate a specific document library. `files search <query> --site <name>` searches across document libraries, or one library with `--library`. Use `--links` on `files list`, `files search`, or `files libraries` to include SharePoint/OneDrive browser links in table output. Use `--json` for machine-readable `id`, `path`, `parentPath`, `library`, and `webUrl` fields. Browser links open the file in Microsoft 365; `files download` downloads file bytes to local disk and accepts either a relative path or a drive item ID from search/list JSON. `--site` and `--library` do case-insensitive partial name matches. Without `--site`, operations target personal OneDrive. Uploads under 4 MB use a single PUT; larger files use chunked upload sessions. Uses the Microsoft Graph token.
- **Copilot**: `copilot ask` sends a prompt to Microsoft 365 Copilot over the StreamHub WebSocket of the Copilot web app (`wss://substrate.office.com/m365Copilot/StreamHub/<oid>@<tid>`). This is not a public API, and Microsoft can change it without notice. `--file` takes a SharePoint/OneDrive URL (such as `webUrl` from `files search --json`), a drive item ID, or a path. Paths and IDs are in OneDrive unless `--site` is given. The CLI reads the file's `sharepointIds` with Graph and attaches it as a `File` message annotation with the ID `SPO_<base64url(siteId,webId,listId)>_<driveItemId>`. Other ID formats make Copilot fail to open the file. `--continue` uses the last conversation ID from the account cache (`copilot-conversation.json`). Copilot keeps the files of earlier turns. The answer text has citation markers (`\ue200...\ue201`) and entity tags (`<Person>...</Person>`) removed. Sources come from the `searchResultsMap` of the citation message. Copilot is not read-only: its own tools can send email and Teams messages, change calendars, and delete or share SharePoint files. The collector records each tool call from the message `invocation` field (`toolCalls`, without `api_tool_list_resources` and hints). It stops at a confirmation, input or sign-in request and reports it as `pendingConfirmations`; it never confirms. Sources are only the search results that the answer cites (`turnNsearchM` points to `searchResultsMap["M"]`). See `docs/copilot-tools.md` for the self-reported tool list, the tools with side effects and the test results.
- **Shell completion**: `completion <bash|zsh|fish|powershell|nushell>` prints a completion script. Bash/zsh/fish/PowerShell use `argcomplete`; Nushell emits static extern definitions.
- **Signatures**: `signature fetch` opens a headless browser with the active account's saved OWA browser session (`session_state/browser_state.json`), intercepts the `OutlookCloudSettings/settings/account` API responses OWA fires on load, and saves the active new-message and reply signatures. No MFA after the first `auth`. On a fresh install, `auth` and `signature fetch` run automatically.
- **Auth**: Built into this repo via `outlook_draft/auth.py`. Outlook features use the Outlook token, Teams uses the Microsoft Graph token, and `copilot` uses the Copilot token (audience `https://substrate.office.com/sydney`, key `substrate.office.com/sydney` in `tokens.json`). After the Outlook login, `auth` opens `https://m365.cloud.microsoft/chat` in the same browser context, clicks the "Do you trust ...?" page, and captures the Copilot token from the request headers or the StreamHub `access_token` parameter. If the account has no Copilot licence, `auth` prints a warning and continues. The Copilot token lasts about an hour. `CopilotTokenManager` refreshes it from the saved `browser_state.json` with no MFA, and falls back to the full login only in the CLI. `auth` saves both API tokens and the full browser session state for the active account. Run `outlook-cli auth` if expired, or `outlook-cli auth --headed` for a visible browser. Existing single-account installs keep using `~/.config/outlook-cli/.env` and do not need to re-authenticate unless you switch to a new profile or tokens are expired.

- **Outlook Web login**: Login opens `outlook.cloud.microsoft` directly. It accepts both `outlook.office.com` and `outlook.cloud.microsoft` after sign-in. Outlook tokens from the new host use the existing `outlook.office.com` cache key. Each captured token is saved during login. Existing tokens for other services are kept. The cache file is replaced only after the new contents are written. If a later login step fails, tokens already saved remain available.

## MCP Server

`outlook-cli mcp` runs a stdio MCP server named `outlook` (package `outlook_draft/mcp/`, built on `mcp.server.fastmcp.FastMCP`, SDK pinned `mcp>=1.10,<2`). It is registered in `~/.pi/agent/mcp.json` and appears as `mcp__outlook` in Pi codemode.

```json
"outlook": {
  "command": "/home/ross/.local/lib/outlook-draft-cli/.venv/bin/outlook-cli",
  "args": ["mcp"],
  "description": "Outlook mail, drafts, calendar, contacts, tasks, Teams and files, with confirmed sending",
  "timeout": 120
}
```

Tools (all IDs are full Outlook or Graph IDs; numeric index caches are not used):

- Mail: `mail_search`, `mail_unread`, `mail_read`, `mail_attachments`, `mail_download_attachments`, `mail_links`, `mail_folders`, `mail_mark_read`, `mail_move`, `mail_archive`
- Drafts: `draft_create`, `draft_reply`, `draft_list`, `draft_get`, `draft_delete` (draft creation does not send; `draft_delete` refuses non-drafts)
- Calendar: `cal_agenda`, `cal_get`, `cal_availability`, `cal_find_time`, `cal_create`, `cal_respond`
- Contacts: `contact_search`
- Tasks: `task_list`, `task_create`, `task_complete`
- Teams (read-only): `teams_list`, `teams_search`, `teams_messages`, `teams_attachments`, `teams_download_attachments`
- Files: `files_sites`, `files_libraries`, `files_list`, `files_search`, `files_download`
- Copilot: `copilot_ask` (prompt, optional `files`, `site`, `library`, `conversation_id`). Not read-only: registered with `readOnlyHint=false` and `destructiveHint=true`.
- Auth: `auth_status` (read-only)
- Sending: `mail_send`, `draft_send`, `teams_send` (preview by default; requires `confirm=true` to send)

Side effects to know about: `cal_create` with attendees emails invitations immediately, and `cal_respond` with `send_response` true emails the organizer.

Rules:

- Send tools `mail_send`, `draft_send` and `teams_send` preview by default. Set `confirm=true` only when the user explicitly requests that send. Do not automatically retry an uncertain send result. CLI sending still requires `--yes`.
- Destructive and account-level commands are not exposed (calendar update/delete/cancel, file upload/move/rename/mkdir, contact create/update, mailbox update, interactive auth, signature fetch, gateway).
- The server never runs the login. If a token is expired, tools return "Run 'outlook-cli auth' in a terminal, then retry." Tokens are re-read on every call, so a fresh `outlook-cli auth` works without restarting the server.
- One exception: `copilot_ask` refreshes the Copilot token from the saved browser session (no password, no MFA, no stdout output). If the saved session has expired, the tool returns the auth message.
- stdout is the protocol channel. Logs go to stderr.

Implementation notes:

- Tools build a client per call through `outlook_draft/mcp/clients.py`. `NoReauthTokenManager` never runs `run_reauth`, so a stale token raises `McpAuthError` instead of starting headless auth.
- `mcp/clients.py` `copilot_client()` uses `CopilotTokenManager(full_login_fallback=False)`. `auth.capture_copilot_token` runs Playwright in a worker thread, because the Playwright sync API fails inside the MCP asyncio loop.
- Tool code must not print to stdout and must not call `sys.exit`. Raise exceptions; FastMCP returns them as `isError` results.
- Pure helpers live in `outlook_draft/formatting.py`. CLI factories and output helpers live in `runtime.py`. Cached reference resolution lives in `references.py`.
- Command modules register their parsers through `register(subparsers)`. `cli.py` handles account selection, dispatch and `CliError`/`OutlookDraftError` rendering. Gateway parser registration lives in `commands/gateway_parser.py` so the daemon module stays unchanged.
- `OutlookClient` owns Outlook, Teams and file HTTP operations. `commands/files.py` keeps a thin `_GraphClient` compatibility constructor with the original file timeout and headers. MCP factories use `OutlookClient` directly and keep `NoReauthTokenManager`.
- Teams detail reads use `concurrency.ordered_map` with eight workers. Results stay in input order. Searches submit batches of eight, so an early match can leave at most seven extra reads in progress.
- Ruff uses the rules in `pyproject.toml`. Long literal lines are allowed. The gateway module is excluded because the refactor plan prohibits edits.
- Do not run `uv run`, `uv sync` or `uv lock` in this repo. They replace `.venv` and break the global `outlook-cli`. Use `uv pip install --python .venv/bin/python <pkg>` and `.venv/bin/python -m pytest -q`.

## Experimental Office-app Copilot bridge

- `copilot extension install` copies `extension/` to the user data directory and writes the Chrome native host manifest and launcher. The user loads the folder in `chrome://extensions`. The fixed extension ID is `lijjlmfmgilgjikalapbblbadhabmnae`.
- `copilot app status` lists Office editor tabs. `copilot app ask` takes one resolved file and can edit it in place. Excel and PowerPoint have live acceptance checks. Word support is experimental. Extension 0.3.2 passed read-only and single-paragraph edit checks on a disposable document. CLI handlers are in `commands/copilot_app.py`; shared logic and the native host are in `copilot_bridge/`.
- `--file` accepts a full SharePoint/OneDrive browser URL, a drive item ID or a path. Bare paths resolve in the user's own OneDrive only. For shared files (SharePoint or another person's OneDrive), pass the browser URL. The tab is matched by `listItemUniqueId` (the `sourcedoc` GUID), not by title.
- Pane readiness: the pane is "presented" only when the frame that embeds it reports the iframe visible (`framePresented` in `sw.js`). Word keeps the pane iframe mounted, with a ready input, while the drawer is closed, so the pane cannot judge this itself. Do not require a visible chain to the top frame: the SharePoint host page reports the Word editor iframe as hidden while it is on screen. `diagnostics.visibilityReason` uses computed CSS and geometry only; `aria-hidden` and `inert` are not visual hiding. Parent frames learn child URLs through a `frame_probe`/`frame_identity` postMessage handshake, because navigated iframes do not expose their URL in `src`.
- Reload rules: after changing `sw.js`, `controller.js` or `diagnostics.js`, reload the extension; the worker reinjects controllers into open tabs. After changing `hook.js` or `adapters/`, also refresh the Office tab, because the main-world hook loads at `document_start`. Check `controllerVersion` and `hookRevision` in `app status` before trusting a live test. The worker reads `chrome.runtime.lastError` on host disconnect and reconnects the transport; it never resends a prompt.
- The browser-wide socket is `config.DATA_DIR/copilot-bridge.sock`, not the active account directory. It uses 0600 permissions, a 0700 parent and Linux `SO_PEERCRED`. The host starts through `python -m outlook_draft.copilot_bridge`. stdout is Chrome's length-framed protocol channel. Logs contain no prompts, document frames or tokens.
- The extension hook runs in the editor's main world. Adapters export only answer text and short action events. Never export raw AugLoop frames: PowerPoint sends the whole document model. Never tag or change the editor DOM. Only the pane prompt control receives text.
- The controller clicks only Copilot, Send and New chat controls. Never add approval controls to its allowlist. Version 0.2.0 has pre-send checks and a passive observer for visible permission dialogs. Live verification and inline/backend approval formats are still incomplete. A pending approval can time out. Timeout/cancel stops collection, not the Office agent. Never retry automatically. Check the document, then reload the tab to clear its uncertain state.
- `copilot_app` is a destructive, non-idempotent MCP tool. Shared-file edits are visible to others. Excel and Word `applied` is null; script source only supplies a `mayWrite` hint. PowerPoint `applied` uses slide activity and client responses, not independent verification. `--wait-saved` uses a later Graph timestamp; other writers can cause it.
- `copilot app inspect --tab-id ID` is read-only. It returns allowlisted UI control states and schema/property names with counts, not document/prompt values, tokens or raw frames. Use it for pane-opening and Word protocol investigation. Word requires hook revision 3 and a refreshed tab after loading extension 0.3.0. Word uses the shared OfficeAgent adapter in `adapters/excel.js`, not startup WordCopilotOutput messages. Never export Word function arguments.
- Verify with `.venv/bin/python -m pytest -q` and `node --test extension/tests/*.test.js`. These use synthetic fixtures. Complete the live checklist in `docs/copilot-app-plan.md` before claiming release acceptance. The spike's checks do not verify this implementation.

### Handoff state (2026-10-23, second session)

The bridge work and the two fixes below are committed on `main` and are not pushed. Last local run: 320 Python passed (1 skipped), 35 JS passed, Ruff clean.

Chrome runs extension 0.3.2 (controller v5, hook revision 3). All Office tabs show controller v5.

Verified live on disposable files:
- Excel `Book 1.xlsx`: read-only summary, then a single-cell edit (Sheet1!E1). The saved file matched. A Graph timestamp advanced before the download contained the edit, so `--wait-saved` alone is not proof.
- PowerPoint `Presentation 4.pptx`: read titles, title-only edit, slide insertion. Saved OOXML compared and slides rendered.
- Word `dummy_document.docx` (item ID `01LT2YGQYBBGUCIETYMVA2PBZIDTQ5ZQMH`): read-only prompt with the drawer open, read-only prompt with the drawer closed (automatic opening worked), and an edit that added the paragraph `Bridge live check 2026-10-23.` at the end. The saved `.docx` was compared with a copy from before the edit; see `docs/copilot-app-plan.md`.

Fixed in this session:
- `auth.py` writes login progress and the MFA number to stderr. Automatic reauth no longer breaks `--json` output or the MCP stdio channel. With `--json`, `cli.py` writes errors to stderr.
- `files download <drive-item-id>` without `--site` now resolves the OneDrive item by ID, then falls back to the path.

Next steps:
1. Live approval detection (no real approval recorded in any app).
2. The `--new-chat` live check.
3. Excel new-table generation.
4. Cold-start reliability of PowerPoint pane opening.
5. A live `copilot_app` MCP call.

Test rules: use only disposable files. Never test on the open working documents (`Innovative-Consulting-David-Update.pptx`, the DELL deck, the Infosum document). A failed or timed-out request marks the tab uncertain; ask Ross to check the document and refresh the tab instead of retrying. Live evidence files are in `/tmp/copilot-live-checks/`. Chrome's `--remote-debugging-port=9222` flag is set but nothing listens on it, so the extension's `inspect` is the only page diagnostic channel.

## File Layout

```
/home/ross/.local/lib/outlook-cli/
  .env                          # MS_EMAIL, MS_PASSWORD, timezone, signature path config, ignored by git
  .env.example                  # Example local config
  ~/.config/outlook-cli/accounts/<name>.env # Optional per-account profile env files
  install.sh                    # Installer/updater script (served at outlook-cli.21436587.xyz)
  session_state/                # Local token + browser session cache, ignored by git
    tokens.json                 # API tokens
    browser_state.json          # Playwright browser session (used by signature fetch)
  ~/.local/share/outlook-cli/accounts/<name>/ # Per-profile data root
    session_state/tokens.json    # Per-profile API tokens
    session_state/browser_state.json # Per-profile Playwright browser session
    cache/                       # Per-profile numeric ref caches
    signature-new.html           # Per-profile new-message signature default
    signature-reply.html         # Per-profile reply signature default
  outlook_draft/
    auth.py                     # Playwright auth, token capture, browser session save
    calendar_time.py            # Calendar timezone parsing and Outlook headers
    cli.py                      # CLI entry point and dispatch
    runtime.py                  # CLI console and client factories
    references.py               # Cached CLI reference resolution
    parser.py                   # Shared argparse options
    concurrency.py              # Bounded, ordered Graph reads
    formatting.py               # Pure text/HTML/datetime/email helpers shared with the MCP server
    signatures.py               # Signature loading and HTML sanitization
    config.py                   # Config and local paths
    errors.py                   # Exceptions
    links.py                    # SharePoint/OneDrive URL extraction and Graph share encoding
    copilot.py                  # Copilot StreamHub client, file annotations, answer parsing
    copilot_bridge/             # Office-app bridge: protocol, native_host, client, app, install, __main__
    outlook_client.py           # Outlook REST API v2.0 client
    token_manager.py            # Token management + reauth trigger
    commands/
      account.py                # Account, auth, config and mailbox handlers/parsers
      drafts.py                 # Draft and mail-send handlers
      completion.py             # Completion command registration
      mcp.py                    # MCP command registration
      gateway_parser.py         # Gateway parser registration only
      calendar.py               # Calendar subcommands
      contacts.py               # Contacts subcommands
      copilot.py                # Copilot ask command and file resolution
      copilot_app.py            # copilot app status/inspect/ask and copilot extension install
      files.py                  # OneDrive and SharePoint file operations via Graph
      mail.py                   # Mail subcommands
      signature.py              # Signature fetch via OWA API interception
      tasks.py                  # Tasks subcommands
      teams.py                  # Teams read subcommands
    mcp/                        # MCP server: server.py, clients.py, one module per tool area
  extension/                    # Chrome MV3 bridge: sw.js, controller.js, diagnostics.js, hook.js, adapters/, tests/
```

## Caveats

- Uses the Outlook REST API v2.0, not Microsoft Graph.
- `copilot` uses an undocumented internal Microsoft protocol. If it stops working, record a new HAR from `m365.cloud.microsoft/chat` with a file attached and compare the `chat_request` frame with `outlook_draft/copilot.py`.
- Token validity is ~24 hours. If expired, the CLI auto-triggers headless reauth (prints MFA number to console).
- Sending (`draft send`, `mail send`, `teams send`) requires an explicit `--yes` flag. Without it the command previews and exits with code 2.
- Email bodies are HTML-to-text converted for console display, truncated at 3000 chars.
