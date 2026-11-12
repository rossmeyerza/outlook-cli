# outlook-cli

CLI tool for managing Outlook email, drafts, and contacts via the Outlook REST API v2.0.

Authentication is built into this repo. The CLI uses Playwright to open Outlook Web, capture Microsoft bearer tokens, and cache them locally. Single-account installs use `~/.local/share/outlook-cli/session_state/`; account profiles use isolated stores under `~/.local/share/outlook-cli/accounts/<name>/`.

## Installation

### One-line installer (recommended)

```bash
curl -fsSL https://outlook-cli.21436587.xyz/install.sh | bash
```

Requires Python 3.12+ and `git`. The script will:

- Clone this repo to `~/.local/lib/outlook-cli`
- Set up a Python virtualenv and install all dependencies
- Install Playwright Chromium for authentication
- Prompt you for your Microsoft 365 email, password, and timezone
- Write your config to `~/.config/outlook-cli/.env` automatically
- Symlink `outlook-cli` to `~/.local/bin`

You can read the full script before running it at:
`https://outlook-cli.21436587.xyz/install.sh`

### Manual installation

```bash
git clone https://github.com/rossmeyerza/outlook-cli.git
cd outlook-cli
python3 -m venv .venv
.venv/bin/pip install -e .
.venv/bin/python -m playwright install chromium
mkdir -p ~/.config/outlook-cli
cp .env.example ~/.config/outlook-cli/.env
ln -sf $(pwd)/.venv/bin/outlook-cli ~/.local/bin/outlook-cli
```

Then edit `~/.config/outlook-cli/.env` and set `MS_EMAIL` and `MS_PASSWORD`.

Check the command is available:

```bash
outlook-cli --help
```

### Upgrading

Re-run the installer at any time to pull the latest changes:

```bash
curl -fsSL https://outlook-cli.21436587.xyz/install.sh | bash
```

The script detects an existing install and runs `git pull` and updates dependencies without touching `~/.config/outlook-cli/.env`.

## Usage

Interactive commands show Rich spinners while waiting on Microsoft APIs. Use `--no-spinner` or `OUTLOOK_CLI_NO_SPINNER=1` to disable them; spinner output goes to stderr so JSON stdout stays parseable.

### Mail

```bash
outlook-cli mail unread                     # list unread emails
outlook-cli mail unread --json              # machine-readable unread emails
outlook-cli mail search colgate             # search emails by keyword
outlook-cli mail search "update" -n 10      # with result limit
outlook-cli mail search colgate --json      # machine-readable search results

outlook-cli mail read 1                     # read by index from last search
outlook-cli mail read <full-id>             # read by message ID
outlook-cli mail read 1 --json              # full message JSON
outlook-cli mail mark-read 1                # mark email as read
outlook-cli mail mark-unread 1              # mark email as unread
outlook-cli mail archive 1                  # archive email
outlook-cli mail folders                    # list mail folders
outlook-cli mail create-folder "Clients"    # create a top-level mail folder
outlook-cli mail create-folder "Project" --parent inbox # create a subfolder
outlook-cli mail folders --json             # machine-readable folders
outlook-cli mail move 1 --folder Archive    # move email to folder
outlook-cli mail attachments 1              # list attachments
outlook-cli mail download-attachments 1 -d ./attachments
```

Results from `mail unread` or `mail search` are cached to disk, so `mail read <n>` works in a separate invocation. Those cached message refs can also be used with `draft reply <n>`.

See [Mail features](docs/mail-features.md) for bulk actions, complete conversations, search filters, attachments, draft edits, forward drafts, categories, rules and safety controls.
Categories and rules need additional Microsoft Graph mailbox permissions.
The current local token does not include those permissions.

### Drafts

```bash
# Create a draft
outlook-cli draft create --to someone@example.com -s "Hello" -b "Hi there"
outlook-cli draft create --to someone@example.com -s "Report" -f ./email.html --format html

# Create a reply draft from an existing email
outlook-cli draft reply 1 -b "Thanks, I'll review this and come back to you."
outlook-cli draft reply <message-id> --reply-all -f ./reply.html --format html

# List drafts
outlook-cli draft list
outlook-cli draft list -n 50

# View or delete
outlook-cli draft show 1          # by index from list
outlook-cli draft delete 1        # by index
```

### Calendar

```bash
outlook-cli cal agenda                  # list upcoming events (next 7 days)
outlook-cli cal agenda --days 14        # list next 14 days
outlook-cli cal agenda --plain          # plain text output
outlook-cli cal agenda --json           # machine-readable JSON
outlook-cli cal agenda --table          # explicit table output

outlook-cli cal show 1                  # show event details by index from agenda
outlook-cli cal show <event-id>         # show by full ID
outlook-cli cal show 1 --json           # machine-readable JSON with recurrence metadata

outlook-cli cal create "Test event" "2026-04-10 14:00" "2026-04-10 15:00" -l "My Desk" --attendee "someone@example.com"
outlook-cli cal rooms --json               # find rooms, if supported by tenant/token
outlook-cli cal availability --attendee someone@example.com --start "2026-04-10 09:00" --end "2026-04-10 17:00"  # busy ranges + common free time
outlook-cli cal availability --attendee someone@example.com --start "2026-04-10 09:00" --end "2026-04-10 17:00" --raw  # raw Outlook JSON
outlook-cli cal find-time --attendee a@example.com --attendee b@example.com --start "2026-04-10 09:00" --end "2026-04-17 17:00" --duration 30  # slots, best first
outlook-cli cal find-time --attendee someone@example.com --start "2026-04-10 09:00" --end "2026-04-10 17:00" --outlook  # Outlook's own suggestions
outlook-cli cal update 1 --start "2026-04-10 15:00" --end "2026-04-10 15:30" --location "Teams"
outlook-cli cal delete 1                # delete by index from agenda
outlook-cli cal accept 1                # accept an invitation
outlook-cli cal tentative 1             # tentatively accept an invitation
outlook-cli cal decline 1 -m "Sorry, I can't make it"  # decline with comment
outlook-cli cal accept 1 --no-send-response            # accept without emailing organizer
outlook-cli cal cancel 1 -m "Cancelled due to conflict" # cancel an event you organize
```

Results from `cal agenda` are cached to disk so `cal show <n>`, `cal update <n>`, `cal delete <n>`, `cal accept <n>`, `cal tentative <n>`, `cal decline <n>`, and `cal cancel <n>` work across invocations. Room discovery is tenant/token dependent; use `cal availability` or `cal find-time` with known room email addresses if `cal rooms` is unavailable.

### Tasks

```bash
outlook-cli task create "Buy milk"        # create a new task
outlook-cli task list                     # list incomplete tasks
outlook-cli task list --json              # machine-readable tasks
outlook-cli task update 1 --due "2026-05-20" --importance High
outlook-cli task complete 1               # mark task as done by index
outlook-cli task delete 1                 # delete task by index
```

### Contacts

```bash
outlook-cli contact search ross
outlook-cli contact search "john smith" -n 20
outlook-cli contact search ross --json
outlook-cli contact org me
outlook-cli contact org person@example.com --chain --json
outlook-cli contact create --name "Ada Lovelace" --email ada@example.com
outlook-cli contact update 1 --company "Analytical Engines Ltd"
```

Searches the org directory and recent contacts. Personal contact create/update uses Outlook contacts.

`contact org` reads profiles, managers and direct reports from the Microsoft Graph directory.
Use `--chain` for the recorded management chain.
Use an email, UPN or directory user ID when a name is ambiguous.
Missing fields and incomplete results remain explicit.
See [organisation details](docs/contact-org.md) for permissions, result fields and limitations.

### Teams

```bash
outlook-cli teams list -n 20
outlook-cli teams search tarik --scan 200 -n 10
outlook-cli teams self
outlook-cli teams show 1
outlook-cli teams messages 1 -n 20
```

Lists and searches Teams chats, shows chat details, and reads messages. `teams messages` fetches the newest messages first and follows Graph paging, so long chats show recent messages and `--count` above 50 works; CLI and MCP results show the newest message first. MCP message and attachment scans now accept the caller's count and use the client's pagination. `teams search` matches chat topics and participants, then caches the matching chats so `teams show <n>` and `teams messages <n>` work from the search result list. `teams list` sorts by the latest received user message, ignoring system events and your own messages where identifiable. Sending is enabled but guarded: `draft send <n>`, `mail send --to ... -s ... -b ...` `teams send <chat> <msg>` and `teams new --to <person> <msg>` print a preview and send nothing unless `--yes` is passed. `teams new` starts a 1:1 chat (or reuses the existing one) or, with two or more `--to` recipients, a new group chat, then sends the first message. Prefer drafts for email and only pass `--yes` when Ross has explicitly asked to send.

`teams send` and `teams new` also accept repeatable `--attach` and `--image` options.
File attachments use an existing OneDrive folder and named chat-member access.
Inline images use Teams hosted content.
See [Teams media](docs/teams-media.md) for previews, limits, permissions and failure handling.

### Teams Gateway

```bash
outlook-cli gateway start
outlook-cli gateway start --self-chat
outlook-cli gateway start --chat-id "19:..." --trigger "@Marlow" --poll 30
outlook-cli gateway start --self-chat --model claude-opus-4.8 --provider wpp
outlook-cli gateway start --self-chat --model sonnet:high
outlook-cli gateway status
outlook-cli gateway stop
```

The gateway watches one configured Teams chat for the trigger, sends the prompt to `pi --mode rpc`, and posts Marlow's response back into the chat. Runtime state lives in `~/.local/share/outlook-cli/session_state/`, including `gateway_state.json`, `gateway.pid`, `gateway.log`, and per-chat Pi session files. Pi runs from a separate per-chat workspace under `~/.local/share/outlook-cli/gateway_workspaces/`, so it does not treat the Outlook CLI install directory as the active project.

`outlook-cli teams self` checks access to your Teams self-chat. `--self-chat` uses the Microsoft Graph special self-chat thread, `48:notes`. This is different from the normal `/me/chats` entry that can appear as a one-person chat.

When a prompt is received, the gateway posts a short `...` receipt message immediately. It then tries to soft-delete that receipt before posting the final response. If Microsoft Graph denies the delete, the final response still posts and the failed cleanup is written to the gateway log.

If the Microsoft Graph token expires, the gateway attempts one headless re-authentication. If that fails, it records the auth error in `gateway_state.json` and exits instead of retrying every poll interval.

Gateway-native chat commands use `!` after the trigger so they do not collide with Teams slash commands:

```text
@Marlow !help
@Marlow !status
@Marlow !new
@Marlow !reset
@Marlow !help model
@Marlow !model
@Marlow !model help
@Marlow !model list
@Marlow !model list sonnet
@Marlow !model claude-opus-4.8
@Marlow !model --provider wpp --model claude-opus-4.8 --thinking high
@Marlow !model reset
@Marlow !pause
@Marlow !resume
@Marlow !tools
@Marlow !files
@Marlow !send report.html
@Marlow !logs
```

`!new` starts a fresh Pi conversation for the chat. `!reset` clears that chat's persisted Pi session before starting fresh. `!pause` ignores normal prompts until `!resume`, while command messages still work.

`!help <command>` shows help for a specific gateway command. `!model` shows the current Pi settings for the chat. `!model list [search]` lists available Pi models through `pi --list-models`. `!model <model>` or `!model --provider <provider> --model <model> --thinking <level>` updates the saved settings and restarts the Pi process on the next prompt while keeping the chat's persisted session files. `!model reset` returns to Pi defaults. Short model names such as `sonnet` use Pi's fuzzy/pattern matching and may change as available models change; use `!model list sonnet` and set an exact model ID when repeatability matters.

Marlow can send generated workspace files back to Teams without exposing a general Teams-send tool to Pi. Files must live under the per-chat workspace and use a supported extension such as `.html`, `.md`, `.txt`, `.csv`, `.json`, `.pdf`, `.png`, `.jpg`, `.docx`, `.pptx`, or `.xlsx`. The gateway uploads the file to `Outlook CLI/Gateway/<chat-hash>/` in Ross's OneDrive, creates an organization view link, and posts that link to Teams. Use `!files` to list sendable workspace files and `!send <path>` to publish one manually.

For automatic publishing, Marlow writes `.marlow-export.json` in the chat workspace after creating files:

```json
{
  "files": ["report.html"],
  "message": "Created the report."
}
```

After the Pi response is posted, the gateway validates the paths, uploads the files, posts the links, and removes the manifest.

### Files (OneDrive and SharePoint)

```bash
# Discover SharePoint sites
outlook-cli files sites
outlook-cli files libraries --site "Tesco"

# Browse
outlook-cli files list                                     # OneDrive root
outlook-cli files list "Documents/Reports"                 # OneDrive subfolder
outlook-cli files list --site "Tesco"                      # SharePoint libraries
outlook-cli files list "General" --site "Tesco" --library "Documents"
outlook-cli files list "General" --site "Tesco" --library "Documents" --links
outlook-cli files list "General" --site "Tesco" --library "Documents" --json

# Search SharePoint
outlook-cli files search budget --site "Tesco"
outlook-cli files search budget --site "Tesco" --library "Documents" --links
outlook-cli files search budget --site "Tesco" --json

# Upload
outlook-cli files upload ./report.pdf "Documents/Reports"
outlook-cli files upload ./deck.pptx "General" --site "Tesco" --library "Documents"

# Download
outlook-cli files download "Documents/Reports/report.pdf"
outlook-cli files download "Documents/Reports/report.pdf" ./downloads/
outlook-cli files download "General/deck.pptx" ./deck.pptx --site "Tesco" --library "Documents"
outlook-cli files download <drive-item-id> ./downloads/ --site "Tesco" --library "Documents"

# Create folders
outlook-cli files mkdir "Documents/Reports/Q2"
outlook-cli files mkdir "Proposals" --site "Tesco" --library "Documents"

# Rename and move
outlook-cli files rename "Documents/old.pdf" "new.pdf"
outlook-cli files rename "General/old.pdf" "new.pdf" --site "Tesco" --library "Documents"
outlook-cli files move "Documents/file.pdf" "Archive"
outlook-cli files move "General/file.pdf" "Archive" --site "Tesco" --library "Documents"
```

`--site` and `--library` match case-insensitively and accept partial names. `files list --site <name>` shows document libraries; pass `--library <name>` to browse or mutate a library. `--links` adds browser links to table output. `--json` returns machine-readable items with `id`, `path`, `parentPath`, `library`, and `webUrl`. `files download` downloads file bytes to disk and accepts either a relative path or a drive item ID from search/list JSON; browser links are for opening the file in SharePoint/OneDrive.

Uploads under 4 MB use a single PUT. Larger files use chunked upload sessions automatically. Downloads save to the current directory by default and require `--overwrite` if the local output file already exists.

### Microsoft 365 Copilot

```bash
# Ask a question
outlook-cli copilot ask "What changed in our travel policy this year?"

# Ask about a file: a SharePoint/OneDrive URL, a drive item ID, or a path
outlook-cli copilot ask "Summarise this deck" --file "https://contoso.sharepoint.com/sites/team/Shared Documents/Plan.pptx"
outlook-cli copilot ask "List the risks" --file "General/Plan.docx" --site "Tesco" --library "Documents"
outlook-cli copilot ask "Compare these" --file "Documents/a.docx" --file "Documents/b.docx"

# Ask a follow-up in the same conversation
outlook-cli copilot ask "What does it say about budget?" --continue
outlook-cli copilot ask "And the timeline?" --conversation <conversation-id>

# Machine-readable answer
outlook-cli copilot ask "Summarise this deck" --file <url> --json
```

`copilot ask` sends a prompt to Microsoft 365 Copilot (the work chat at `m365.cloud.microsoft/chat`). `--file` attaches a OneDrive or SharePoint file as context. It is repeatable. A URL can be the `webUrl` from `files search --json`. Paths and drive item IDs are in your OneDrive, unless you give `--site`. The answer streams to the terminal, then the sources and the conversation ID show. `--continue` uses the last conversation of the active account. Copilot keeps the files of earlier turns, so a follow-up does not need `--file` again.

Caution: Copilot is not read-only. It can create drafts, post Teams messages, change calendars and delete or share SharePoint files. A prompt can cause these actions. After the answer, the CLI lists each tool that Copilot called (`toolCalls` in `--json`). If Copilot asks for a confirmation, the CLI stops and does not confirm it (`pendingConfirmations`). Copilot's document tools write edits to a new file, not to the original SharePoint file. For the full list of tools and the test results, see [docs/copilot-tools.md](docs/copilot-tools.md).

Warning: Microsoft has no public API for this access. The command uses the internal WebSocket protocol of the Copilot web app. Microsoft can change this protocol without notice, and then the command stops working.

The command needs a Microsoft 365 Copilot licence. `outlook-cli auth` captures the Copilot token after the normal login. The token is valid for about an hour. When it expires, the CLI gets a new token from the saved browser session, with no password or MFA step.

### Copilot in Office web apps (experimental)

This bridge sends prompts to the Copilot pane in Excel or PowerPoint in your Chrome session. It can edit the live file, not a copy. Co-authors see changes at once. Word support is experimental in extension 0.3.0. Its OfficeAgent chat schemas were observed on a dummy document, but automated Word acceptance checks remain pending.

```bash
outlook-cli copilot extension install
# Open chrome://extensions, enable Developer mode, then Load unpacked.
# Select the folder printed by the install command.
outlook-cli copilot app status --json
outlook-cli copilot app ask "Add a sample table in A1:C4" --file "Documents/Test.xlsx" --json
outlook-cli copilot app ask "Add an agenda slide" --file "Documents/Test.pptx" --open --wait-saved
```

Use disposable test files first. Excel read-only and single-cell edit checks passed. PowerPoint read-only, title-edit and slide-insertion checks passed after manual pane opening. Saved files were independently checked. Automatic PowerPoint pane opening also passed after a tab refresh, including a background-tab read; cold-start reliability still needs checks. Approval detection is not verified. It never clicks approval controls or changes **Allow editing**. An approval, timeout or lost connection does not undo edits. Inspect the document before retrying. Do not use the pane manually while a request runs.

`--new-chat` starts a new pane conversation. Otherwise, the pane keeps its current conversation. `--open` permits a new tab. Without it, the file must already be open. `--wait-saved` checks Graph modification time. A changed timestamp is evidence of a file update, not proof that these specific edits were saved.

The `copilot_app` MCP tool returns the same result. It is marked as destructive and non-idempotent. The bridge uses a same-user Unix socket and Chrome native messaging. It does not capture Office tokens. See [the implementation plan and release checklist](docs/copilot-app-plan.md).

### Auth and config

```bash
outlook-cli account add wpp --email ross@wpp.com --switch
outlook-cli account add ikea --email ross@ikea.example
outlook-cli account list
outlook-cli account current
outlook-cli account switch ikea
outlook-cli account switch default
outlook-cli --account wpp mail unread

outlook-cli auth                         # run headless login
outlook-cli auth status                  # show token status
outlook-cli auth status --json           # machine-readable token status
outlook-cli auth clear                   # delete local token cache
outlook-cli auth scopes                  # list safe token metadata and scopes
outlook-cli mailbox show                 # show mailbox settings
outlook-cli mailbox update --timezone "GMT Standard Time"
outlook-cli config check                 # validate local setup without printing secrets
outlook-cli config check --json

outlook-cli signature fetch              # fetch new and reply signatures from OWA
outlook-cli signature fetch --headed     # run with a visible browser if headless fails
```

Account profiles are separate `.env` files in `~/.config/outlook-cli/accounts/<name>.env`. Each profile gets its own tokens, browser session, command caches, signatures, and gateway state under `~/.local/share/outlook-cli/accounts/<name>/`. `account switch <name>` sets the default profile; `account switch default` returns to the legacy `~/.config/outlook-cli/.env`; `--account <name>` overrides it for one command. If no profile is active, `account list` shows the legacy config as `(default)`.

`outlook-cli signature fetch` opens a headless browser using your saved OWA session, intercepts the signature API responses that OWA fires on load, and saves the active new-message and reply signatures to the active account's signature files. No MFA required after the first `auth`. On a fresh install, `auth` and `signature fetch` run automatically.

`outlook-cli auth` runs the built-in headless auth flow: headless Chromium, enters credentials from the active account env file, prints the MFA challenge number to the console (works with Okta Verify, Microsoft Authenticator, and similar push-based MFA), waits for approval, and saves tokens to the active account's `session_state/tokens.json`.

Login opens `outlook.cloud.microsoft` directly. It accepts both `outlook.office.com` and `outlook.cloud.microsoft` after sign-in. Outlook tokens from the new host use the existing `outlook.office.com` cache key. Each captured token is saved during login. Existing tokens for other services are kept. The cache file is replaced only after the new contents are written. If a later login step fails, tokens already saved remain available.

Use a visible browser instead:

```bash
outlook-cli auth --headed
```

### Shell completions

`outlook-cli completion <shell>` prints a completion script for `bash`, `zsh`,
`fish`, `powershell`, or `nushell`.

bash / zsh / fish / powershell are powered by [argcomplete](https://github.com/kislyuk/argcomplete),
which completes against the live command tree (subcommands, flags, and
choices like `--importance Low/Normal/High`). Register it once in your shell:

```bash
# bash (~/.bashrc) or zsh (~/.zshrc)
eval "$(outlook-cli completion bash)"   # use zsh in zsh

# fish (~/.config/fish/config.fish)
outlook-cli completion fish | source
```

```powershell
# PowerShell ($PROFILE)
outlook-cli completion powershell | Out-String | Invoke-Expression
```

Nushell has no argparse bridge, so the script is a generated module of
`extern` definitions. Save it and source it from your config:

```nushell
outlook-cli completion nushell | save -f ~/.cache/outlook-cli.nu
# then in config.nu:
source ~/.cache/outlook-cli.nu
```

Re-run the command after upgrading to pick up new subcommands or flags.

### MCP server

`outlook-cli mcp` runs a local stdio MCP server (server name `outlook`) over the same client code as the CLI. Use `outlook-cli --account <name> mcp` to serve a specific account profile.

Register it with Pi in `~/.pi/agent/mcp.json`:

```json
"outlook": {
  "command": "/home/ross/.local/lib/outlook-draft-cli/.venv/bin/outlook-cli",
  "args": ["mcp"],
  "description": "Outlook mail, drafts, calendar, contacts, tasks, Teams and files, with confirmed sending",
  "timeout": 120
}
```

Check it with `pi mcp list`.

Tools (all IDs are full Outlook or Graph IDs; numeric index caches are not used):

- Mail: `mail_search`, `mail_unread`, `mail_read`, `mail_attachments`, `mail_download_attachments`, `mail_links`, `mail_download_links`, `mail_folders`, `mail_create_folder`, `mail_mark_read`, `mail_move`, `mail_archive`
- Drafts: `draft_create`, `draft_reply`, `draft_update`, `draft_forward`, `draft_list`, `draft_get`, `draft_delete` (draft creation does not send; `draft_delete` refuses non-drafts)
- Calendar: `cal_agenda`, `cal_get`, `cal_availability`, `cal_find_time`, `cal_create`, `cal_respond`, `cal_delete`, `cal_cancel`, `cal_rooms`, `cal_update`
- Contacts: `contact_search`, `contact_create`, `contact_update`, `contact_org`
- Tasks: `task_list`, `task_create`, `task_complete`, `task_update`, `task_delete`
- Teams (read-only): `teams_list`, `teams_search`, `teams_get`, `teams_self`, `teams_messages`, `teams_attachments`, `teams_download_attachments`
- Files: `files_sites`, `files_libraries`, `files_list`, `files_search`, `files_download`, `files_upload`, `files_mkdir`, `files_rename`, `files_move`
- Copilot: `copilot_ask` (prompt with optional `files`; pass `conversation_id` for a follow-up). Not read-only: Copilot can take actions with its own tools. The result lists `toolCalls` and `pendingConfirmations`. The tool never confirms an action.
- Auth: `auth_status` (read-only)
- Sending: `mail_send`, `draft_send`, `teams_send` (preview by default; requires `confirm=true` to send)

Side effects to know about: `cal_create` with attendees emails invitations immediately, and `cal_respond` with `send_response` true emails the organizer.

Rules:

- Send tools `mail_send`, `draft_send` and `teams_send` preview by default. Set `confirm=true` only when the user explicitly requests that send. Do not automatically retry an uncertain send result. CLI sending still requires `--yes`.
- All 74 CLI operations across mail, drafts, calendar, contacts, tasks, Teams and files have MCP equivalents.
- New contact, task, calendar and file changes preview by default and require explicit confirmation.
- File uploads require explicit overwrite approval to replace an existing file. New remote writes have no automatic retries.
- Mailbox settings, account administration, interactive authentication, signatures and the gateway remain outside this parity scope.
- Calendar removal tools preview by default. Set `confirm=true` only after explicit approval of the exact event and scope.
- Whole-series removal also requires `allow_series=true`. Organiser meetings with attendees require `cal_cancel`, which notifies attendees.
- Use `cal_respond` to decline an invitation. Calendar removal does not delete the invitation email.
- Calendar delete and cancel writes have no automatic retries.

See [the complete CLI-to-MCP table](docs/cli-mcp-parity.md) for all commands, partial support and missing tools.
- The server never runs the login. If a token is expired, tools return "Run 'outlook-cli auth' in a terminal, then retry." Tokens are re-read on every call, so a fresh `outlook-cli auth` works without restarting the server.
- One exception: `copilot_ask` gets a new Copilot token from the saved browser session when the token expires. This step has no password or MFA. If the saved session has expired, the tool returns the auth message.
- stdout is the protocol channel. Logs go to stderr.

## Configuration

Config file:

`~/.config/outlook-cli/.env`

Account profile files:

`~/.config/outlook-cli/accounts/<name>.env`

```
MS_EMAIL=your.email@company.com
MS_PASSWORD=your-password
LOCAL_TIMEZONE=Europe/London
OUTLOOK_TIMEZONE=GMT Standard Time
SIGNATURE_NEW_FILE=signature-new.html
SIGNATURE_REPLY_FILE=signature-reply.html
```

For the legacy/global config, the token file is at `~/.local/share/outlook-cli/session_state/tokens.json`. For account profiles, it is at `~/.local/share/outlook-cli/accounts/<name>/session_state/tokens.json`. Browser session state follows the same pattern as `browser_state.json`.

`LOCAL_TIMEZONE` is the Python timezone used to interpret calendar times you type. `OUTLOOK_TIMEZONE` is the Microsoft timezone sent to Outlook. For the UK, use `Europe/London` and `GMT Standard Time`.

Draft signatures are loaded automatically from `SIGNATURE_NEW_FILE` and `SIGNATURE_REPLY_FILE`. Relative paths are resolved from the active account data directory. If unset, the CLI defaults to:

- `signature-new.html` for new drafts
- `signature-reply.html` for reply drafts

Email composition accepts Markdown (default) or HTML through `--format markdown|html`.
The MCP email tools expose `body_format` with the same two choices.
Markdown renders as HTML with native lists, compact paragraph margins and Aptos before the saved signature.
Supplied HTML, saved signatures and quoted originals retain their existing markup.

## Architecture

```
outlook_draft/
  auth.py              # Built-in Playwright auth, token capture, browser session save
  calendar_time.py     # Calendar timezone parsing and Outlook timezone headers
  cli.py               # CLI with argparse subcommands
  formatting.py        # Pure HTML/text/datetime/email-body helpers shared by CLI and MCP
  signatures.py        # Signature path loading and HTML sanitization
  config.py            # Paths and env vars
  errors.py            # Exception types
  links.py             # SharePoint/OneDrive URL extraction and Graph share ID encoding
  copilot.py           # Microsoft 365 Copilot StreamHub client, file annotations, answer parsing
  outlook_client.py    # Outlook REST API v2.0 client (mail, drafts, contacts)
  token_manager.py     # Token loading, validation, triggers reauth
  commands/
    calendar.py        # Calendar subcommands
    contacts.py        # Contacts subcommands
    copilot.py         # Copilot ask command and file resolution
    mail.py            # Mail subcommands
    signature.py       # Signature fetch via OWA API interception
    tasks.py           # Tasks subcommands
    teams.py           # Teams read subcommands
  mcp/
    server.py          # FastMCP app, tool registration, stdio run()
    clients.py         # Per-call clients with no-reauth token handling
    mail.py drafts.py calendar.py contacts.py tasks.py teams.py files.py copilot.py auth.py  # Tool modules
```

Auth flow: `token_manager.py` reads local tokens from `~/.local/share/outlook-cli/session_state/tokens.json`. Outlook features use the Outlook token, Teams browsing and files use the Microsoft Graph token, and `copilot` uses the Copilot token (audience `https://substrate.office.com/sydney`, saved under the key `substrate.office.com/sydney`). If a token is expired, the CLI runs the built-in auth flow from `outlook_draft/auth.py`. For the Copilot token, it first tries the saved browser session.

## Tests

```bash
.venv/bin/pip install -e '.[dev]'
.venv/bin/python -m pytest -q
```

Live Outlook calendar tests are gated because they mutate real calendar data:

```bash
OUTLOOK_CLI_LIVE=1 .venv/bin/python -m pytest -q tests/test_live_calendar.py
```

## Dependencies

- httpx (HTTP client)
- pyjwt (token decoding)
- python-dotenv (env loading)
- python-dateutil (date parsing)
- playwright (browser auth automation)
- requests (token validation)
- rich (terminal output)
- mcp (official MCP Python SDK, 1.x, for `outlook-cli mcp`)
- websockets (Copilot StreamHub connection)

## License

MIT License. Copyright (c) 2026 Ross Meyer.
