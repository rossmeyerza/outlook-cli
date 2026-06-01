# outlook-cli

CLI tool for managing Outlook email, drafts, and contacts via the Outlook REST API v2.0.

Authentication is built into this repo. The CLI uses Playwright to open Outlook Web, capture Microsoft bearer tokens, and cache them locally under `~/.local/share/outlook-cli/session_state/`.

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
outlook-cli mail folders --json             # machine-readable folders
outlook-cli mail move 1 --folder Archive    # move email to folder
outlook-cli mail attachments 1              # list attachments
outlook-cli mail download-attachments 1 -d ./attachments
```

Results from `mail unread` or `mail search` are cached to disk, so `mail read <n>` works in a separate invocation. Those cached message refs can also be used with `draft reply <n>`.

### Drafts

```bash
# Create a draft
outlook-cli draft create --to someone@example.com -s "Hello" -b "Hi there"
outlook-cli draft create --to someone@example.com -s "Report" -f ./email.html --html

# Create a reply draft from an existing email
outlook-cli draft reply 1 -b "Thanks, I'll review this and come back to you."
outlook-cli draft reply <message-id> --reply-all -f ./reply.html --html

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
outlook-cli cal availability --attendee someone@example.com --start "2026-04-10 09:00" --end "2026-04-10 17:00"
outlook-cli cal find-time --attendee someone@example.com --start "2026-04-10 09:00" --end "2026-04-10 17:00"
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
outlook-cli contact create --name "Ada Lovelace" --email ada@example.com
outlook-cli contact update 1 --company "Analytical Engines Ltd"
```

Searches the org directory and recent contacts. Personal contact create/update uses Outlook contacts.

### Teams

```bash
outlook-cli teams list -n 20
outlook-cli teams self
outlook-cli teams show 1
outlook-cli teams messages 1 -n 20
```

Lists Teams chats, shows chat details, and reads messages. `teams list` sorts by the latest received user message, ignoring system events and your own messages where identifiable. Teams sending is intentionally disabled for agent safety.

### Teams Gateway

```bash
outlook-cli gateway start
outlook-cli gateway start --self-chat
outlook-cli gateway start --chat-id "19:..." --trigger "@Marlow" --poll 30
outlook-cli gateway status
outlook-cli gateway stop
```

The gateway watches one configured Teams chat for the trigger, sends the prompt to `pi --mode rpc`, and posts Marlow's response back into the chat. Runtime state lives in `~/.local/share/outlook-cli/session_state/`, including `gateway_state.json`, `gateway.pid`, `gateway.log`, and per-chat Pi session files.

`--self-chat` uses the Microsoft Graph special self-chat thread, `48:notes`. This is different from the normal `/me/chats` entry that can appear as a one-person chat.

Gateway-native chat commands use `!` after the trigger so they do not collide with Teams slash commands:

```text
@Marlow !help
@Marlow !status
@Marlow !new
@Marlow !reset
@Marlow !pause
@Marlow !resume
@Marlow !tools
@Marlow !logs
```

`!new` starts a fresh Pi conversation for the chat. `!reset` clears that chat's persisted Pi session before starting fresh. `!pause` ignores normal prompts until `!resume`, while command messages still work.

### Files (OneDrive and SharePoint)

```bash
# Discover SharePoint sites
outlook-cli files sites

# Browse
outlook-cli files list                                     # OneDrive root
outlook-cli files list "Documents/Reports"                 # OneDrive subfolder
outlook-cli files list --site "Tesco"                      # SharePoint site root
outlook-cli files list --site "Tesco" "Shared Documents"   # SharePoint subfolder

# Upload
outlook-cli files upload ./report.pdf "Documents/Reports"
outlook-cli files upload ./deck.pptx "Shared Documents" --site "Tesco"

# Download
outlook-cli files download "Documents/Reports/report.pdf"
outlook-cli files download "Documents/Reports/report.pdf" ./downloads/
outlook-cli files download "Shared Documents/deck.pptx" ./deck.pptx --site "Tesco"

# Create folders
outlook-cli files mkdir "Documents/Reports/Q2"
outlook-cli files mkdir "Shared Documents/Proposals" --site "Tesco"

# Rename and move
outlook-cli files rename "Documents/old.pdf" "new.pdf"
outlook-cli files rename "Shared Documents/old.pdf" "new.pdf" --site "Tesco"
outlook-cli files move "Documents/file.pdf" "Archive"
outlook-cli files move "Shared Documents/file.pdf" "Archive" --site "Tesco"
```

`--site` matches case-insensitively and accepts partial names — `--site nestle` will find "MAP Nestle" without needing the full name. If the match is ambiguous it tells you which names matched. Without `--site`, all operations target your personal OneDrive.

Uploads under 4 MB use a single PUT. Larger files use chunked upload sessions automatically. Downloads save to the current directory by default and require `--overwrite` if the local output file already exists.

### Auth and config

```bash
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

`outlook-cli signature fetch` opens a headless browser using your saved OWA session, intercepts the signature API responses that OWA fires on load, and saves the active new-message and reply signatures to the files configured in `~/.config/outlook-cli/.env`. No MFA required after the first `auth`. On a fresh install, `auth` and `signature fetch` run automatically.

`outlook-cli auth` runs the built-in headless auth flow: headless Chromium, enters credentials from `~/.config/outlook-cli/.env`, prints the MFA challenge number to the console (works with Okta Verify, Microsoft Authenticator, and similar push-based MFA), waits for approval, and saves tokens to `~/.local/share/outlook-cli/session_state/tokens.json`.

Use a visible browser instead:

```bash
outlook-cli auth --headed
```

## Configuration

Config file:

`~/.config/outlook-cli/.env`

```
MS_EMAIL=your.email@company.com
MS_PASSWORD=your-password
LOCAL_TIMEZONE=Europe/London
OUTLOOK_TIMEZONE=GMT Standard Time
SIGNATURE_NEW_FILE=signature-new.html
SIGNATURE_REPLY_FILE=signature-reply.html
```

The token file is at `~/.local/share/outlook-cli/session_state/tokens.json`. The browser session state (used by `signature fetch` to avoid re-auth) is at `~/.local/share/outlook-cli/session_state/browser_state.json`.

`LOCAL_TIMEZONE` is the Python timezone used to interpret calendar times you type. `OUTLOOK_TIMEZONE` is the Microsoft timezone sent to Outlook. For the UK, use `Europe/London` and `GMT Standard Time`.

Draft signatures are loaded automatically from `SIGNATURE_NEW_FILE` and `SIGNATURE_REPLY_FILE`. Relative paths are resolved from `~/.local/share/outlook-cli`. If unset, the CLI defaults to:

- `signature-new.html` for new drafts
- `signature-reply.html` for reply drafts

Draft bodies are sent as HTML and wrapped with Aptos styling before the saved signature is appended.

## Architecture

```
outlook_draft/
  auth.py              # Built-in Playwright auth, token capture, browser session save
  calendar_time.py     # Calendar timezone parsing and Outlook timezone headers
  cli.py               # CLI with argparse subcommands
  signatures.py        # Signature path loading and HTML sanitization
  config.py            # Paths and env vars
  errors.py            # Exception types
  links.py             # SharePoint/OneDrive URL extraction and Graph share ID encoding
  outlook_client.py    # Outlook REST API v2.0 client (mail, drafts, contacts)
  token_manager.py     # Token loading, validation, triggers reauth
  commands/
    calendar.py        # Calendar subcommands
    contacts.py        # Contacts subcommands
    mail.py            # Mail subcommands
    signature.py       # Signature fetch via OWA API interception
    tasks.py           # Tasks subcommands
    teams.py           # Teams read subcommands
```

Auth flow: `token_manager.py` reads local tokens from `~/.local/share/outlook-cli/session_state/tokens.json`. Outlook features use the Outlook token, and Teams browsing uses the Microsoft Graph token. If expired, it runs the built-in auth flow from `outlook_draft/auth.py`.

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

## License

MIT License. Copyright (c) 2026 Ross Meyer.
