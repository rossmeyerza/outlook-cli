# outlook-cli

CLI tool for managing Outlook email, drafts, and contacts via the Outlook REST API v2.0.

Authentication is built into this repo. The CLI uses Playwright to open Outlook Web, capture Microsoft bearer tokens, and cache them locally under `session_state/`.

## Installation

Clone the private repo and install it into a local virtualenv:

```bash
git clone git@github.com:rossmeyerza/outlook-draft-cli.git
cd outlook-draft-cli
python3 -m venv .venv
.venv/bin/pip install -e .
.venv/bin/python -m playwright install chromium
cp .env.example .env
ln -sf $(pwd)/.venv/bin/outlook-cli ~/.local/bin/outlook-cli
```

Check the command is available:

```bash
outlook-cli --help
```

## Usage

### Mail

```bash
outlook-cli mail unread                     # list unread emails
outlook-cli mail search colgate             # search emails by keyword
outlook-cli mail search "update" -n 10      # with result limit

outlook-cli mail read 1                     # read by index from last search
outlook-cli mail read <full-id>             # read by message ID
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

outlook-cli cal show 1                  # show event details by index from agenda
outlook-cli cal show <event-id>         # show by full ID
outlook-cli cal show 1 --json           # machine-readable JSON with recurrence metadata

outlook-cli cal create "Test event" "2026-04-10 14:00" "2026-04-10 15:00" -l "My Desk" --attendee "someone@example.com"
outlook-cli cal update 1 --start "2026-04-10 15:00" --end "2026-04-10 15:30" --location "Teams"
outlook-cli cal delete 1                # delete by index from agenda
outlook-cli cal accept 1                # accept an invitation
outlook-cli cal tentative 1             # tentatively accept an invitation
outlook-cli cal decline 1 -m "Sorry, I can't make it"  # decline with comment
outlook-cli cal accept 1 --no-send-response            # accept without emailing organizer
outlook-cli cal cancel 1 -m "Cancelled due to conflict" # cancel an event you organize
```

Results from `cal agenda` are cached to disk so `cal show <n>`, `cal delete <n>`, `cal accept <n>`, `cal tentative <n>`, `cal decline <n>`, and `cal cancel <n>` work across invocations.

### Tasks

```bash
outlook-cli task create "Buy milk"        # create a new task
outlook-cli task list                     # list incomplete tasks
outlook-cli task complete 1               # mark task as done by index
outlook-cli task delete 1                 # delete task by index
```

### Contacts

```bash
outlook-cli contact search ross
outlook-cli contact search "john smith" -n 20
```

Searches the org directory and recent contacts. Shows name, email, and contact type.

### Teams

```bash
outlook-cli teams list -n 20
outlook-cli teams show 1
outlook-cli teams messages 1 -n 20
```

Lists Teams chats, shows chat details, and reads chat messages. This is read-only.

### Auth and config

```bash
outlook-cli auth                         # run headless login
outlook-cli auth status                  # show token status
outlook-cli auth status --json           # machine-readable token status
outlook-cli auth clear                   # delete local token cache
outlook-cli config check                 # validate local setup without printing secrets
outlook-cli config check --json
```

`outlook-cli auth` runs the built-in headless auth flow: headless Chromium, enters credentials from `.env`, prints the Okta MFA verification number to the console, waits for push approval, and saves tokens to `session_state/tokens.json`.

Use a visible browser instead:

```bash
outlook-cli auth --headed
```

## Configuration

`.env` in the project root:

```
MS_EMAIL=your.email@company.com
MS_PASSWORD=your-password
LOCAL_TIMEZONE=Europe/London
OUTLOOK_TIMEZONE=GMT Standard Time
SIGNATURE_NEW_FILE=signature-new.html
SIGNATURE_REPLY_FILE=signature-reply.html
```

The token file is local to this repo at `session_state/tokens.json`. Both `.env` and `session_state/` are ignored by git.

`LOCAL_TIMEZONE` is the Python timezone used to interpret calendar times you type. `OUTLOOK_TIMEZONE` is the Microsoft timezone sent to Outlook. For the UK, use `Europe/London` and `GMT Standard Time`.

Draft signatures are loaded automatically from `SIGNATURE_NEW_FILE` and `SIGNATURE_REPLY_FILE`. Relative paths are resolved from the project root. If unset, the CLI defaults to:

- `signature-new.html` for new drafts
- `signature-reply.html` for reply drafts

Draft bodies are sent as HTML and wrapped with Aptos styling before the saved signature is appended.

## Architecture

```
outlook_draft/
  auth.py            # Built-in Playwright auth and token capture
  calendar_time.py   # Calendar timezone parsing and Outlook timezone headers
  cli.py             # CLI with argparse subcommands
  signatures.py      # Signature path loading and HTML sanitization
  config.py          # Paths and env vars
  errors.py          # Exception types
  outlook_client.py  # Outlook REST API v2.0 client (mail, drafts, contacts)
  token_manager.py   # Token loading, validation, triggers reauth
```

Auth flow: `token_manager.py` reads local tokens from `session_state/tokens.json`. Outlook features use the Outlook token, and Teams browsing uses the Microsoft Graph token. If expired, it runs the built-in auth flow from `outlook_draft/auth.py`.

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
