# outlook-cli

CLI tool for managing Outlook email, drafts, and contacts via the Outlook REST API v2.0.

Authentication is shared with the other Microsoft projects (onedrive-fuse, sharepoint-fuse) through [ms-graph-explorer](../ms-graph-explorer/), which handles headless SSO login with Okta MFA.

## Installation

Clone the private repo and install it into a local virtualenv:

```bash
git clone git@github.com:rossmeyerza/outlook-draft-cli.git
cd outlook-draft-cli
python3 -m venv .venv
.venv/bin/pip install -e .
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

outlook-cli cal show 1                  # show event details by index from agenda
outlook-cli cal show <event-id>         # show by full ID

outlook-cli cal create "Test event" "2026-04-10 14:00" "2026-04-10 15:00" -l "My Desk" --attendee "someone@example.com"
```

Results from `cal agenda` are cached to disk so `cal show <n>` works across invocations.

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

### Force re-authentication

```bash
outlook-cli auth
```

Runs ms-graph-explorer's headless auth flow: headless Chromium, enters credentials, prints the Okta MFA verification number to the console, waits for push approval.

## Configuration

`.env` in the project root:

```
MS_EMAIL=your.email@company.com
```

Credentials (email + password) live in ms-graph-explorer's `.env`. The token file is shared at `../ms-graph-explorer/session_state/tokens.json`.

Draft signatures are loaded automatically from:

- `signature-new.html` for new drafts
- `signature-reply.html` for reply drafts

Draft bodies are sent as HTML and wrapped with Aptos styling before the saved signature is appended.

## Architecture

```
outlook_draft/
  cli.py             # CLI with argparse subcommands
  config.py          # Paths, env vars, points to ms-graph-explorer
  errors.py          # Exception types
  outlook_client.py  # Outlook REST API v2.0 client (mail, drafts, contacts)
  token_manager.py   # Token loading, validation, triggers reauth
```

Auth flow: `token_manager.py` reads tokens from ms-graph-explorer's `tokens.json`. Outlook features use the Outlook token, and Teams browsing uses the Microsoft Graph token. If expired, it runs `auth.py --headless` from ms-graph-explorer's venv.

## Dependencies

- httpx (HTTP client)
- pyjwt (token decoding)
- python-dotenv (env loading)
- rich (terminal output)

## License

MIT License. Copyright (c) 2026 Ross Meyer.
