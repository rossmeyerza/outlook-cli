# outlook-cli

## Overview

CLI for managing Outlook email, drafts, and contacts via the Outlook REST API v2.0. Globally available as `outlook-cli`.

## Quick Reference

```bash
# Search and read emails
outlook-cli mail unread                    # list unread emails
outlook-cli mail search <query>            # search emails by keyword
outlook-cli mail search colgate -n 10      # with result limit
outlook-cli mail read <n>                  # read email by index from last unread/search
outlook-cli mail read <full-id>            # read by Outlook message ID

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
outlook-cli cal show <n>                   # show event details by index
outlook-cli cal create "Subj" "Start" "End" # create event (e.g. "2026-04-10 14:00")

# Manage tasks
outlook-cli task create "<name>"           # create a new task
outlook-cli task list                      # list active tasks
outlook-cli task complete <n>              # mark task as done
outlook-cli task delete <n>                # delete task

# Search contacts
outlook-cli contact search <query>         # search by name or email
outlook-cli contact search ross -n 20      # with result limit

# Browse Teams chats (read-only)
outlook-cli teams list -n 20               # list chats
outlook-cli teams show <chat-ref>          # show chat details
outlook-cli teams messages <chat-ref> -n 20 # read messages

# Auth
outlook-cli auth                           # force headless re-authentication
```

## Key Details

- **Email search**: `mail unread` and `mail search` results are cached to disk so `mail read <n>` works in a separate invocation.
- **Email reading**: `mail read` strips HTML to plain text for console display. Truncates at 3000 chars.
- **Calendar**: `cal agenda` results are cached to disk so `cal show <n>` works in a separate invocation. Shows attendee responses.
- **Tasks**: `task list` caches to disk so `task complete <n>` works in a separate invocation.
- **Recipients**: `--to`, `--cc`, `--bcc` are repeatable and accept comma-separated addresses.
- **Body source**: `--body` for inline text, `--body-file` for file. Add `--html` when the provided body is already HTML.
- **Reply drafts**: `draft reply` creates a draft tied to an existing message. Use a cached mail index from `mail unread` / `mail search`, a cached ID suffix, or a full message ID.
- **Draft formatting**: New and reply drafts are always saved as HTML, use Aptos for the message body, and append the saved signature HTML from `signature-new.html` or `signature-reply.html` automatically.
- **Importance**: `--importance Low|Normal|High` (default: Normal).
- **Contacts**: `contact search` searches org directory and recent email contacts. Returns name, email, and type.
- **Teams**: `teams list`, `teams show`, and `teams messages` browse Teams chats and messages read-only via Microsoft Graph.
- **Draft references**: `draft show` and `draft delete` accept a numeric index from `draft list`, a partial ID suffix, or a full ID.
- **Auth**: Built into this repo via `outlook_draft/auth.py`. Outlook features use the Outlook token, Teams uses the Microsoft Graph token. Run `outlook-cli auth` if expired, or `outlook-cli auth --headed` for a visible browser.

## File Layout

```
/home/ross/.local/lib/outlook-draft-cli/
  .env                          # MS_EMAIL and MS_PASSWORD, ignored by git
  .env.example                  # Example local config
  session_state/                # Local token cache, ignored by git
  outlook_draft/
    auth.py                     # Built-in Playwright auth and token capture
    cli.py                      # CLI entry point
    config.py                   # Config and local paths
    errors.py                   # Exceptions
    outlook_client.py           # Outlook REST API v2.0 client
    token_manager.py            # Token management + reauth trigger
```

## Caveats

- Uses the Outlook REST API v2.0, not Microsoft Graph.
- Token validity is ~24 hours. If expired, the CLI auto-triggers headless reauth (prints MFA number to console).
- Sending is intentionally not supported. Drafts are created for review in Outlook before sending.
- Email bodies are HTML-to-text converted for console display, truncated at 3000 chars.
