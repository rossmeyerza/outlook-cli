# outlook-cli

## Overview

CLI for managing Outlook email, drafts, calendar, tasks, contacts, Teams browsing, and signatures via the Outlook REST API v2.0 and Microsoft Graph. Globally available as `outlook-cli`. Repo: `rossmeyerza/outlook-cli`.

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
outlook-cli files list                     # list OneDrive root
outlook-cli files list "Documents/Reports" # list OneDrive subfolder
outlook-cli files list --site "Tesco"      # list SharePoint site root
outlook-cli files list --site "Tesco" "Shared Documents"  # list SharePoint subfolder
outlook-cli files upload ./file.pdf "Documents"            # upload to OneDrive
outlook-cli files upload ./file.pdf "Shared Documents" --site "Tesco"  # upload to SharePoint
outlook-cli files mkdir "Documents/Q2"     # create OneDrive folder
outlook-cli files mkdir "Shared Docs/Q2" --site "Tesco"   # create SharePoint folder
outlook-cli files rename "Documents/old.pdf" "new.pdf"    # rename
outlook-cli files move "Documents/file.pdf" "Archive"     # move

# Signatures
outlook-cli signature fetch                # fetch new + reply signatures from OWA (no MFA after first auth)
outlook-cli signature fetch --headed       # run with visible browser if headless fails

# Auth and config
outlook-cli auth                           # force headless re-authentication
outlook-cli auth status                    # show token status
outlook-cli auth clear                     # delete local token cache
outlook-cli auth scopes                    # list safe token metadata/scopes
outlook-cli mailbox show                   # show mailbox settings
outlook-cli mailbox update --timezone "GMT Standard Time"
outlook-cli config check                   # validate local config without printing secrets
```

## Key Details

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
- **Teams**: `teams list`, `teams show`, and `teams messages` browse Teams chats and messages via Microsoft Graph. `teams list` sorts by latest received user message, ignoring system events and self messages where identifiable. Teams sending is intentionally disabled for agent safety.
- **Teams self-chat**: `teams self` checks access to the Microsoft Graph special self-chat thread, `48:notes`. This is the real Teams "chat with yourself" stream and is different from normal one-person `/me/chats` entries.
- **Teams gateway**: `gateway start --self-chat` watches `48:notes` for `@Marlow`, sends prompts to `pi --mode rpc`, and posts responses back into Teams. The gateway posts a short `...` receipt, soft-deletes it before the final response where Graph allows it, and can surface compact Pi tool progress. If the Graph token expires and headless reauth fails, the gateway records the auth error and exits instead of retrying every poll interval.
- **Gateway state**: Runtime state is outside git under `~/.local/share/outlook-cli/session_state/`. Pi sessions live under `session_state/gateway_sessions/<chat-hash>/`; per-chat workspaces live under `~/.local/share/outlook-cli/gateway_workspaces/<chat-hash>/`. These are local data/state, not tracked or pushed.
- **Gateway model controls**: Start with `outlook-cli gateway start --self-chat --model <model> [--provider <provider>] [--thinking <level>]`. Inside Teams use `@Marlow !model`, `@Marlow !model help`, `@Marlow !model list [search]`, `@Marlow !model <model>`, `@Marlow !model --provider <provider> --model <model> --thinking <level>`, or `@Marlow !model reset`. Short names such as `sonnet` rely on Pi fuzzy/pattern matching; use `!model list sonnet` and an exact model ID when repeatability matters.
- **Gateway file publishing**: Pi must not send Teams messages directly. To send generated files back to Teams, create files under the per-chat workspace and write `.marlow-export.json` with JSON like `{"files":["report.html"],"message":"Created the report."}`. The gateway validates relative paths, uploads supported file types to OneDrive under `Outlook CLI/Gateway/<chat-hash>/`, creates organization view links, posts them, and removes the manifest.
- **Gateway commands in Teams**: Use `@Marlow !help`, `@Marlow !commands`, or `@Marlow !help <command>` to list commands. Current commands include `!status`, `!new`, `!reset`, `!model`, `!pause`, `!resume`, `!tools`, `!files`, `!send`, and `!logs`.
- **Draft references**: `draft show` and `draft delete` accept a numeric index from `draft list`, a partial ID suffix, or a full ID.
- **Files**: `files sites` lists SharePoint sites via M365 group membership. `--site` does a case-insensitive partial name match. Without `--site`, operations target personal OneDrive. Uploads under 4 MB use a single PUT; larger files use chunked upload sessions. Uses the Microsoft Graph token.
- **Signatures**: `signature fetch` opens a headless browser with the saved OWA browser session (`session_state/browser_state.json`), intercepts the `OutlookCloudSettings/settings/account` API responses OWA fires on load, and saves the active new-message and reply signatures. No MFA after the first `auth`. On a fresh install, `auth` and `signature fetch` run automatically.
- **Auth**: Built into this repo via `outlook_draft/auth.py`. Outlook features use the Outlook token, Teams uses the Microsoft Graph token. `auth` saves both API tokens and the full browser session state. Run `outlook-cli auth` if expired, or `outlook-cli auth --headed` for a visible browser.

## File Layout

```
/home/ross/.local/lib/outlook-cli/
  .env                          # MS_EMAIL, MS_PASSWORD, timezone, signature path config, ignored by git
  .env.example                  # Example local config
  install.sh                    # Installer/updater script (served at outlook-cli.21436587.xyz)
  session_state/                # Local token + browser session cache, ignored by git
    tokens.json                 # API tokens
    browser_state.json          # Playwright browser session (used by signature fetch)
  outlook_draft/
    auth.py                     # Playwright auth, token capture, browser session save
    calendar_time.py            # Calendar timezone parsing and Outlook headers
    cli.py                      # CLI entry point
    signatures.py               # Signature loading and HTML sanitization
    config.py                   # Config and local paths
    errors.py                   # Exceptions
    links.py                    # SharePoint/OneDrive URL extraction and Graph share encoding
    outlook_client.py           # Outlook REST API v2.0 client
    token_manager.py            # Token management + reauth trigger
    commands/
      calendar.py               # Calendar subcommands
      contacts.py               # Contacts subcommands
      files.py                  # OneDrive and SharePoint file operations via Graph
      mail.py                   # Mail subcommands
      signature.py              # Signature fetch via OWA API interception
      tasks.py                  # Tasks subcommands
      teams.py                  # Teams read subcommands
```

## Caveats

- Uses the Outlook REST API v2.0, not Microsoft Graph.
- Token validity is ~24 hours. If expired, the CLI auto-triggers headless reauth (prints MFA number to console).
- Sending is intentionally not supported. Drafts are created for review in Outlook before sending.
- Email bodies are HTML-to-text converted for console display, truncated at 3000 chars.
