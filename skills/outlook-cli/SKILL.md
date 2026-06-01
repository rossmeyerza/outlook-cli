---
name: outlook-cli
description: "Use when an agent needs to work with Microsoft 365 through the local outlook-cli: Outlook mail search/read/drafts/attachments/links, calendar agenda/events/availability/find-time/rooms/invitations, contacts, tasks, Teams chats/messages/attachments, OneDrive/SharePoint files, mailbox settings, signatures, or auth/config checks."
---

# Outlook CLI

Use `outlook-cli` for Microsoft 365 work. It is installed globally and authenticated for the local user.

Default output is human tables. Prefer `--json` for agent reads when a command supports it. `--table` is useful when answering a human directly.

## Operating Rules

- Start read-only. Mutating commands require explicit user intent.
- Sending real email and Teams messages is intentionally disabled; create drafts instead of sending.
- Use cached numeric indexes only after running the matching list/search command, or when cache freshness is acceptable.
- For ambiguous people, search contacts first, then use the confirmed email address.
- For calendar work, use exact local times like `"2026-06-01 14:00"`; the CLI uses configured local/Outlook time zones.
- Config lives at `~/.config/outlook-cli/.env`.
- Runtime state lives under `~/.local/share/outlook-cli/`, especially `session_state/tokens.json` and `session_state/browser_state.json`.

## Health And Auth

Use these before deeper work if auth or config may be stale:

```bash
outlook-cli config check --json
outlook-cli auth status --json
outlook-cli auth scopes --json
```

Re-authentication:

```bash
outlook-cli auth
outlook-cli auth --headed
outlook-cli auth clear
```

Headless auth enters the configured email/password, prints the MFA challenge number when Okta/Microsoft shows one, and waits for approval. If Okta rejects credentials before MFA, update `MS_PASSWORD` in `~/.config/outlook-cli/.env`.

## Mail

Read/search commands:

```bash
outlook-cli mail unread --count 50 --json
outlook-cli mail search "search terms" --count 20 --json
outlook-cli mail read 1 --json
outlook-cli mail folders --count 100 --json
```

Message identifiers can be:
- a numeric index from the most recent unread/search result
- a cached ID suffix
- a full Outlook message ID

Useful inspection:

```bash
outlook-cli mail attachments 1 --json
outlook-cli mail links 1 --json
outlook-cli mail links 1 --share-only --json
```

Download only when requested:

```bash
outlook-cli mail download-attachments 1 --dir ./attachments
outlook-cli mail download-attachments 1 --dir ./attachments --include-inline
outlook-cli mail download-links 1 --dir ./attachments
```

Mutations require explicit user approval:

```bash
outlook-cli mail mark-read 1
outlook-cli mail mark-unread 1
outlook-cli mail archive 1
outlook-cli mail move 1 --folder Archive
```

Typical workflows:
- “Find the latest email about X”: `mail search "X" --json`, then `mail read <index> --json`.
- “Get the deck from that email”: `mail attachments <index> --json`, then download the requested attachment.
- “Find SharePoint links in that email”: `mail links <index> --share-only --json`.

## Drafts

Drafts are the safe path for outbound email. Do not send.

```bash
outlook-cli draft create --to person@example.com --subject "Subject" --body "Body text"
outlook-cli draft create --to person@example.com --cc other@example.com --subject "Subject" --body-file ./body.html --html
outlook-cli draft reply 1 --body "Thanks, I will review and come back to you."
outlook-cli draft reply 1 --reply-all --body-file ./reply.txt
outlook-cli draft list --count 20
outlook-cli draft show 1
outlook-cli draft delete 1
```

For generated email text, confirm recipients, subject, and body intent before creating the draft if any part is ambiguous.

## Calendar

Read agenda and event details:

```bash
outlook-cli cal agenda --days 7 --count 20 --json
outlook-cli cal agenda --days 14 --plain
outlook-cli cal show 1 --json
```

Availability/free-busy:

```bash
outlook-cli cal availability \
  --attendee person@example.com \
  --start "2026-06-01 09:00" \
  --end "2026-06-01 17:00" \
  --interval 30 \
  --json
```

For multiple people, repeat `--attendee`:

```bash
outlook-cli cal availability \
  --attendee a@example.com \
  --attendee b@example.com \
  --start "2026-06-01 09:00" \
  --end "2026-06-01 17:00" \
  --json
```

Find meeting slots:

```bash
outlook-cli cal find-time \
  --attendee person@example.com \
  --start "2026-06-01 09:00" \
  --end "2026-06-01 17:00" \
  --duration 30 \
  --count 10 \
  --json
```

Rooms:

```bash
outlook-cli cal rooms --json
outlook-cli cal rooms --room-list rooms@example.com --json
```

Create/update events only when requested:

```bash
outlook-cli cal create "Subject" "2026-06-01 14:00" "2026-06-01 14:30" \
  --attendee person@example.com \
  --location "Meeting room" \
  --body "Agenda"

outlook-cli cal create "Subject" "2026-06-01 14:00" "2026-06-01 14:30" \
  --attendee person@example.com \
  --teams

outlook-cli cal update 1 --start "2026-06-01 15:00" --end "2026-06-01 15:30"
outlook-cli cal update 1 --teams
outlook-cli cal update 1 --no-teams
outlook-cli cal delete 1
```

Invitation responses:

```bash
outlook-cli cal accept 1 --no-send-response
outlook-cli cal tentative 1 --comment "Tentative for now"
outlook-cli cal decline 1 --comment "Sorry, I cannot make it"
outlook-cli cal cancel 1 --comment "No longer needed"
```

Calendar decision pattern:
- To answer “when is X free?”, use `cal availability`.
- To propose meeting times, use `cal find-time`.
- To inspect current schedule, use `cal agenda`, then `cal show`.
- To book a meeting, first confirm attendees, subject, date/time, duration, Teams/location, and body.

## Contacts

```bash
outlook-cli contact search "name or email" --count 10 --json
outlook-cli contact create --name "Full Name" --email person@example.com --company "Company"
outlook-cli contact update 1 --mobile "+44..."
```

Use `contact search` before emailing or scheduling if the user gives only a name.

## Tasks

```bash
outlook-cli task list --count 50 --json
outlook-cli task create "Follow up with client"
outlook-cli task update 1 --subject "New subject" --due "2026-06-03 17:00" --importance High
outlook-cli task update 1 --status InProgress
outlook-cli task complete 1
outlook-cli task delete 1
```

Valid task statuses: `NotStarted`, `InProgress`, `Completed`, `WaitingOnOthers`, `Deferred`.

## Teams

Teams support is for browsing and downloading attachments/links. Sending Teams messages is disabled.

```bash
outlook-cli teams list --count 20
outlook-cli teams list --count 20 --sort-received
outlook-cli teams show 1
outlook-cli teams messages 1 --count 20
outlook-cli teams attachments 1 --scan 50
```

Download Teams attachments or SharePoint/OneDrive links only when requested:

```bash
outlook-cli teams download-attachments 1 --scan 50 --dir ./attachments
outlook-cli teams download-attachments 1 --attachment 2 --dir ./attachments
outlook-cli teams download-attachments 1 --attachment "filename-or-url-substring" --overwrite
```

Teams identifiers can be a chat index from `teams list`, cached ID suffix, or full chat ID.

## Files: OneDrive And SharePoint

Browse:

```bash
outlook-cli files sites
outlook-cli files list
outlook-cli files list "Documents"
outlook-cli files list "Shared Documents" --site "Client"
```

Download/upload/mutate only when requested:

```bash
outlook-cli files download "Documents/report.pdf" ./downloads/
outlook-cli files download "Shared Documents/report.pdf" ./downloads/ --site "Client" --overwrite
outlook-cli files upload ./deck.pptx "Shared Documents" --site "Client"
outlook-cli files mkdir "Documents/Reports"
outlook-cli files rename "Documents/old.pdf" "new.pdf"
outlook-cli files move "Documents/file.pdf" "Archive"
```

Use `files sites` to discover SharePoint site names before using `--site`. Site matching is partial by name.

## Mailbox And Signatures

Mailbox settings:

```bash
outlook-cli mailbox show --json
outlook-cli mailbox update --timezone "GMT Standard Time"
outlook-cli mailbox update --auto-reply-status Disabled
outlook-cli mailbox update --auto-reply-status AlwaysEnabled \
  --internal-reply "Internal message" \
  --external-reply "External message"
```

Signature capture from OWA:

```bash
outlook-cli signature fetch
outlook-cli signature fetch --headed
```

Signature files default to:
- `~/.local/share/outlook-cli/signature-new.html`
- `~/.local/share/outlook-cli/signature-reply.html`

## Output Notes

- `--json` works broadly for mail, calendar, contacts, tasks, mailbox, auth, and config.
- Some Teams/files commands are table/text oriented; parse cautiously or run `--help` for the exact command if unsure.
- Top-level flags usually work before the domain (`outlook-cli --json cal agenda`) and many commands also accept `--json` after the subcommand (`outlook-cli cal agenda --json`).
- Use `--help` on any command before mutating if argument order is uncertain.
