---
name: outlook-cli
description: Use when an agent needs to read/search Outlook mail, inspect calendar events, manage drafts, search contacts, list tasks, browse Teams chats, or work with OneDrive/SharePoint files through the local outlook-cli command. Use this for Microsoft 365 and Outlook automation where safe CLI operations are preferred.
---

# Outlook CLI

Use `outlook-cli` for Outlook/Microsoft 365 work: mail, drafts, calendar, contacts, tasks, Teams browsing, and files.

The CLI is human-first by default. Use `--json` for agent-readable output on read/search/list commands.

## Safety

- Prefer read-only commands before mutation.
- Sending email and Teams messages is intentionally disabled.
- Draft creation is allowed, but do not create drafts unless recipients, subject, and body are clear.
- Calendar/contact/task/file mutations should only be run when the user explicitly asks.
- Use cached numeric refs only after running the relevant list/search command in this session, or when cache freshness is acceptable.

## Auth

```bash
outlook-cli auth status --json
outlook-cli auth
outlook-cli auth --headed
```

## Mail

```bash
outlook-cli mail unread --json
outlook-cli mail search "colgate" --json
outlook-cli mail read 1 --json
outlook-cli mail folders --json
outlook-cli mail attachments 1 --json
outlook-cli mail links 1 --json
```

Archive or move only when explicitly requested:

```bash
outlook-cli mail archive 1
outlook-cli mail move 1 --folder Archive
```

## Drafts

```bash
outlook-cli draft create --to person@example.com --subject "Subject" --body "Body text"
outlook-cli draft reply 1 --body "Thanks, I will review and come back to you."
outlook-cli draft list
outlook-cli draft show 1
outlook-cli draft delete 1
```

## Calendar

```bash
outlook-cli cal agenda --json
outlook-cli cal show 1 --json
outlook-cli cal rooms --json
outlook-cli cal availability --attendee person@example.com --start "2026-06-01 09:00" --end "2026-06-01 17:00" --json
outlook-cli cal find-time --attendee person@example.com --start "2026-06-01 09:00" --end "2026-06-01 17:00" --json
```

Create/update/delete/accept/decline only when explicitly requested:

```bash
outlook-cli cal create "Subject" "2026-06-01 14:00" "2026-06-01 14:30" --attendee person@example.com
outlook-cli cal update 1 --start "2026-06-01 15:00" --end "2026-06-01 15:30"
outlook-cli cal delete 1
outlook-cli cal accept 1 --no-send-response
outlook-cli cal decline 1 --comment "Sorry, I cannot make it"
```

## Contacts And Tasks

```bash
outlook-cli contact search "ross" --json
outlook-cli task list --json
outlook-cli task create "Follow up with client"
outlook-cli task complete 1
```

## Teams

Teams browsing is read-only:

```bash
outlook-cli teams list
outlook-cli teams show 1
outlook-cli teams messages 1
outlook-cli teams attachments 1
```

## Files

Browse:

```bash
outlook-cli files sites
outlook-cli files list
outlook-cli files list --site "Client"
```

Download/upload/move/rename only when explicitly requested:

```bash
outlook-cli files download "Documents/report.pdf" ./downloads/
outlook-cli files upload ./deck.pptx "Shared Documents" --site "Client"
outlook-cli files mkdir "Documents/Reports"
outlook-cli files rename "Documents/old.pdf" "new.pdf"
outlook-cli files move "Documents/file.pdf" "Archive"
```
