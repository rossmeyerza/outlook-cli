# Mail features

The CLI and MCP server support the features below.
These commands do not require permanent email deletion.

## Bulk actions and trash

```bash
outlook-cli mail bulk archive <id1> <id2>
outlook-cli mail bulk archive <id1> <id2> --yes
outlook-cli mail bulk move <id1> <id2> --folder "Clients" --yes
outlook-cli mail bulk mark-read <id1> <id2> --yes
outlook-cli mail trash <id1> <id2>
outlook-cli mail trash <id1> <id2> --yes
```

The first command shows a preview.
Add `--yes` only after you check the exact message list and action.
The commands accept full message IDs or cached references.
The CLI resolves folder names and indices from `mail folders`.
The MCP tools require full IDs.

`trash` moves an email to Deleted Items.
It does not permanently delete the email.
The supported bulk actions are `archive`, `trash`, `move`, `mark-read` and `mark-unread`.
Each result includes the original message ID and its status.
A move can return a new message ID.

Caution: A failed request can have an uncertain result.
Check the mailbox before you retry a failed or uncertain write.
The operation stops after its first failed write.
The result identifies messages that the CLI did not attempt.
Bulk writes have no automatic retries.

MCP tools: `mail_bulk`, `mail_trash`.
Both tools preview by default.
Set `confirm=true` only after the user approves the change.

## Complete conversations

```bash
outlook-cli mail conversation <message-id> --json
```

The command accepts any message in the conversation.
It follows every available page and returns messages in date order.
It includes message bodies and attachment metadata.
It does not download attachment bytes.
The result covers messages that the active account can access.
It does not reconstruct deleted or inaccessible messages.

MCP tool: `mail_conversation`.
The MCP tool returns plain-text bodies.

## Search filters and pages

```bash
outlook-cli mail search --from person@example.com --after 2026-11-01 --all --json
outlook-cli mail search --to person@example.com --has-attachments --folder inbox --all --json
outlook-cli mail search --unread --before 2026-11-01 -n 100 --json
outlook-cli mail search "budget" --all --json
```

Simple keyword searches keep the existing output format.
Searches with filters, `--all` or `--page` return a report with page and count fields.
`-n` sets the page size, from 1 to 1000, for these searches.
Filters apply to fetched messages, not to an assumed mailbox total.
An address filter matches an exact address without regard to letter case.
`--to` checks both To and CC recipients.

`--after` is inclusive.
`--before` is exclusive.
Dates and timestamps without a time zone use UTC.
`--all` follows every available page.
A keyword search can still have a limit that the server controls.

The report includes these fields:

- `matchedCount`: the number of returned messages after the filters.
- `scannedCount`: the number of fetched messages before the filters.
- `nextPage`: the next URL, or null when no next page exists.
- `complete`: true only for an exhaustive structured query from its first page.
- `countExact`: true only when the count covers that complete query.
- `warning`: a description of any known coverage limit.

Use `--page '<nextPage>'` to continue a query.
Keep its keyword, folder and filters unchanged.
A continuation result counts only the messages that it returns.
The client rejects continuation URLs outside the active mail API.
The search cache contains only messages from the latest result.

MCP tool: `mail_query`.
Its arguments include `sender`, `recipient`, `after`, `before`, `has_attachments`, `unread`, `folder_id`, `page`, `all_pages` and `page_size`.

## Attachments and draft edits

```bash
outlook-cli draft create --to person@example.com -s "Report" -b "Please review." --attach ./report.pdf
outlook-cli draft reply <message-id> -b "Please see the file." --attach ./report.pdf
outlook-cli draft update <draft-id> -s "Updated subject" --attach ./report.pdf --json
outlook-cli draft update <draft-id> --body-file ./body.html --html --json
outlook-cli draft update <draft-id> --cc "" --json
outlook-cli mail send --to person@example.com -s "Report" -b "Please review." --attach ./report.pdf
```

`--attach` is repeatable.
Each attachment must be a regular local file of at most 3 MiB.
The CLI validates all files before it creates or changes a draft.
Larger attachment uploads are not supported.
The server can impose a lower total size for the request or message.
A send still requires `--yes`.

A draft update changes only the supplied fields.
An empty value clears a supplied field.
An omitted body stays unchanged.
A replacement body has no extra signature unless you request `--signature new` or `--signature reply`.
The client refuses an update or attachment upload when the message is not a draft.

Caution: A draft can remain saved after an attachment or field update fails.
Check the draft ID in the error before any retry.
The CLI does not delete a partial draft.
It does not automatically retry an attachment upload.

MCP tools: `draft_create`, `draft_reply`, `draft_update` and `mail_send` accept an `attachments` list of local paths.

## Forward drafts

```bash
outlook-cli draft forward <message-id> --to person@example.com -b "For your review." --json
```

A forward preserves the quoted original that Outlook supplies.
Outlook also supplies the original attachments.
The CLI adds the reply signature before the quoted original.
Use `--attach` to add other local files.
The command saves a draft and does not send it.

MCP tool: `draft_forward`.

## Categories and inbox rules

These features use Microsoft Graph v1.0, not the older Outlook REST mail API.
They require a Graph token with the relevant mailbox permissions.
A valid Teams token does not necessarily include these permissions.
The client reports a permission error instead of an undocumented API fallback.

| Operation | Required delegated permission |
|---|---|
| List categories or rules | `MailboxSettings.Read` or `MailboxSettings.ReadWrite` |
| Create, change or delete categories or rules | `MailboxSettings.ReadWrite` |
| Read message categories | `Mail.Read` or `Mail.ReadWrite` |
| Change message categories | `Mail.ReadWrite` |

The current local Graph token returned HTTP 403 for category and rule reads during the implementation check.
A token with the additional permissions is required before live use.
The normal browser token capture does not guarantee those permissions.
The tenant can require an administrator to approve them.

```bash
outlook-cli mail categories list --json
outlook-cli mail categories create "Client work" --colour preset4
outlook-cli mail categories create "Client work" --colour preset4 --yes
outlook-cli mail categories update <category-id> --colour preset7 --yes
outlook-cli mail categories delete <category-id> --yes
outlook-cli mail categorise <graph-message-id> --add "Client work"
outlook-cli mail categorise <graph-message-id> --add "Client work" --yes
outlook-cli mail rules list --json
outlook-cli mail rules create "Client mail" --conditions '{"subjectContains":["Client"]}' --actions '{"assignCategories":["Client work"],"stopProcessingRules":true}'
outlook-cli mail rules create "Client mail" --conditions '{"subjectContains":["Client"]}' --actions '{"assignCategories":["Client work"]}' --yes
outlook-cli mail rules update <rule-id> --enabled --yes
outlook-cli mail rules delete <rule-id> --yes
```

Category colours accept `none` and `preset0` through `preset24`.
Microsoft Graph does not permit a master category name change.
Category and rule changes preview by default.
`categorise` adds or removes only the requested categories.
It preserves unrelated categories and uses an ETag when Graph returns one.

Use full Graph IDs for category, rule, folder and message references in these commands.
Cached CLI indices are not Graph IDs.
Outlook REST IDs can differ from Graph IDs.
The CLI does not translate them automatically.

New rules start disabled.
Use `--enabled` explicitly to enable a rule.
A rule must have a sender or subject condition.
The supported conditions are `fromAddresses` and `subjectContains`.
The supported actions are `moveToFolder`, `assignCategories`, `markAsRead` and `stopProcessingRules`.
Unknown fields, automatic forward actions and delete actions fail validation.
The CLI checks the complete effective rule before it enables or changes an active rule.

MCP tools: `mail_categories`, `mail_category_create`, `mail_category_update`, `mail_category_delete`, `mail_categorise`, `mail_rules`, `mail_rule_create`, `mail_rule_update`, `mail_rule_delete`.
Mutation tools require `confirm=true` to apply their changes.

## Safety controls

```bash
outlook-cli --readonly mail search --unread --all --json
outlook-cli mail trash <id1> <id2> --yes --dry-run
outlook-cli --no-send draft create --to person@example.com -s "Review" -b "Draft only."
outlook-cli --enable-commands mail.search,mail.read mail search "budget"
outlook-cli --disable-commands mail.send,draft.send mail search "budget"
```

Safety flags work before or after a command.
`--readonly` blocks remote write requests even when a command has `--yes`.
`--no-send` blocks message sends and calendar changes that can send invitations.
Both modes block Copilot prompts because Copilot can change remote data outside the mail HTTP client.
Both modes also block gateway starts because child processes do not inherit the CLI policy.
Neither mode changes the account's saved permissions.

`--dry-run` prints an operation plan before it creates a client or authenticates.
It does not validate remote IDs, resolve recipients or read local attachment files.
Use the operation's normal preview when you need those checks.
The plan excludes body text and credentials.
A dry run does not claim that the operation will succeed.

Command restrictions accept comma-separated prefixes.
A disabled prefix takes precedence over an enabled prefix.
Dots or spaces separate command names, for example `mail.rules.create`.

The MCP server also supports process-level controls:

- `OUTLOOK_CLI_READONLY=1` hides mutating tools and blocks write requests.
- `OUTLOOK_CLI_NO_SEND=1` hides send and Copilot tools and blocks send requests.
- `OUTLOOK_CLI_ENABLE_TOOLS=mail_query,mail_conversation` exposes only the named tools.
- `OUTLOOK_CLI_DISABLE_TOOLS=mail_send,draft_send` hides the named tools.

Tool restrictions use exact MCP names.
Restart the MCP server after you change its environment or code.
Remote read-only mode does not prevent local attachment downloads.
