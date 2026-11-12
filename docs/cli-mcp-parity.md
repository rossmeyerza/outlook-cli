# CLI and MCP feature comparison

Checked on 11 November 2026 against the executable CLI parser, registered MCP tools and their implementations.
This table compares the CLI with the MCP server, not with the Microsoft service as a whole.

## Coverage definitions

- **Full:** The MCP provides the same main capability, sometimes through a different tool name or several calls.
- **Partial:** The MCP provides part of the capability, but some options or data are absent.
- **Missing:** An application capability exists in the CLI but has no dedicated MCP equivalent.
- **Local only:** Local account, authentication, bridge or process administration has no equivalent MCP tool.

Different output formats and cached numeric references do not determine coverage.
MCP tools normally require full IDs.
Safety controls can differ even when the main capability is full.
Presence in this table does not prove that the active account has the required permissions.

There are **95 executable parser paths** and **98 distinct operations** in the table.
`auth` uses four positional operation choices rather than four subparsers, which accounts for the difference.
The parser registers no command aliases.
The MCP server registers **74 tools**.
All **74 CLI operations** in mail, drafts, calendar, contacts, tasks, Teams, OneDrive and SharePoint have MCP equivalents.
Equivalent options can use different names or several tool calls.
Confirmation and local-path safeguards remain stricter in MCP.
Account administration, mailbox settings and Copilot are outside this parity change.

| Domain | CLI operations |
|---|---:|
| Account | 5 |
| Authentication | 4 |
| Configuration | 1 |
| Mail | 26 |
| Drafts | 8 |
| Calendar | 12 |
| Contacts | 4 |
| Tasks | 5 |
| Teams | 10 |
| Mailbox | 2 |
| Files | 9 |
| Copilot | 6 |
| Signatures | 1 |
| Gateway | 3 |
| Shell completion | 1 |
| MCP process | 1 |
| **Total** | **98** |

## Account, authentication and configuration

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `account list` | Local only | None | Lists local account profiles. |
| `account current` | Partial | `auth_status` | Returns the active account name, not profile paths or complete profile details. |
| `account add` | Local only | None | Creates credentials and local account configuration. |
| `account switch` | Local only | None | No MCP operation to switch its active account. |
| `account remove` | Local only | None | Can also delete local profile data. |
| `auth login` | Local only | None | Bare `auth` selects login; supports a visible browser through `--headed`. |
| `auth status` | Full | `auth_status` | Reports token presence, validity and remaining time. |
| `auth clear` | Local only | None | Deletes cached authentication. |
| `auth scopes` | Local only | None | Token scopes are not part of `auth_status`. |
| `config check` | Local only | None | Checks the local configuration. |

## Mail

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `mail unread` | Full | `mail_unread` | Caller controls the count; no extra MCP upper limit. |
| `mail search` | Full | `mail_search`, `mail_query` | The second tool provides filters, continuation URLs and all-page retrieval. |
| `mail read` | Full | `mail_read` | Optional HTML; the default body limit is 20,000 characters and the link list stops at 50. |
| `mail mark-read` | Full | `mail_mark_read` | Set `is_read=true`; the change is immediate. |
| `mail mark-unread` | Full | `mail_mark_read` | Set `is_read=false`; the change is immediate. |
| `mail archive` | Full | `mail_archive` | The change is immediate. |
| `mail move` | Full | `mail_move` | Accepts a well-known folder, display name or ID; rejects ambiguous names. |
| `mail folders` | Full | `mail_folders` | Caller controls the count; no extra MCP upper limit. |
| `mail create-folder` | Full | `mail_create_folder` | Optional `parent_id`; creates the folder immediately. |
| `mail attachments` | Full | `mail_attachments` | Returns IDs, sizes, types and inline status. |
| `mail download-attachments` | Full | `mail_download_attachments` | Absolute destination; optional attachment IDs, inline files and overwrite. |
| `mail links` | Full | `mail_links` | MCP defaults to share links; set `share_only=false` for all links. |
| `mail download-links` | Full | `mail_download_links` | Downloads share links; returns successful paths and explicit per-link errors. |
| `mail send` | Full | `mail_send` | Recipients, format, importance, attachments, signature selection and optional Sent Items copy; previews unless confirmed. |
| `mail conversation` | Full | `mail_conversation` | Returns accessible conversation messages, bodies and attachment metadata. |
| `mail bulk` | Full | `mail_bulk` | Archive, trash, move, mark-read and mark-unread; previews unless confirmed. |
| `mail trash` | Full | `mail_trash` | Moves to Deleted Items after confirmation; no permanent deletion. |
| `mail categories list` | Full | `mail_categories` | Requires Microsoft Graph mailbox-settings permission. |
| `mail categories create` | Full | `mail_category_create` | Confirmation required; colour is none or preset0 through preset24. |
| `mail categories update` | Full | `mail_category_update` | Changes colour, not the category name; confirmation required. |
| `mail categories delete` | Full | `mail_category_delete` | Confirmation required. |
| `mail categorise` | Full | `mail_categorise` | Adds or removes categories on full Graph message IDs; confirmation required. |
| `mail rules list` | Full | `mail_rules` | Requires Microsoft Graph mailbox-settings permission. |
| `mail rules create` | Full | `mail_rule_create` | Confirmation required; new rules start disabled and reject unsupported actions. |
| `mail rules update` | Full | `mail_rule_update` | Accepts name, order, enabled state, conditions and actions; confirmation required. |
| `mail rules delete` | Full | `mail_rule_delete` | Confirmation required. |

`mail_query` supports sender, recipient, date, attachment, unread, folder and page filters.
Its counts cover fetched results, not an assumed mailbox total.
Keyword searches can retain a limit that the server controls.
The last live check returned HTTP 403 for category and rule reads because the token lacked mailbox-settings permissions.
That permission failure is separate from MCP implementation coverage.

## Drafts

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `draft create` | Full | `draft_create` | Recipients, Markdown or HTML, importance, signature and attachments; creates without a send. |
| `draft reply` | Full | `draft_reply` | Reply-all, format, attachments and optional importance. |
| `draft update` | Full | `draft_update` | Replacement fields, format, importance, signature selection and attachments. |
| `draft forward` | Full | `draft_forward` | Recipients, format, attachments and optional importance. |
| `draft list` | Full | `draft_list` | Caller controls the count; no extra MCP upper limit. |
| `draft show` | Full | `draft_get` | Includes optional HTML. |
| `draft delete` | Full | `draft_delete` | Deletes immediately; rejects messages that are not drafts. |
| `draft send` | Full | `draft_send` | Preview by default; confirmation performs the send. |

Email attachments must not exceed 3 MiB each.
MCP composition accepts the body value rather than a CLI body-file path.
Markdown is the default input format.

## Calendar

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `cal agenda` | Full | `cal_agenda` | Caller controls the days and count; no extra MCP upper limits. |
| `cal show` | Full | `cal_get` | Recurrence, organiser status and attendees; body limit 10,000 characters. |
| `cal create` | Full | `cal_create` | Location, body, attendees and online meeting; invitations go out immediately when attendees exist. |
| `cal rooms` | Full | `cal_rooms` | Discovers rooms, optionally within a room list; tenant support and permissions still apply. |
| `cal availability` | Full | `cal_availability` | Interval, workday, weekdays, tentative treatment and optional raw schedule data. |
| `cal find-time` | Full | `cal_find_time` | Calculated slots or native Outlook suggestions through outlook=true. |
| `cal update` | Full | `cal_update` | Preview, confirmation and series controls; subject, times, location, body and online-meeting state. |
| `cal delete` | Full | `cal_delete`, `cal_cancel` | New guarded tools; organiser meetings with attendees require explicit cancellation through `cal_cancel`. |
| `cal accept` | Full | `cal_respond` | Use response accept; supports comment and the send-response flag. |
| `cal decline` | Full | `cal_respond` | Use response decline; supports comment and the send-response flag. |
| `cal tentative` | Full | `cal_respond` | Use response tentative; supports comment and the send-response flag. |
| `cal cancel` | Full | `cal_cancel` | New tool; organiser-only, with preview and explicit confirmation of attendee notifications. |

Calendar update, delete and cancel tools preview by default.
A whole-series operation also requires `allow_series=true` and explicit approval for the whole series.
An occurrence ID affects only that occurrence; the tool does not follow its series-master ID automatically.
An optional `expected_change_key` rejects a changed event after a preview.
These tools disable automatic retries for writes.
The CLI update, delete and cancel commands retain their existing immediate behaviour.

`cal_respond` defaults to send_response true, which emails the organiser.
It has no separate confirmation argument.
Calendar deletion does not delete an invitation email from the mailbox.

## Contacts

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `contact search` | Full | `contact_search` | Directory and recent contacts; caller controls the count. |
| `contact org` | Full | `contact_org` | Company, department, title, manager and all direct reports; optional management chain and explicit completeness flags. |
| `contact create` | Full | `contact_create` | Personal contact fields; previews unless confirmed. |
| `contact update` | Full | `contact_update` | Personal contact fields; previews unless confirmed; a supplied email replaces the email collection. |

## Tasks

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `task list` | Full | `task_list` | Incomplete tasks; caller controls the count. |
| `task create` | Full | `task_create` | Creates a task with its subject. |
| `task update` | Full | `task_update` | Subject, due date, importance and status; previews unless confirmed. |
| `task complete` | Full | `task_complete` | Completes the task immediately. |
| `task delete` | Full | `task_delete` | Previews the exact task before confirmed deletion. |

## Teams

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `teams list` | Full | `teams_list` | Caller controls the count; optional latest-received sort within fetched chats. |
| `teams search` | Full | `teams_search` | Caller controls the result count, scan and member count; includes one-on-one restriction. |
| `teams with` | Full | `teams_search`, `teams_messages` | Search with one_on_one true; read messages after a unique match. |
| `teams show` | Full | `teams_get` | Chat metadata, created time and members, without a message fetch. |
| `teams self` | Full | `teams_self` | Self-chat discovery, access check and member details. |
| `teams messages` | Full | `teams_messages` | Caller controls the count; uses client pagination beyond 50 messages. |
| `teams attachments` | Full | `teams_attachments` | Caller controls the message scan; no extra MCP upper limit. |
| `teams download-attachments` | Full | `teams_download_attachments` | Selector and overwrite; caller controls the scan; absolute destination required. |
| `teams send` | Full | `teams_send` | Message, HTML, attachments, inline images and upload folder; previews unless confirmed. |
| `teams new` | Full | `teams_new` | Recipients, group topic and media; previews unless confirmed. |

A file preview resolves the existing folder and the intended recipients without remote changes.
Confirmation permits file uploads and read grants to internal members as well as the message send.
Existing owner and administrator access remains unchanged.
A new group chat is not automatically reusable.
Do not repeat an uncertain group creation or message send.

## Mailbox

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `mailbox show` | Missing | None | No mailbox-settings read tool. |
| `mailbox update` | Missing | None | No timezone or automatic-reply settings tool. |

## Files

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `files sites` | Full | `files_sites` | Lists sites through Microsoft 365 group membership. |
| `files libraries` | Full | `files_libraries` | Resolves a site and lists its libraries. |
| `files list` | Full | `files_list` | OneDrive paths and SharePoint site or library selection. |
| `files search` | Full | `files_search` | SharePoint site required; optional library and caller-controlled result count. |
| `files upload` | Full | `files_upload` | Preview of file and exact destination; confirmation required; replacement also requires overwrite=true. |
| `files download` | Full | `files_download` | Paths, item IDs and SharePoint/OneDrive URLs; absolute local destination and optional overwrite. |
| `files mkdir` | Full | `files_mkdir` | Preview of one folder in an existing parent; confirmation required. |
| `files rename` | Full | `files_rename` | Preview of the exact source and new name; confirmation required. |
| `files move` | Full | `files_move` | Preview of source and destination within one drive; confirmation required. |

MCP downloads reuse the CLI's share-URL and OneDrive ID resolution.
New file tools resolve the selected site, library and folder before a change.
Upload defaults reject replacement; explicit overwrite approval permits it.
Small uploads and upload sessions use an explicit conflict policy.
New contact, task, calendar and file changes have no automatic retries.
Uploads and downloads retain the CLI's in-memory file buffers.

## Copilot

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `copilot ask` | Partial | `copilot_ask` | Files, site, library and explicit conversation ID; no cached continue, timeout or stream controls. |
| `copilot app status` | Local only | None | No bridge-status tool. |
| `copilot app inspect` | Local only | None | No bridge-inspection tool. |
| `copilot app reset` | Local only | None | No operation to clear an uncertain tab state. |
| `copilot app ask` | Partial | `copilot_app` | File, app, site, library, new chat, open tab and save wait; no timeout option. |
| `copilot extension install` | Local only | None | Installs the local browser extension and native host. |

Neither Copilot tool is read-only.
Copilot can change remote data, and the app tool can edit shared documents immediately.
Neither tool has a confirmation flag that makes arbitrary prompts safe.
Approval detection remains a separate limitation.

## Signatures, gateway and local infrastructure

| CLI command | MCP coverage | MCP tool | Difference or limit |
|---|---|---|---|
| `signature fetch` | Local only | None | Retrieves signatures through a browser session. |
| `gateway start` | Local only | None | Starts the Teams gateway and selects its model and chat. |
| `gateway stop` | Local only | None | Stops the local gateway. |
| `gateway status` | Local only | None | Reports local gateway state. |
| `completion` | Local only | None | Accepts bash, zsh, fish, powershell or nushell as an argument. |
| `mcp` | Local only | None | Starts the stdio MCP process; it is not an MCP operation. |

## Names and aliases

The parser registers no command aliases.
Bare `auth` and `auth login` select the same operation.
`teams with` is a separate command, not a registered alias.
Shell names for completion are argument values rather than separate commands.
Short flags, colour/color and the compatibility HTML flag do not increase the command count.

No registered MCP tool lacks a CLI capability equivalent.
Some MCP tool names differ from the corresponding CLI commands:

| MCP tool | CLI equivalent |
|---|---|
| `mail_query` | `mail search` with filters, page or all-page options |
| `cal_respond` | `cal accept`, `cal tentative`, `cal decline` |
| `mail_mark_read` with is_read false | `mail mark-unread` |
| `draft_get` | `draft show` |
| `cal_get` | `cal show` |
| `copilot_app` | `copilot app ask` |

## General differences

CLI sends and guarded bulk operations use `--yes`; MCP equivalents use `confirm=true`.
Some MCP changes remain immediate: draft creation, updates and deletion; individual mail changes; tasks; calendar creation and invitation responses.
The MCP server normally does not start interactive authentication.
Copilot can refresh its token through the saved browser session.

The CLI has global read-only, no-send, dry-run and command restrictions.
The MCP server has equivalent process-level read-only, no-send and tool restrictions through environment variables.
It has no general per-call dry-run flag.
Download tools that have no remote write still create local files.
Copilot does not reliably replace a missing dedicated tool.

## Verification source

The comparison uses `outlook_draft/cli.py`, every module under `outlook_draft/commands/`, and the modules in `mcp/server.py::MODULES`.
The table contains no live mailbox counts.
Synthetic tests verify the new tools, options, confirmations and policies without live mutations.
The table tests require coverage for every executable CLI operation and every registered MCP tool.
