# Microsoft 365 Copilot chat tools

This file lists the tools that the Microsoft 365 Copilot work chat can use when `outlook-cli copilot ask` or the `copilot_ask` MCP tool sends it a prompt.

Source: Copilot gave this list itself, in the work chat, on 2026-10-22. The list is self-reported. The real tool names and parameters can be different. The tenant and the licence can add or remove tools, and Microsoft can change them at any time. Use this list to understand what a prompt can cause. Do not use it as an API reference.

## Important points for this repo

- **`copilot_ask` is not read-only.** Copilot can create drafts and events, post Teams messages, delete files, share files and change calendars. A prompt to `copilot ask` can start these actions. The MCP tool has the annotations `readOnlyHint=false` and `destructiveHint=true`.
- **The CLI shows each tool call.** The answer lists the tools that Copilot called, with their arguments (`toolCalls` in `--json` and in the MCP result). The list does not include tool discovery (`api_tool_list_resources`) or language hints.
- **The client never confirms an action.** If Copilot sends a confirmation, input or sign-in request (`ConfirmationCard`, `TriggerConfirmation`, `TriggerUserInputRequest`, `TriggerPluginAuth`), the client stops, reports it as `pendingConfirmations`, and does not confirm it.
- **Document edits make a new file.** `edit_docx`, `edit_presentation`, `edit_workbook` and `edit_pdf` return a new file. They do not change the original SharePoint file. Copilot reports that it has no direct edit of a cell in a SharePoint Excel file.
- **The CLI returns only text.** `copilot ask` does not yet return or download the files that Copilot creates.
- **Use the guarded CLI commands to send.** To send email or Teams messages, use `mail send`, `draft send` or `teams send`. These commands preview and need `--yes`. The MCP send tools need `confirm=true`.

## Test results (2026-10-22)

These tests used only Ross's own mailbox.

| Prompt | Result |
|---|---|
| Create a draft email to self | Copilot called `outlook_mail_draft_message` and created the draft. There was no confirmation step. |
| Send an email to self now | Copilot did not send. Its Outlook tool only creates drafts. In the web app, the user sends with the Send button in the compose widget, which this client does not have. |
| Find the latest email about a topic | Copilot called `office365_search` and `search_office365`. |

The real tool names on the wire are different from the self-reported names below. For example, the draft tool is `outlook_mail_draft_message`, not `draft_email`. Copilot finds its tools at run time with `api_tool_list_resources`. The Outlook tools run on an MCP server at `https://substrate.office.com/outlook-mcp/mcp`.

The client did not get a confirmation card in these tests, so the format of the card is not known. Actions on Teams, SharePoint files and the calendar are not tested.

## Tool list

### Microsoft 365 search and retrieval

```text
search_m365(domain, query, filters?, response_length?)
    domain = files | emails | chats | meetings | people | transcripts
open_m365_item(reference_id)
find_in_m365_item(reference_id, patterns[])
```

### SharePoint

```text
find_sharepoint_site(search_query?)
find_sharepoint_file(file_name)
get_sharepoint_file_metadata(file_id, document_library_id)
list_sharepoint_folder(document_library_id, folder_id?)
copy_sharepoint_file(source_file_id, source_library_id, destination_library_id, destination_folder_id)
move_sharepoint_file(source_file_id, source_library_id, destination_library_id, destination_folder_id)
rename_sharepoint_file(file_id, library_id, new_name)
delete_sharepoint_file(file_id, library_id)
share_sharepoint_file(file_id, library_id, recipients[], permissions[])
set_file_sensitivity_label(file_id, library_id, label_id)
create_sharepoint_list(site_id, name, columns?)
update_sharepoint_list_item(site_id, list_id, item_id, fields)
update_sharepoint_column(site_id, list_id, column_id, properties)
```

### Excel

```text
create_workbook(specification) -> .xlsx
edit_workbook(input_xlsx, edits) -> .xlsx
analyze_workbook(input_xlsx, analysis_request)
recalculate_workbook(workbook)
render_workbook(workbook)
```

An edit can target a sheet and a range with a formula, a format or a value operation. The edit applies to a workbook file (an artifact). It does not apply to a cell in a SharePoint file.

### Word

```text
create_docx(content, formatting?) -> .docx
edit_docx(input_docx, edits) -> .docx
render_docx(docx)
```

### PowerPoint

```text
create_presentation(content, design?) -> .pptx
edit_presentation(input_pptx, edits) -> .pptx
render_presentation(pptx)
```

### PDF

```text
create_pdf(content, layout?) -> .pdf
edit_pdf(input_pdf, edits) -> .pdf
```

### Code sandbox

```text
exec_shell(command)
exec_python(code)
open_image(file_path)
download_supported_asset(url, destination)
```

The sandbox keeps files in `/mnt/data`.

### Image generation

```text
generate_image(prompt, orientation?, transparent_background?)
edit_image(image, edit_request)
```

### Public web

```text
web_search(query)
open_web_page(url_or_search_result)
find_on_web_page(page, text)
follow_web_link(page, link)
```

### Outlook email

Email retrieval uses `search_m365("emails", query)` and `open_m365_item(email_reference)`.

```text
draft_email(recipients, subject, body)
send_email(recipients, subject, body)
reply_email(email, body)
forward_email(email, recipients, body?)
delete_email(email_refs[])
archive_email(email_refs[])
move_email(email_refs[], folder)
copy_email(email_refs[], folder)
mark_email_read(email_refs[])
mark_email_unread(email_refs[])
categorize_email(email_refs[], categories[])
pin_email(email_refs[])
unpin_email(email_refs[])
flag_email(email_refs[], due_date?, reminder_date?)
unflag_email(email_refs[])
complete_email_flag(email_refs[])
mark_email_junk(email_refs[])
mark_email_not_junk(email_refs[])
```

### Outlook rules

```text
get_inbox_rules()
create_inbox_rule(rule)
update_inbox_rule(rule_id, changes)
delete_inbox_rule(rule_ids[])
```

### Calendar and meetings

Meeting search uses `search_m365("meetings", query)`.

```text
find_availability(attendees[], date_range, duration, constraints?)
create_calendar_event(subject, start, duration?, attendees[], options?)
accept_meeting(meeting)
decline_meeting(meeting)
tentatively_accept_meeting(meeting)
follow_meeting(meeting)
delete_calendar_event(meeting)
find_meeting_room(time, requirements?)
create_scheduling_poll(attendees[], candidate_times[])
get_scheduling_poll(poll)
finalize_scheduling_poll(poll, selected_time)
manage_calendar_categories(operation, category)
create_calendar_instruction(rule)
list_calendar_instructions()
delete_calendar_instruction(rule)
get_calendar_instruction_activity()
```

### Workplace and room location

```text
get_workplace_location(person, date?)
find_building(name)
find_rooms(building, floor?)
find_desks(building, floor?)
get_room_availability(room, dates)
```

### Teams

Chat search uses `search_m365("chats", query)`.

```text
list_joined_teams()
list_team_channels(team)
get_teams_thread(chat_or_channel)
create_teams_chat(members[], topic?)
send_teams_message(chat, content)
post_teams_channel_message(team, channel, content)
reply_to_teams_message(thread, message, content)
edit_teams_message(message, content)
delete_teams_message(message)
add_teams_chat_member(chat, person)
add_teams_channel_member(team, channel, person)
update_teams_channel(team, channel, changes)
```

### People

```text
search_m365("people", query)
open_person(person_reference)
```

### Planner and tasks

```text
search_planner_plans(scope)
search_planner_tasks(filters?)
```

Copilot reports that it can get more Planner actions when a request needs them.

### Enterprise connectors

```text
search_viva_engage(query, filters?)
search_viva_learning(query, filters?)
```

Copilot can find other tenant connectors at run time.

### Microsoft product help

```text
search_microsoft_product_help(question)
```

### Earlier Copilot conversations

```text
search_previous_conversations(keywords[])
```

### Pages

```text
create_page(name, content, type)
```

### Memory

```text
remember_user_information(information)
forget_user_information(information)
```

### Long-running tasks

```text
create_background_task(name, request)
get_background_task(task)
update_background_task(task, instruction)
```

### Tool discovery

```text
discover_tools(service_or_capability, query?)
```

## Tools with side effects

These tools change data or send content to other people. A prompt to `copilot ask` can cause them.

| Area | Tools |
|---|---|
| Email | `send_email`, `reply_email`, `forward_email`, `delete_email`, `move_email`, `mark_email_junk` |
| Outlook rules | `create_inbox_rule`, `update_inbox_rule`, `delete_inbox_rule` |
| Calendar | `create_calendar_event`, `delete_calendar_event`, `accept_meeting`, `decline_meeting`, `finalize_scheduling_poll` |
| Teams | `send_teams_message`, `post_teams_channel_message`, `reply_to_teams_message`, `edit_teams_message`, `delete_teams_message`, `create_teams_chat`, member and channel changes |
| SharePoint | `delete_sharepoint_file`, `move_sharepoint_file`, `rename_sharepoint_file`, `share_sharepoint_file`, `set_file_sensitivity_label`, list and column changes |
| Memory | `remember_user_information`, `forget_user_information` |
| Background work | `create_background_task`, `update_background_task` |
