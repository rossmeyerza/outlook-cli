# Teams files and inline images

The CLI and MCP server support local files and inline images in Teams chat messages.
Both existing chats and new chats support these options.
A send still requires `--yes` or `confirm=true`.

## CLI examples

```bash
# Preview a file and an inline image in your self-chat
outlook-cli teams send '48:notes' 'Please review these files.' --attach ./report.pdf --image ./preview.png --json

# Apply the send after you check the preview
outlook-cli teams send '48:notes' 'Please review these files.' --attach ./report.pdf --image ./preview.png --yes --json

# Send only an image
outlook-cli teams send <chat-id> --image ./preview.png --yes --json

# Start a new chat with a file
outlook-cli teams new --to person@example.com 'Please review the report.' --attach ./report.pdf --yes --json

# Use a different existing private OneDrive folder
outlook-cli teams send <chat-id> 'Report attached.' --attach ./report.pdf --upload-folder 'Documents/Reports' --yes --json
```

`--attach` and `--image` are repeatable.
You can omit the message text when you supply media.
The commands also accept `--html` for message text.
The CLI returns JSON for media previews and results.
A preview exits with code 2 and makes no remote changes.

## File attachments

Teams stores file attachments in OneDrive rather than in the message body.
The command uploads each file to the existing `Microsoft Teams Chat Files` folder by default.
It does not create a folder.
Use `--upload-folder` to select another existing private OneDrive folder.
The command rejects a folder that already has shared access.

The command adds a unique prefix to each remote filename.
It does not replace an existing file deliberately.
It uses the uploaded file's SharePoint GUID and WebDAV URL to create a native Teams reference attachment.
The attachment card retains the original local filename.

For an existing chat, the command checks the current membership before it grants file access.
It grants read access to each verified internal user other than the sender.
It uses directory object IDs rather than user-supplied email addresses.
It does not create an organisation-wide or anonymous link.
It does not send a separate invitation email.
The sender already owns files in the sender's self-chat, so that case needs no access grant.
Existing owner and administrator access remains unchanged.
A private folder does not exclude access that the tenant already grants to service accounts.

File attachments currently support internal 1:1 chats, internal group chats and the self-chat `48:notes`.
The command rejects meeting chats, guests, external members and members whose identity or tenant cannot be verified.
The client refuses a file send when it cannot verify all members within its 1000-member limit.
Permissions cover the verified members at the time of the send.
A later chat member does not automatically receive file access from this command.

## Inline images

Inline images use the message's Graph `hostedContents` collection.
The command sends the image bytes with the message.
It does not create a separate OneDrive file for an inline image.
The HTML image reference matches its temporary hosted-content ID.

The client accepts PNG, JPEG and GIF files with a matching extension and file signature.
It does not accept SVG, WebP, remote image URLs or arbitrary data URLs through `--image`.
The live check uses PNG.
JPEG and GIF share the same protocol, but their live acceptance remains unverified.

## Validation and limits

The command validates and reads all local media before its first remote write.
It rejects symlinks, directories, pipes, devices and filenames that OneDrive cannot accept.
The validated bytes remain fixed throughout one operation.
The command also checks the final escaped HTML size before it uploads a file or creates a chat.

The initial client limits are:

- 25 MiB for each file attachment.
- 1 MiB for each inline image.
- 2 MiB for all inline images together, before Base64 encoding.
- 50 MiB for all local media together.
- 10 files and inline images together.
- 20,000 UTF-8 bytes for message text.
- 28,000 UTF-8 bytes for the final HTML body, including escaped text and media references.

These are client limits, not claims about the full Teams service limits.
The image total leaves space under Graph's documented 4 MB hosted-content limit after Base64 encoding.
The tenant can impose other limits or block a file type.
The command reports those errors rather than silently dropping media.

## Preview and safety

A file preview reads the destination folder and chat membership.
It reports local names, sizes, the destination folder and the user IDs that will receive access.
It does not upload files, grant access, create a chat or send a message.
A preview for a new chat resolves the proposed recipients without creating the chat.
An image-only preview needs no remote write or file upload.

Caution: Confirmation permits file uploads and access grants as well as the message send.
Check the files, chat and access list before you confirm.
The command checks read-only and no-send policy before the first upload.
It checks the policy again before each access grant and final send.
A global `--dry-run` prints a plan without client creation, local file validation or remote checks.

Uploads, access grants and message sends have no automatic retries.
A failed operation reports its phase, chat ID and any remote file details that it knows.
A successful upload can remain in OneDrive when a later step fails.
The command never silently deletes such files.
A new chat can also remain after its first message fails.

Caution: A timeout or network error can leave an uncertain result.
Check the chat and reported remote filenames before any retry.
Do not repeat a new group chat command automatically.
The command does not roll back a successful access grant.

## MCP

`teams_send` and `teams_new` accept these additional arguments:

```json
{
  "attachments": ["/absolute/path/report.pdf"],
  "images": ["/absolute/path/preview.png"],
  "upload_folder": "Microsoft Teams Chat Files",
  "confirm": false
}
```

`content` defaults to an empty string so a message can contain only media.
The existing text-only interfaces remain supported.
A file send requires OneDrive write permissions in addition to the permissions for Teams chats.
A send to other members also requires permission to grant those members access to the file.
The token must permit the applicable Graph operations.
The client does not switch to a broader link when a permission grant fails.

Restart the MCP server after you update its code.
The Teams gateway keeps its separate file export workflow.
These options do not change the gateway.

## Live verification, 11 November 2026

Two confirmed MCP calls sent a PNG and a text file to Ross's self-chat, `48:notes`.
Separate read requests found both saved messages.
The hosted image bytes matched the local PNG exactly.
The OneDrive file bytes matched the local text file exactly.
The saved file message contained a native reference attachment.
Ross confirmed that Teams displayed the blue and green image and the text attachment.

The test created no new access grants or share links.
The file retained the same two owner records as the existing destination folder.
Those records cover Ross and WPP's existing migration service account.
The test did not change either owner record.

The test did not send to another person or create a new chat.
Member access grants, new-chat media and JPEG/GIF acceptance have synthetic tests but no live acceptance check.
The local evidence remains at `/tmp/outlook-teams-live-01a1cb0f.json`.

## Microsoft references

- [Send a chat message, including files and hosted images](https://learn.microsoft.com/en-us/graph/api/chatmessage-post?view=graph-rest-1.0)
- [Chat message attachment](https://learn.microsoft.com/en-us/graph/api/resources/chatmessageattachment?view=graph-rest-1.0)
- [Hosted content](https://learn.microsoft.com/en-us/graph/api/resources/chatmessagehostedcontent?view=graph-rest-1.0)
- [Upload file contents](https://learn.microsoft.com/en-us/graph/api/driveitem-put-content?view=graph-rest-1.0)
- [Grant access to a drive item](https://learn.microsoft.com/en-us/graph/api/driveitem-invite?view=graph-rest-1.0)
- [Drive recipient](https://learn.microsoft.com/en-us/graph/api/resources/driverecipient?view=graph-rest-1.0)
