# Teams chat activity

Chat `lastUpdatedDateTime` is not a last-message timestamp. New messages can leave this field unchanged.

CLI and MCP lists now request `lastMessagePreview` and `members` through Graph `$expand`. Graph orders the chat collection by `lastMessagePreview/createdDateTime desc` before paging. Both interfaces use that timestamp for local sorting. A missing preview triggers a bounded concurrent read of one message. Empty or inaccessible chats have no message timestamp. They sort last within the retrieved set. A failed fallback produces a warning.

The default is latest activity, including messages sent by the current user. CLI `--sort-received` remains an explicit, slower option. It can require message reads for each returned chat.

## CLI

```sh
outlook-cli teams list
outlook-cli teams search Vagelis
outlook-cli teams search Vagelis --one-on-one
outlook-cli teams with Vagelis
```

Tables show the last message time, sender and a one-line preview. Full chat IDs appear below the table. Use a full ID with `teams messages`, `teams show` or `teams send` to avoid changing numeric indexes. Sending still requires explicit confirmation.

`teams with` searches only 1:1 chats. It also reads recent messages when there is one matching result. If there are several matches, use the full ID for the correct person.

## MCP

Existing `teams_list` and `teams_search` return types remain unchanged. Each chat has these additional fields:

- `lastMessageTime`: the actual message timestamp, or an empty string.
- `lastSender`: the sender label, when available.
- `preview`: plain text on one line. System events and empty content have explicit labels.

`lastUpdated` retains its original meaning for compatibility.

`teams_search` accepts optional `one_on_one` and `include_messages` booleans. Both default to false. Set both to true to find a person's 1:1 chat and read recent messages. Each returned result includes `searchIncomplete`. A true value means Graph has chats beyond the scan limit, or at least one matching result was omitted by the result limit. An exact count alone does not trigger a warning. The server also logs a warning, including when no results are returned.

The client preserves Graph continuation metadata in a list-compatible chat collection. Search checks up to one extra match within the scan to detect omitted results.

Searches remain bounded. Increase `--scan`/`scan` or `--count`/`count` when a limit warning appears. MCP scan and result limits remain 500 and 50. Member expansion is reused only when it has no continuation link. Otherwise the existing paged member lookup is used.

No date, sender or unread filters were added in this fix.
