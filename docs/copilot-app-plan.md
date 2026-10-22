# Plan: Copilot in Excel, PowerPoint and Word through a browser extension

Status: not built. This plan is kept for reference (2026-10-22).

## Context

`outlook-cli copilot ask` uses the main Microsoft 365 Copilot chat. That chat can read files, but it cannot edit an open document. Its document tools only make a new copy.

The Copilot agent inside each Office app can edit the document. Ross wants the CLI and the MCP server to send a prompt to that agent for a given file and app, and to get the answer back.

The in-app agents need the real Office web editor, because:

- The editor sends its document model to the agent over AugLoop.
- The agent sends edits back, and the editor JavaScript in the browser applies them. In Excel, the agent sends Office.js code (`executeScript`). In PowerPoint, it sends `CopilotAction` operations.
- The AugLoop token comes from the editor session.

So the CLI controls an Office tab in Ross's own Chrome, through an extension and a native messaging host. `suimi8/augloop-copilot-proxy` connects to AugLoop without an editor. It reads tokens from the memory of the Office process, fakes the client and cannot edit. Do not use that method.

## Spike results (2026-10-22)

A spike extension in the scratchpad controlled real tabs for `Book 1.xlsx` and `Presentation 4.pptx`.

| | Excel | PowerPoint |
|---|---|---|
| Open the pane | Button `aria-label="Chat with Copilot"` in the editor frame | Same |
| Editor frame | `*.officeapps.live.com/x/_layouts/xlviewerinternal.aspx` | `*.officeapps.live.com/pods/ppt.aspx` |
| Pane frame | `*.resources.office.net/.../taskpane.html` ("Excel Copilot Chat Add-in") | Same host ("Copilot Chat") |
| Prompt box | `contenteditable`, `aria-label="Chat input"` | `contenteditable`, `aria-label="Type your message"` |
| Type the prompt | `document.execCommand("insertText")` | Same |
| Send | Button `aria-label="Send"`. It appears after the box has text. | Same |
| AugLoop endpoint | `wss://augloop.office.com` | `wss://msit.augloop.svc.cloud.microsoft` |
| Agent protocol | `AugLoop_OfficeAgent_*` | `AugLoop_PptAgentExperimental_*` |
| Answer text | `OfficeAgentAssistantMessage.text` (pieces, then the full text) | `PptAgentExperimentalOutputAnnotation` (`chunkContent`, then `finalResponse.assistantResponse`) |
| End of turn | `status: "completed"` | `responseStatus: "complete"` |
| Edit | `OfficeAgentInvokeFunctionAnnotation`, `namespace: "excel"`, `functionName: "executeScript"` | `CopilotActionResponseObject` from the client |
| Saved in OneDrive | Yes, some minutes after the edit | Yes |
| Follow-up prompt | Not tested | Works in the same pane |

Lessons:

1. **Never change the editor DOM.** The spike wrote a `data-*` attribute on PowerPoint's editable slide nodes, and PowerPoint failed with "Sorry, we ran into a problem". With ids kept in a JavaScript map, the error did not come back.
2. **A reloaded extension loses its scripts in open tabs.** The service worker must inject its scripts again into open Office tabs when it starts.
3. **The edit shows at once, but OneDrive gets it later.** For Excel, it took some minutes.
4. **A main-world WebSocket hook sees the AugLoop traffic** in the editor frame, in both apps.

## Goals

- `outlook-cli copilot app ask "<prompt>" --file <url|id|path>` sends the prompt to the in-app agent and streams the answer.
- The MCP tool `copilot_app` does the same for agents.
- Excel and PowerPoint first. Word after a short spike.
- The tool uses only Ross's own browser session. It does not capture or store tokens.

## Non-goals

- No direct AugLoop client and no token capture for AugLoop.
- No automatic approval of agent actions or permission requests.
- No support for browsers other than Chrome in the first version.

## Architecture

```
outlook-cli / MCP ──Unix socket──> native host (Python) ──stdio──> extension service worker
                                   started by Chrome                     │
                                                          content scripts in the Office tab
                                                          ├─ hook.js (main world, editor frame):
                                                          │    reads AugLoop frames, app adapter
                                                          │    turns them into events
                                                          └─ controller.js (isolated world):
                                                               opens the pane, types, sends
```

### 1. Extension (`extension/`)

Manifest V3, loaded unpacked. A `key` in the manifest gives a fixed extension ID.

- **Permissions:** `nativeMessaging`, `tabs`, `scripting`. Host permissions only for `https://*.sharepoint.com/*`, `https://*.officeapps.live.com/*`, `https://*.resources.office.net/*` and `https://*.cloud.microsoft/*`. There is no `<all_urls>` and no localhost access.
- **`sw.js` (service worker):**
  - Calls `chrome.runtime.connectNative("com.outlook_cli.copilot")` when it starts. The open port keeps the worker running.
  - Injects the content scripts again into open Office tabs when it starts.
  - Finds the tab for a file, or opens one if the request allows it.
  - Lets only one request run in a tab at a time.
  - Routes commands to frames with `chrome.tabs.sendMessage(tabId, msg, {frameId})`.
  - Relays events to the native host.
- **`hook.js` (main world, editor frames only):**
  - Wraps `WebSocket` and watches only AugLoop URLs.
  - Passes each frame to the adapter of the app. It sends only the events that the adapter makes, never raw frames. PowerPoint sends the whole document model over AugLoop, and that content must not leave the page.
- **`adapters/excel.js`, `adapters/powerpoint.js`, later `adapters/word.js`:** each one is a pure function from an AugLoop message to zero or more events:
  - `started` (request ID, conversation ID)
  - `progress` (header text)
  - `text` (the new piece of the answer)
  - `action` (a short summary: Excel script length and target ranges, or PowerPoint action names)
  - `approval_required` (the type and text of the request)
  - `completed` (the full text)
  - `error`
- **`controller.js` (isolated world, all Office frames):** it has these commands:
  - `status`: app, editor ready, pane open, edit mode.
  - `open_pane`
  - `new_chat`
  - `send_prompt`
  - `read_pane`: a fallback answer from the pane text.

  It finds controls by `aria-label` and role. It keeps element references in a `Map` and never writes to the DOM.

### 2. Native host (`outlook_draft/copilot_bridge/`)

- **`native_host.py`:** Chrome starts it through a launcher script (`outlook-cli copilot native-host`).
  - Native messaging framing on stdin and stdout: a 4-byte length in native byte order, then UTF-8 JSON. Nothing else may write to stdout. Logs go to `session_state/copilot-bridge.log`.
  - It listens on `~/.local/share/outlook-cli/copilot-bridge.sock`. The folder has mode 0700 and the socket has mode 0600. It accepts only peers with the same UID (`SO_PEERCRED`).
  - It forwards requests and streams events back to the socket client that sent the request.
- **`client.py`:** the socket client for the CLI and the MCP server. It has `status()` and `ask(...)`, which returns an iterator of events and a final result.
- **Protocol:** newline-delimited JSON, with `{"v": 1, ...}` in each message. When it connects, the client checks the protocol version of the extension and of the host.

### 3. CLI (`outlook_draft/commands/copilot.py`)

```
outlook-cli copilot app ask "<prompt>" --file <url|id|path> [--site S] [--library L]
                         [--app excel|powerpoint|word] [--new-chat] [--open]
                         [--wait-saved] [--timeout 300] [--json]
outlook-cli copilot app status            # host, extension, open Office tabs
outlook-cli copilot extension install     # native host manifest and launcher
outlook-cli copilot native-host           # started by Chrome only (hidden)
```

- Resolve `--file` with `resolve_copilot_file()`. This gives the drive item, `webUrl` and the `sourcedoc` GUID that identifies the tab.
- Get the app from the file extension: `.xlsx` and `.xlsm` are Excel, `.pptx` is PowerPoint and `.docx` is Word. `--app` overrides this.
- `--open` lets the extension open the file in a new tab when no tab has it. Without `--open`, the command fails with "Open the file in Chrome first."
- Stream `text` events to stdout. Print `action` and `approval_required` events to stderr as they happen.
- Return the result: the answer, the actions, any pending approval, and `applied: true`.
- `--wait-saved` polls the Graph `lastModifiedDateTime` of the file until it is later than the start time, or until the time limit. The result then reports `saved: true` or `saved: false`.
- Exit codes: 0 for completed, 3 for approval required, and 1 for errors.
- `extension install` writes `~/.config/google-chrome/NativeMessagingHosts/com.outlook_cli.copilot.json`. The file has `allowed_origins` set to the fixed extension ID and a launcher in `~/.local/share/outlook-cli/`. The command prints the folder to load with **Load unpacked**.

### 4. MCP tool (`outlook_draft/mcp/copilot.py`)

`copilot_app(prompt, file, app=None, site=None, library=None, new_chat=False, open_tab=False, wait_saved=False)`:

- Register it with `readOnlyHint=false`, `destructiveHint=true` and `idempotentHint=false`.
- Its description says that the tool edits the live document, that other people see the edits, and that SharePoint version history can restore the file.
- It returns the same result as `--json`.

## Safety rules

1. Never approve an agent action, a permission request or a plan proposal. Stop, and return `approval_required`.
2. Never click **Done**, **Undo**, **Keep** or the **Allow editing** control. Report the **Allow editing** setting in `status`.
3. Never change the editor DOM.
4. Never send document content, tokens or raw AugLoop frames out of the page. Send only adapter events.
5. The socket accepts only the same user, and the extension accepts messages only from its native host.
6. One request runs in a tab at a time. A second request gets the error "busy".
7. If the file is not in Ross's own OneDrive, the CLI prints a caution, because other people see the edits at once.

## Implementation phases

1. **Bridge:** extension skeleton, native host, socket, `extension install`, `app status`. Done when `status` lists the open Office tabs and their apps.
2. **Excel:** controller, Excel adapter, `app ask` with streaming. Done when the Book 1 test passes from the CLI, and the result reports the `executeScript` action.
3. **PowerPoint:** adapter. Done when the title test and the agenda slide test pass, and no error dialog appears.
4. **Approvals:** set **Allow editing** to ask in each app, record the approval messages, and add `approval_required` to the adapters. Done when the CLI stops and reports the request, and the document does not change.
5. **MCP and docs:** the `copilot_app` tool, README, `AGENTS.md`, `docs/copilot-tools.md`.
6. **Word:** spike the pane and the protocol, then add an adapter.
7. **Additions:** `--open`, `--new-chat`, `--wait-saved`, and behaviour in background tabs.

## Tests

- **Python (pytest):**
  - native message framing
  - the socket protocol and the version check
  - the same-UID check
  - CLI app detection and errors
  - `extension install` output, written to a temporary folder
  - the MCP tool with a fake bridge client
  - the exit codes
- **JavaScript (`node --test`, no build step):** the adapters as pure functions, with synthetic fixtures based on the spike frames. The fixtures have no tokens, tenant IDs or document content. Tests for controller selectors use a small fake DOM.
- **Manual checks** (a checklist in this file for each release):
  - Excel table prompt
  - PowerPoint title prompt and slide prompt
  - a follow-up prompt
  - the approval case
  - a reload of the extension while tabs are open
  - Chrome closed, which gives a clear error

## Open questions

1. What does an approval request look like when **Allow editing** asks first? (Phase 4)
2. Does the agent work when its tab is in the background, or does Chrome slow its timers too much? Do we need to bring the tab to the front?
3. Does Word use `OfficeAgent` like Excel, or its own agent?
4. Should the tool refuse files that are not Ross's by default, and need a flag such as `--allow-shared`?
5. The pane keeps its own conversation. Is that enough, or do we need `--new-chat` for each call by default?
6. Do we need Edge support? The native host manifest folder and the extension ID are different there.

## Risks

- **The UI or the protocol can change.** The selectors and the AugLoop message names are not public. Keep them in the adapters, and keep the manual checklist current.
- **Policy.** Ross's own browser runs the agent, so this is closer to normal use than a token proxy. But it is still automation of Copilot that Microsoft does not support, and it edits live files.
- **Edits to shared files.** Co-authors see each change at once. Version history is the only way back.

## Clean-up of the spike

- Stop the spike collector on `127.0.0.1:8765`.
- Remove the spike extension from `chrome://extensions`. It has access to all sites.
- Delete the scratchpad folders `copilot-ext/` and `augloop-proxy-review/`.
- Delete the test files `Book 1.xlsx` and `Presentation 4.pptx`, or keep them for phase 2 and 3 tests.
