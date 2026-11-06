# Plan: Copilot in Excel, PowerPoint and Word through a browser extension

Status: experimental implementation added (2026-10-22). Live tab discovery passed on 2026-10-23. Excel read-only and single-cell edit checks passed. PowerPoint read-only, title-edit and slide-insertion checks passed after manual pane opening. Saved files were independently checked. Unit and local pipe/socket integration tests pass. Do not treat the spike results below as release verification.

## Implementation status

- Version 0.4.0 (2026-11-06) fixes the problems in "Incident 2026-11-06" below. Synthetic tests pass. Live checks are still to do; see the release checklist.

- Bridge, installer, tab status, Excel and PowerPoint adapters, CLI and MCP entry points: implemented. CLI live reads and edits passed for Excel and PowerPoint. PowerPoint automatic pane opening passed with extension 0.2.0 after a refreshed tab was left closed. Earlier opening attempts failed, so cold-start reliability still needs more checks. MCP has fake-client tests, not live acceptance.
- Version 0.3.1 corrects drawer readiness: a mounted pane frame can report a locally visible input even when its parent drawer is closed. Ross confirmed this after the first automated Word send failed. Parent-side iframe geometry, CSS, clipping and hidden ancestor checks now gate pane selection and pre-send readiness. A complete visible embedding chain to the top frame is required. Synthetic tests pass; live validation remains pending. Controller version 3 is reinjected on extension reload; the Word hook remains revision 3.
- `--open`, `--new-chat`, `--wait-saved`: implemented, needs live checks. Save evidence is a later Graph modification timestamp. Other writers can change it.
- Approvals: the controller has a fixed click allowlist. It never clicks approval controls. Version 0.2.0 adds pre-send checks and a 250 ms observer for visible permission dialogs in editor/pane frames. This returns `approval_required` without clicking. Synthetic tests pass; live UI checks are still needed. Inline cards and backend annotation formats remain unverified. Excel pending-action detection still uses an unverified annotation shape. An unrecognised approval can time out.
- Word: experimental support added in extension 0.3.0 after Ross sent a manual read-only prompt in `dummy_document.docx`. The actual chat turn uses `OfficeAgentSignal`, `OfficeAgentAnnotation`, `OfficeAgentAssistantMessage`, progress and function-invocation schemas, shared with Excel. The earlier `WordCopilotOutputAnnotation` messages were observed before this turn and are not used for its parser. The shared adapter tags actions as Word and exports no function arguments. Word returns `applied: null`. Extension 0.3.2 passed automated read and edit checks on 2026-10-23. Hook revision 3 requires a refreshed Word tab after loading extension 0.3.0.
- Version 0.2.0 adds `copilot app inspect --tab-id ID`. It exports an allowlisted set of control labels/states and AugLoop schema/property names with counts. It never exports document text, prompt values, tokens, scripts or raw frames. A missing hook revision causes `ask` to require an Office tab refresh. This diagnostic build needs a Chrome extension reload before further live investigation.
- The native launcher runs `python -m outlook_draft.copilot_bridge`, not the general CLI dispatcher. This keeps CLI output off Chrome's stdout protocol channel.
- The worker also uses `alarms` for reconnects. A missing host is retried once per minute.
- Timed-out, interrupted and approval-pending tabs remain blocked until the user checks the document and reloads the tab. Cancellation stops collection, not the Office agent.
- Excel and Word return `applied: null`. Script source patterns are exported as `mayWrite`, not proof of a write. Generated helper functions caused a false positive during a read-only live check. PowerPoint uses explicit slide activity and client action responses. Independent saved-deck checks confirmed the live title and insertion tests.

### Release checklist (new implementation)

- [x] Load the installed extension in Chrome and list real Office tabs with `app status` (2026-10-23).
- [x] Excel single-cell edit in the disposable workbook; independently verify the saved E1 marker, all 12 existing cells and unchanged workbook styles (2026-10-23).
- [ ] Excel new-table generation check.
- [x] PowerPoint title edit and title-and-content slide insertion. Independently check saved OOXML and inspect rendered slides (2026-10-23). Manual pane opening was required.
- [x] Follow-up edit in the same Excel conversation after a read-only prompt (2026-10-23). A later read prompt was blocked by an unsent draft, which was not cleared.
- [ ] Record approvals in each app, implement missing detection, verify no approval is clicked.
- [ ] Reload the extension while Office tabs are open.
- [x] `--open` finds/opens disposable test files; `--wait-saved` returns timestamp evidence. PowerPoint automatic pane opening passed once after a refreshed tab was left closed (extension 0.2.0); cold-start reliability remains under test.
- [x] Background PowerPoint tab: automatic pane opening and read-only answer passed (2026-10-23).
- [ ] Test `--new-chat`.
- [ ] Test Chrome closed and disconnected host errors.
- [x] Word pane/protocol spike and adapter. Extension 0.3.2: read-only prompts with the drawer open and closed, and a single-paragraph edit with saved-file checks (2026-10-23).
- [ ] 0.4.0: a multi-line PowerPoint prompt with bullets returns by itself, with `completion: event` and `binding` set. Record `statuses` and `signals`.
- [ ] 0.4.0: `--open` on a closed disposable deck opens the pane by itself. Record the phases.
- [ ] 0.4.0: `--new-chat` on a fresh pane (skipped) and on a pane with a conversation (clicked).
- [ ] 0.4.0: a short `--timeout` ends with `status: timeout` and a phase; `app status` shows `uncertain`; `app reset` clears it.
- [ ] 0.4.0: find the label of the PowerPoint Stop control with `inspect` while Copilot works. The settle fallback needs it, or a finished task status.

Live checks (2026-10-23): after readiness and exact-prompt correlation changes, Excel returned the correct sample-table description. A follow-up set only Sheet1!E1 to `Bridge live check 2026-10-23`. The saved workbook was downloaded and checked: the marker is present, all 12 original cells are unchanged, and `xl/styles.xml` is unchanged. Graph reported a newer modification time before downloads contained the edit, so timestamp evidence is not sufficient. A later read was refused because the prompt box contained a draft; it was not cleared. Evidence files are in `/tmp/copilot-live-checks/`.

PowerPoint initially opened the disposable `Presentation 4.pptx`, but its Copilot pane did not become ready within 30 seconds. After Ross opened the pane manually, the CLI read the two slide titles, changed only slide 1's title to `Copilot bridge live check`, and appended slide 3, `Bridge verification`, with Read/Edit/Saved bullets in the same conversation. The saved title edit matched the original slide XML with only its title replaced; slide 2 was byte-identical. Office also refreshed date fields in the master/layouts from 2026-10-22 to 2026-10-23. After insertion, both existing slide XML files were byte-identical to their prior versions. Slide 3 uses the existing `slideLayout2.xml`. All three final slides were rendered through LibreOffice and visually checked. No Office lock files were present. Evidence is in `/tmp/copilot-live-checks/ppt-*`.

Extension 0.2.0 diagnostics confirmed the closed PowerPoint pane was mounted but lacked a visible prompt. A subsequent CLI request opened it automatically and returned slide 1's correct title, with no action events. Word's dummy document has a ready prompt and an **Allow editing** menu; its protocol differs from Excel. No Word prompt was sent by the CLI.

Word checks (2026-10-23, extension 0.3.2, controller v5): with the drawer open, the CLI sent a read-only prompt to `dummy_document.docx`. Copilot returned a one-sentence summary and called only `get_active_document_html`. With the drawer closed, the bridge opened the drawer and the same prompt completed. A third prompt added one paragraph at the end. The actions were `get_active_document_html`, `begin_edit_session`, `run_js_code` and `end_edit_session`. No approval came up. The saved `.docx` was downloaded and compared with a copy from before the edit. The new last paragraph is `Bridge live check 2026-10-23.` (Copilot added the full stop). All 17 original paragraphs keep their text and order. Word Online wrote the bold run of the previous last paragraph again as `<w:b w:val="1"/><w:bCs w:val="1"/>`, with the same meaning. Only `word/settings.xml` (one new rsid) and `docProps/core.xml` changed in the other parts. Evidence is in `/tmp/copilot-live-checks/word-v5-*` and `/tmp/copilot-live-checks/word-edit/`.

Chrome reload initially raced the previous native host's socket lock. A bounded lock wait and regression test now cover that case.

Local checks: `.venv/bin/python -m pytest -q` (315 passed, 1 skipped) and `node --test extension/tests/*.test.js` (32 passed). The JS fixtures are synthetic and contain no real document or tenant data.

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

## Incident 2026-11-06 (extension 0.3.2)

Target: a shared SharePoint deck. The prompt had several paragraphs and a bulleted list. It asked Copilot to edit slide 6 only.

1. `--open --new-chat`, tab not open: the tab opened, then "The Copilot pane did not open in time." The bridge clicked the Copilot button as soon as it appeared. The editor was probably not ready to handle the click.
2. `--new-chat`, pane open from run 1: "The pane has no New chat button." A fresh pane has no New chat button. The button appears only after a conversation starts.
3. No flags: Copilot edited only slide 6, and the file was saved. The command did not return. Ross stopped it after about six minutes (`--timeout 400`).

Cause of run 3: the hook bound a request only when the signal `query` was exactly the typed prompt. The pane is a rich text editor, so typed newlines, list markers and spaces come back changed. The hook never bound, the adapters got no records, and the worker had only its timer. The 0.3.2 timer started after the pane opened. After the interrupt, the cancel marked the tab uncertain, and `status` showed that as `busy: True`.

Signal count: `inspect` showed two `PptAgentExperimentalSignal` sends. These counts cover the life of the page, not one request. Runs 1 and 2 stopped before arm and before `send_prompt`, which is the only call that types. So the bridge did not send twice. The second signal is either a pane start signal or a task operation (the signal has `allowResume` and `taskOperationParams`). Version 0.4.0 reports per-request signal classes to settle this in the next live run.

Observed PowerPoint schemas (names and keys only): `PptAgentExperimentalOutputAnnotation` (`finalResponse`, `responseStatus`, `queryId`, `taskId`, `sydneyConversationId`), `PptAgentExperimentalOutputAnnotationV1` (`pptEditsPayload`, `responseStatus`, `queryId`), `Core_TaskInformationAnnotation` (`taskStatus`, `ownerId`). Five V1 annotations matched five `CopilotActionResponseObject` sends, so V1 carries the edit batches.

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
- **Completion (0.4.0):** a `completed` event from the adapter ends a request. Without it, the worker polls the pane every 2 s. It settles the turn when Copilot was seen working (a Stop control, or a finished task status), the pane then shows no Stop control for 3 polls, and the hook has been quiet for 10 s. A settled result has `completion: settled`. If the hook saw nothing, `applied` is null.
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
- Return the answer, the actions and any pending approval. Do not infer `applied: true` from Excel script source. The current implementation reports `applied: null` for Excel.
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
