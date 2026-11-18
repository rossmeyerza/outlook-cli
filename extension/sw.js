// Service worker: connects to the native host, finds the Office tab for a file,
// drives the controller in its frames, and relays adapter events to the host.
"use strict";

const PROTOCOL = 1;
const HOST_NAME = "com.outlook_cli.copilot";
const VERSION = chrome.runtime.getManifest().version;
const OFFICE_URLS = [
  "https://*.sharepoint.com/*",
  "https://*.officeapps.live.com/*",
  "https://*.cloud.microsoft/*",
];
const HOOK_FILES = ["adapters/common.js", "adapters/excel.js", "adapters/powerpoint.js", "hook.js"];
const HOOK_REVISION = 6;
// Fallback end of a turn when no final answer event arrives: the pane shows no
// Stop control for several polls, and the hook has been quiet for a while.
const TIMING = {
  settlePollMs: 2000, settleIdlePolls: 3, settleQuietMs: 10000, staleGraceMs: 30000, minTimeoutMs: 10000,
  editorPollMs: 1000, panePollMs: 500,
  // The ribbon button can render before the editor handles a click. Wait for
  // the AugLoop socket, or this long after the button appears.
  editorSettleMs: 10000, beforeOpenMs: 1000,
  paneFirstWaitMs: 12000, paneWaitMs: 30000, inputSettleMs: 1500,
  // A frozen or busy tab can leave executeScript waiting for ever. Each frame
  // call has a limit, so status and the request waits keep checking time.
  frameCallMs: 10000, sendCallMs: 30000,
};
const TASK_DONE = /^(complete|completed|success|succeeded|done|finished)$/i;

let port = null;
let lastHostError = null;
const requests = new Map(); // tab id -> active request
// tab id -> { since, reason }. A request ended after Send without a final
// answer. A timeout does not stop the Office agent, so the document can still
// change. A new request waits until the user checks the document and resets.
const uncertainTabs = new Map();
const cancelledRequests = new Set();

const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function send(msg) {
  if (port) port.postMessage({ v: PROTOCOL, ...msg });
}

function connect() {
  if (port) return;
  try {
    port = chrome.runtime.connectNative(HOST_NAME);
  } catch (e) {
    lastHostError = String(e.message || e);
    port = null;
    return;
  }
  const connectedPort = port;
  connectedPort.onMessage.addListener((msg) => onHostMessage(msg).catch((e) => console.warn("ocli", e)));
  connectedPort.onDisconnect.addListener(() => {
    // Chrome requires lastError to be read inside this callback. Otherwise an
    // expected reload disconnect appears as an unchecked runtime error.
    const error = chrome.runtime.lastError;
    if (port !== connectedPort) return;
    lastHostError = error ? error.message : "The native host disconnected.";
    port = null;
    for (const req of requests.values()) finish(req, "error", { message: lastHostError });
    // Reconnect the transport only. Never replay an interrupted Office request.
    chrome.alarms.create("ocli-reconnect", { delayInMinutes: 0.05 });
  });
  send({ type: "hello", extension: VERSION, protocol: PROTOCOL });
}

// ---- frames -------------------------------------------------------------

async function inject(tabId) {
  const target = { tabId, allFrames: true };
  await chrome.scripting.executeScript({ target, files: HOOK_FILES, world: "MAIN" }).catch(() => {});
  await chrome.scripting.executeScript({ target, files: ["diagnostics.js", "controller.js"] }).catch(() => {});
}

async function reinjectAll() {
  const tabs = await chrome.tabs.query({ url: OFFICE_URLS });
  await Promise.all(tabs.map((tab) => inject(tab.id)));
}

// The tab did not answer a frame call in time. It can be frozen, or busy with a large file.
class FrameTimeout extends Error {}

async function runInFrames(tabId, cmd, frameIds, timeoutMs = TIMING.frameCallMs) {
  const target = frameIds ? { tabId, frameIds } : { tabId, allFrames: true };
  let timer;
  const limit = new Promise((resolve, reject) => {
    timer = setTimeout(() => reject(new FrameTimeout(
      `The tab did not answer ${cmd.action} within ${Math.round(timeoutMs / 1000)} s. It can be frozen or busy.`)), timeoutMs);
  });
  let results;
  try {
    results = await Promise.race([chrome.scripting.executeScript({
      target,
      func: (c) => (globalThis.__ocliController ? globalThis.__ocliController.run(c) : null),
      args: [cmd],
    }), limit]);
  } finally {
    clearTimeout(timer);
  }
  return results.filter((r) => r.result).map((r) => ({ frameId: r.frameId, ...r.result }));
}

// The Copilot drawer is open when the frame that embeds the pane shows it.
// Word keeps the pane iframe mounted with a ready input while the drawer is
// closed, so the pane cannot judge this itself. Only the direct parent is
// checked: the SharePoint host page reports the Word editor iframe as hidden
// even while it is on screen, so a full chain to the top frame fails.
function framePresented(frame, frames) {
  if (frame.isTop) return true;
  return frames.some((parent) => parent.frameId !== frame.frameId && (parent.embeddedFrames || [])
    .some((entry) => entry.url === frame.url && entry.visible === true));
}

async function probe(tabId) {
  let unresponsive = false;
  const frames = await runInFrames(tabId, { action: "describe" }).catch((e) => {
    unresponsive = e instanceof FrameTimeout;
    return [];
  });
  const editors = frames.filter((f) => f.role === "editor");
  const panes = frames.filter((f) => f.role === "pane" && framePresented(f, frames));
  const mounted = frames.filter((f) => f.role === "pane").length;
  return { frames, editors, panesMounted: mounted, unresponsive, pane: panes.find((p) => p.promptBox) || panes[0] || null };
}

async function runIn(tabId, frameId, cmd, timeoutMs) {
  const [result] = await runInFrames(tabId, cmd, [frameId], timeoutMs);
  if (!result) throw new Error(`No controller answered ${cmd.action}.`);
  return result;
}

// ---- tabs ---------------------------------------------------------------

function tabMatches(tab, file) {
  const url = (tab.url || "").toLowerCase();
  let decoded = url;
  try { decoded = decodeURIComponent(url); } catch (e) { /* keep the raw URL */ }
  const id = (file.uniqueId || "").toLowerCase().replace(/[{}]/g, "");
  if (id && decoded.includes(id)) return true;
  return Boolean(file.webUrl) && url.split("#")[0] === file.webUrl.toLowerCase();
}

function appOf(probeResult) {
  const editor = probeResult.editors[0];
  return editor ? editor.app : null;
}

// The action value of an Office URL (view, edit, default ...). Never the URL itself.
function urlAction(url) {
  let action = null;
  try { action = new URL(url).searchParams.get("action"); } catch (e) { return null; }
  if (!action) return null;
  return /^[a-z]{1,20}$/i.test(action) ? action.toLowerCase() : "other";
}

const ageSeconds = (since) => (since ? Math.round((Date.now() - since) / 1000) : null);

async function describeTab(tab) {
  const p = await probe(tab.id);
  const editor = p.editors[0] || null;
  const req = requests.get(tab.id);
  const uncertain = uncertainTabs.get(tab.id);
  return {
    tabId: tab.id,
    title: tab.title,
    active: tab.active,
    app: appOf(p),
    editorReady: Boolean(editor),
    unresponsive: p.unresponsive,
    urlAction: urlAction(tab.url),
    mode: editor ? editor.mode || null : null,
    hooked: Boolean(editor && editor.hook && editor.hook.seen),
    hookRevision: editor && editor.hook ? editor.hook.revision : null,
    controllerVersion: editor ? editor.version : null,
    augloopSockets: editor && editor.hook ? editor.hook.sockets : 0,
    paneOpen: Boolean(p.pane && p.pane.promptBox),
    allowEditing: p.pane ? p.pane.allowEditing : null,
    busy: Boolean(req),
    busyAgeSeconds: req ? ageSeconds(req.startedAt) : null,
    phase: req ? req.phase : null,
    uncertain: Boolean(uncertain),
    uncertainReason: uncertain ? uncertain.reason : null,
    uncertainAgeSeconds: uncertain ? ageSeconds(uncertain.since) : null,
  };
}

async function status() {
  const tabs = await chrome.tabs.query({ url: OFFICE_URLS });
  // In parallel: each probe has its own limit, so one frozen tab cannot block the others.
  const described = await Promise.all(tabs.map(describeTab));
  return { extension: VERSION, protocol: PROTOCOL, lastHostError,
    tabs: described.filter((info) => info.editorReady || info.unresponsive) };
}

async function officeTab(tabId) {
  const tab = await chrome.tabs.get(tabId);
  let url;
  try { url = new URL(tab.url); } catch (e) { throw new Error("The tab has no valid URL."); }
  if (url.protocol !== "https:" || !/(^|\.)(sharepoint\.com|officeapps\.live\.com|cloud\.microsoft)$/.test(url.hostname))
    throw new Error("Only Office browser tabs can be used.");
  return tab;
}

async function inspectTab(tabId) {
  await officeTab(tabId);
  return { extension: VERSION, tabId, frames: await runInFrames(tabId, { action: "inspect" }) };
}

// Diagnostic: click the Copilot button once and report what the pane did.
// It never types or sends. The button only shows or hides the pane.
async function openPaneOnce(msg) {
  const tab = await officeTab(msg.tabId);
  if (requests.get(tab.id)) throw new Error("busy: a request is running in this tab.");
  const before = await probe(tab.id);
  const report = { extension: VERSION, tabId: tab.id, active: Boolean(tab.active), clicked: false,
    paneOpenBefore: Boolean(before.pane && before.pane.promptBox), panesMountedBefore: before.panesMounted };
  if (report.paneOpenBefore) return { ...report, paneOpen: true, message: "The pane was already open. Nothing was clicked." };
  const editor = before.editors.find((e) => e.paneButton);
  if (!editor) throw new Error(before.unresponsive ? "The tab did not answer." : "The Copilot button is not on the page.");
  report.augloopSockets = editor.hook ? editor.hook.sockets : 0;
  const clicked = await runIn(tab.id, editor.frameId, { action: "open_pane" });
  if (!clicked.ok) throw new Error(clicked.error);
  report.clicked = true;
  const start = Date.now();
  const waitMs = Math.min(Math.max(Number(msg.waitMs) || 15000, 1000), 60000);
  let after = before;
  while (Date.now() - start < waitMs) {
    await sleep(TIMING.panePollMs);
    after = await probe(tab.id);
    if (after.pane && after.pane.promptBox) break;
  }
  const open = Boolean(after.pane && after.pane.promptBox);
  return { ...report, paneOpen: open, openedAfterMs: open ? Date.now() - start : null,
    panesMountedAfter: after.panesMounted, unresponsive: after.unresponsive };
}

// Several tabs can show the same file. Use the active one, then the one used
// last: that is the tab where the user is most likely to have the pane open.
async function findTab(file, open) {
  const tabs = (await chrome.tabs.query({})).filter((t) => tabMatches(t, file));
  tabs.sort((a, b) => (Number(Boolean(b.active)) - Number(Boolean(a.active))) || ((b.lastAccessed || 0) - (a.lastAccessed || 0)));
  if (tabs.length) return { tab: tabs[0], matches: tabs };
  if (!open) throw new Error("Open the file in Chrome first, or use --open.");
  if (!file.webUrl) throw new Error("The file has no web URL to open.");
  const created = await chrome.tabs.create({ url: file.webUrl, active: true });
  return { tab: created, matches: [created] };
}

// Thrown by a wait when its request already ended (timeout, cancel, disconnect).
class Stopped extends Error {}

function untilDone(req, timeoutMs) {
  return Math.min(Date.now() + timeoutMs, req ? req.deadline : Infinity);
}

// What the current wait saw last. A timeout message names it.
function noteWait(req, state) {
  if (req) req.waitState = state;
}

function checkActive(req) {
  if (req && req.done) throw new Stopped("The request already ended.");
}

// Ready means the pane is already open, or the Copilot button is on the page
// and the editor's AugLoop socket is connected (or the button has been there
// for a while). A new tab shows the button before the editor handles clicks.
async function waitForEditor(tabId, timeoutMs, req) {
  const end = untilDone(req, timeoutMs);
  let buttonSince = 0;
  let last = null;
  while (Date.now() < end) {
    checkActive(req);
    const p = await probe(tabId);
    last = p;
    noteWait(req, p.unresponsive ? "the tab did not answer" : !p.editors.length ? "no editor frame"
      : p.editors.some((e) => e.paneButton) ? "Copilot button seen" : "no Copilot button");
    if (p.editors.length && p.pane && p.pane.promptBox) return p;
    const editor = p.editors.find((e) => e.paneButton);
    if (editor) {
      buttonSince = buttonSince || Date.now();
      const connected = Boolean(editor.hook && editor.hook.seen && editor.hook.sockets > 0);
      if (connected || Date.now() - buttonSince >= TIMING.editorSettleMs) return p;
    } else {
      buttonSince = 0;
    }
    await sleep(TIMING.editorPollMs);
  }
  const seen = last && last.unresponsive ? "the tab did not answer; it can be frozen or busy with a large file"
    : !last || !last.editors.length ? "no editor frame" : buttonSince ? "Copilot button without an AugLoop connection" : "no Copilot button";
  throw new Error(`The Office editor did not load in time (${seen}).`);
}

async function waitForPane(tabId, timeoutMs, req) {
  const end = untilDone(req, timeoutMs);
  let mounted = 0;
  let unresponsive = false;
  while (Date.now() < end) {
    checkActive(req);
    const p = await probe(tabId);
    if (p.pane && p.pane.promptBox) return p;
    // A probe that timed out says nothing about the pane: keep the last count.
    if (!p.unresponsive) mounted = p.panesMounted || 0;
    unresponsive = p.unresponsive;
    noteWait(req, unresponsive ? "the tab did not answer" : mounted ? "pane frame loaded but not shown" : "pane frame not loaded");
    await sleep(TIMING.panePollMs);
  }
  const state = unresponsive ? "the tab did not answer" : mounted ? "pane frame loaded but not shown" : "pane frame not loaded";
  throw new Error(`The Copilot pane did not open in time (${state}).`);
}

// Click the Copilot button. If no pane frame loads at all, and the button
// does not report itself as open, click once more. The button only toggles
// the pane; it sends nothing.
async function openPane(req, tabId, p) {
  setPhase(req, "open_pane");
  await sleep(TIMING.beforeOpenMs);
  checkActive(req);
  const editor = p.editors.find((e) => e.paneButton) || p.editors[0];
  const opened = await runIn(tabId, editor.frameId, { action: "open_pane" });
  if (!opened.ok) throw new Error(opened.error);
  setPhase(req, "wait_pane");
  try {
    p = await waitForPane(tabId, TIMING.paneFirstWaitMs, req);
  } catch (e) {
    if (e instanceof Stopped) throw e;
    const now = await probe(tabId);
    const button = now.editors.find((x) => x.paneButton);
    if (!now.unresponsive && !(now.pane && now.pane.promptBox) && !now.panesMounted && button && button.paneButtonState !== "true") {
      setPhase(req, "open_pane_retry");
      const retry = await runIn(tabId, button.frameId, { action: "open_pane" });
      if (!retry.ok) throw new Error(retry.error);
      setPhase(req, "wait_pane");
    }
    p = await waitForPane(tabId, TIMING.paneWaitMs, req);
  }
  // The input can be on screen before it accepts typing.
  await sleep(TIMING.inputSettleMs);
  return waitForPane(tabId, 5000, req);
}

// ---- requests -----------------------------------------------------------

function finish(req, status, extra = {}) {
  if (req.done) return;
  req.done = true;
  clearTimeout(req.timer);
  clearTimeout(req.settleTimer);
  if (requests.get(req.tabId) === req) requests.delete(req.tabId);
  if (req.sent && status !== "completed") uncertainTabs.set(req.tabId, { since: Date.now(), reason: status });
  for (const frameId of req.armedFrames || []) {
    runIn(req.tabId, frameId, { action: "arm", on: false }).catch(() => {});
  }
  const result = {
    status,
    app: req.app,
    tabId: req.tabId,
    text: extra.text !== undefined ? extra.text : req.text,
    conversationId: extra.conversationId || req.conversationId || null,
    agentRequestId: req.agentRequestId || null,
    progress: req.progress,
    actions: req.actions,
    approval: extra.approval || null,
    message: extra.message || null,
    completion: extra.completion || null,
    phase: req.phase,
    sent: req.sent === true,
    newChat: req.newChat || null,
    signals: req.signals,
    binding: req.binding || null,
    statuses: req.statuses,
    ignoredSignals: req.ignoredSignals,
  };
  // A function invocation is not proof that edits were applied. Excel scripts
  // can be reads. Report only a successful client response after a write signal.
  const writeObserved = req.actions.some((a) => a.writes === true || a.name === "ppt_edits"
    || (a.name === "slide_activity" && a.status === "success"));
  const succeeded = req.actions.some((a) => a.type === "action_result" && a.ok === true);
  const failed = req.actions.some((a) => a.type === "action_result" && a.ok === false);
  result.applied = status === "completed" && writeObserved && succeeded && !failed;
  result.appliedEvidence = "Observed write signals and client action responses, not an independent document check.";
  if (req.sent && status !== "completed") {
    // Copilot can finish its edits after the request ends (seen 2026-11-13).
    result.applied = null;
    result.appliedEvidence = `Not known. The prompt was sent, and the request ended (${status}) before a final answer. Check the document.`;
  } else if (req.sent && !req.binding) {
    // The hook passes on answers and edits only after it binds. Without a
    // binding it saw none of this turn (2026-11-16: an edit applied, applied was false).
    result.applied = null;
    result.appliedEvidence = "Not known. The hook did not see this request. Check the document.";
  }
  result.matchingTabs = req.matchingTabs || 1;
  result.oversizedFrames = req.oversizedFrames || 0;
  send({ id: req.id, done: true, result });
}

function onEvent(req, event) {
  if (req.done) return;
  if (event.type === "oversized") {
    // The hook skipped a large AugLoop frame. Evidence only: it does not mean the hook saw this turn.
    req.oversizedFrames = (req.oversizedFrames || 0) + 1;
    return;
  }
  req.lastEventAt = Date.now();
  req.hookEvents += 1;
  if (event.type === "signal") {
    // Prompt signals seen while armed: bound, same, related, repeated or ignored.
    const kind = ["bound", "same", "related", "repeated", "ignored"].includes(event.kind) ? event.kind : "other";
    req.signals[kind] = (req.signals[kind] || 0) + 1;
    if (kind === "ignored" && req.ignoredSignals.length < 5)
      req.ignoredSignals.push({ reason: event.reason || null, score: typeof event.score === "number" ? event.score : null, operation: event.operation || null });
  } else if (event.type === "bound") {
    req.binding = event.mode || "unknown";
  } else if (event.type === "status" || event.type === "task_status") {
    const name = `${event.source || "task"}:${event.value}`;
    if (!req.statuses.includes(name) && req.statuses.length < 20) req.statuses.push(name);
    if (event.type === "task_status" && TASK_DONE.test(String(event.value))) req.taskDone = true;
  } else if (event.type === "started") {
    req.agentRequestId = event.requestId || req.agentRequestId;
    req.conversationId = event.conversationId || req.conversationId;
  } else if (event.type === "text") {
    req.text += event.text;
  } else if (event.type === "progress") {
    req.progress.push(event.text);
  } else if (event.type === "action" || event.type === "action_result") {
    req.actions.push(event);
  }
  send({ id: req.id, event });
  if (event.type === "completed") {
    finish(req, "completed", { text: event.text || req.text, conversationId: event.conversationId, completion: "event" });
  } else if (event.type === "error") {
    finish(req, "error", { message: event.message });
  } else if (event.type === "approval_required") {
    finish(req, "approval_required", { approval: { kind: event.kind, text: event.text } });
  }
}

// The pane UI is the fallback for the end of a turn. A final answer event from
// the hook always wins. Settle only after Copilot was seen working (a Stop
// control or a finished task status), so a slow start is not taken as the end.
function watchSettle(req) {
  const tick = async () => {
    if (req.done) return;
    let pane = null;
    try { pane = (await probe(req.tabId)).pane; } catch (e) { pane = null; }
    if (req.done) return;
    if (pane && pane.stopButton) {
      req.sawBusy = true;
      req.idlePolls = 0;
    } else if (pane) {
      req.idlePolls += 1;
    }
    const quiet = Date.now() - req.lastEventAt >= TIMING.settleQuietMs;
    if ((req.sawBusy || req.taskDone) && req.idlePolls >= TIMING.settleIdlePolls && quiet) {
      finish(req, "completed", { completion: "settled",
        message: "Copilot stopped working, but no final answer event arrived. The text can be incomplete." });
      return;
    }
    req.settleTimer = setTimeout(tick, TIMING.settlePollMs);
  };
  req.settleTimer = setTimeout(tick, TIMING.settlePollMs);
}

function setPhase(req, phase) {
  checkActive(req);
  req.phase = phase;
  req.waitState = null;
  send({ id: req.id, event: { type: "phase", phase } });
}

async function ask(msg) {
  const file = msg.file || {};
  const startedAt = Date.now();
  const timeoutMs = Math.max(TIMING.minTimeoutMs, (Number(msg.timeout) || 300) * 1000);
  const { tab, matches } = await findTab(file, msg.open === true);
  if (cancelledRequests.delete(msg.id) || !port) return;
  // Every tab of this file shows the same document. Check them all.
  for (const other of matches) {
    const running = requests.get(other.id);
    if (running) throw new Error(`busy: another request is running in tab ${other.id} for this file (phase ${running.phase}, ${ageSeconds(running.startedAt)} s).`);
    const doubt = other.id !== tab.id && uncertainTabs.get(other.id);
    if (doubt) throw new Error(`Tab ${other.id} for this file ended a request without a final answer (${doubt.reason}, ${ageSeconds(doubt.since)} s ago). Check the document, then reload that tab or run 'outlook-cli copilot app reset --tab-id ${other.id}'.`);
  }
  const uncertain = uncertainTabs.get(tab.id);
  if (uncertain) throw new Error(`The previous request ended without a final answer (${uncertain.reason}, ${ageSeconds(uncertain.since)} s ago). Copilot can still have changed the document. Check it, then reload the tab or run 'outlook-cli copilot app reset --tab-id ${tab.id}'.`);
  const req = {
    id: msg.id, tabId: tab.id, app: null, text: "", progress: [], actions: [], armedFrames: [], done: false,
    startedAt, deadline: startedAt + timeoutMs, phase: "find_tab", matchingTabs: matches.length,
    statuses: [], signals: {}, ignoredSignals: [], sendCalls: 0, binding: null, hookEvents: 0, taskDone: false, sawBusy: false, idlePolls: 0, lastEventAt: Date.now(),
  };
  requests.set(tab.id, req);
  // One timer for every phase. The waits below also stop at this deadline.
  req.timer = setTimeout(() => finish(req, "timeout", {
    message: `Copilot did not finish in time. Stopped in phase ${req.phase}${req.waitState ? ` (${req.waitState})` : ""}.` }),
  Math.max(0, req.deadline - Date.now()));
  try {
    setPhase(req, "wait_editor");
    let p = await waitForEditor(tab.id, msg.open ? 90000 : 30000, req);
    if (req.done) return;
    req.app = appOf(p);
    if (msg.app && req.app && msg.app !== req.app) throw new Error(`The tab is ${req.app}, not ${msg.app}.`);
    send({ id: req.id, event: { type: "tab", tabId: tab.id, app: req.app, title: tab.title, matchingTabs: matches.length } });
    if (!(p.pane && p.pane.promptBox)) p = await openPane(req, tab.id, p);
    if (req.done) return;
    if (msg.newChat) {
      setPhase(req, "new_chat");
      const fresh = await runIn(tab.id, p.pane.frameId, { action: "new_chat" });
      if (!fresh.ok) throw new Error(fresh.error);
      req.newChat = fresh.newChat || "clicked";
      if (req.newChat === "clicked") {
        await sleep(TIMING.inputSettleMs);
        p = await waitForPane(tab.id, 15000, req);
      }
    }
    if (req.done) return;
    setPhase(req, "arm");
    for (const editor of p.editors) {
      req.armedFrames.push(editor.frameId);
      const armed = await runIn(tab.id, editor.frameId, { action: "arm", on: true, requestId: req.id, prompt: msg.prompt });
      if (!armed.hook || !armed.hook.seen || !(armed.hook.revision >= HOOK_REVISION))
        throw new Error("The AugLoop hook is missing or out of date. Reload the Office tab.");
    }
    if (req.done) return;
    req.armedFrames.push(p.pane.frameId);
    await runIn(tab.id, p.pane.frameId, { action: "arm", on: true, requestId: req.id });
    if (!tabMatches(await chrome.tabs.get(tab.id), file)) throw new Error("The tab no longer matches the requested file.");
    const ready = await probe(tab.id);
    if (!ready.pane || ready.pane.frameId !== p.pane.frameId || !ready.pane.promptBox)
      throw new Error("The Copilot drawer is not visibly open. Nothing was sent.");
    setPhase(req, "send");
    for (const editor of p.editors) {
      const marked = await runIn(tab.id, editor.frameId, { action: "sending", requestId: req.id });
      if (!marked.ok) throw new Error(marked.error);
    }
    if (req.done) return;
    // At most one send per request. Nothing below retries it.
    if (req.sendCalls++ > 0) throw new Error("This request already sent its prompt.");
    req.sent = true; // Until the pane says otherwise: a lost reply can follow a real click.
    req.sentAt = Date.now();
    req.lastEventAt = req.sentAt;
    const sent = await runIn(tab.id, p.pane.frameId, { action: "send_prompt", text: msg.prompt, requestId: req.id }, TIMING.sendCallMs);
    if (!sent.ok) {
      // The pane answered and did not click Send, so Copilot did not start.
      if (!sent.clicked) req.sent = false;
      throw new Error(sent.error);
    }
    send({ id: req.id, event: { type: "sent" } });
    setPhase(req, "wait_answer");
    watchSettle(req);
  } catch (e) {
    if (e instanceof Stopped) return;
    // A wait that ran out at the request deadline is a timeout, not an error.
    const status = Date.now() >= req.deadline ? "timeout" : "error";
    finish(req, status, { message: `${String((e && e.message) || e)} (phase ${req.phase})` });
  }
}

// Clears a stale state for one tab after the user has checked the document.
// A request inside its time limit is never cleared: wait for it, or for its timeout.
function reset(msg) {
  const tabId = msg.tabId;
  const cleared = { request: false, uncertain: false };
  const req = requests.get(tabId);
  if (req) {
    if (Date.now() < req.deadline + TIMING.staleGraceMs)
      throw new Error(`A request is running in this tab (phase ${req.phase}, ${ageSeconds(req.startedAt)} s). Wait for it to finish or time out.`);
    finish(req, "error", { message: "The user reset a stale request." });
    requests.delete(tabId);
    cleared.request = true;
  }
  cleared.uncertain = uncertainTabs.delete(tabId);
  return { tabId, cleared };
}

function cancel(msg) {
  for (const req of requests.values()) {
    if (req.id === msg.id) {
      finish(req, "cancelled", { message: "The client cancelled the request." });
      return;
    }
  }
  cancelledRequests.add(msg.id);
  // A bounded tombstone list covers requests still finding their tab.
  if (cancelledRequests.size > 1000) cancelledRequests.delete(cancelledRequests.values().next().value);
}

async function onHostMessage(msg) {
  if (!msg || msg.v !== PROTOCOL) {
    send({ type: "error", message: `Unsupported protocol version ${msg && msg.v}.` });
    return;
  }
  if (msg.type === "hello") return;
  try {
    if (msg.type === "status") {
      send({ id: msg.id, done: true, result: await status() });
    } else if (msg.type === "inspect") {
      send({ id: msg.id, done: true, result: await inspectTab(msg.tabId) });
    } else if (msg.type === "open_pane") {
      send({ id: msg.id, done: true, result: await openPaneOnce(msg) });
    } else if (msg.type === "reset") {
      send({ id: msg.id, done: true, result: reset(msg) });
    } else if (msg.type === "ask") {
      await ask(msg);
    } else if (msg.type === "cancel") {
      cancel(msg);
    } else {
      send({ id: msg.id, done: true, error: `Unknown request type ${msg.type}.` });
    }
  } catch (e) {
    send({ id: msg.id, done: true, error: String((e && e.message) || e) });
  }
}

// Events from the controller in an editor frame. Only accept them from this
// extension's own content scripts, for the request that armed the frame.
chrome.runtime.onMessage.addListener((msg, sender) => {
  if (sender.id !== chrome.runtime.id || !sender.tab || !msg || msg.type !== "augloop") return;
  const req = requests.get(sender.tab.id);
  if (req && req.id === msg.requestId && req.armedFrames.includes(sender.frameId)) onEvent(req, msg.event);
});

chrome.tabs.onRemoved.addListener((tabId) => {
  uncertainTabs.delete(tabId);
  const req = requests.get(tabId);
  if (req) finish(req, "error", { message: "The Office tab closed. Edits may already have been applied." });
});
chrome.tabs.onUpdated.addListener((tabId, change) => {
  if (change.status === "loading") {
    uncertainTabs.delete(tabId);
    const req = requests.get(tabId);
    if (req && req.sent) finish(req, "error", { message: "The Office tab navigated during the request." });
  }
});
chrome.runtime.onStartup.addListener(connect);
chrome.runtime.onInstalled.addListener(() => reinjectAll().catch(() => {}));
chrome.alarms.create("ocli-connect", { periodInMinutes: 1 });
chrome.alarms.onAlarm.addListener(() => connect());
connect();
reinjectAll().catch(() => {});
