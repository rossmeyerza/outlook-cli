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

let port = null;
let lastHostError = null;
const requests = new Map(); // tab id -> active request
const uncertainTabs = new Set(); // A timeout does not stop the Office agent.
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

async function runInFrames(tabId, cmd, frameIds) {
  const target = frameIds ? { tabId, frameIds } : { tabId, allFrames: true };
  const results = await chrome.scripting.executeScript({
    target,
    func: (c) => (globalThis.__ocliController ? globalThis.__ocliController.run(c) : null),
    args: [cmd],
  });
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
  const frames = await runInFrames(tabId, { action: "describe" }).catch(() => []);
  const editors = frames.filter((f) => f.role === "editor");
  const panes = frames.filter((f) => f.role === "pane" && framePresented(f, frames));
  return { frames, editors, pane: panes.find((p) => p.promptBox) || panes[0] || null };
}

async function runIn(tabId, frameId, cmd) {
  const [result] = await runInFrames(tabId, cmd, [frameId]);
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

async function describeTab(tab) {
  const p = await probe(tab.id);
  const editor = p.editors[0] || null;
  return {
    tabId: tab.id,
    title: tab.title,
    active: tab.active,
    app: appOf(p),
    editorReady: Boolean(editor),
    hooked: Boolean(editor && editor.hook && editor.hook.seen),
    hookRevision: editor && editor.hook ? editor.hook.revision : null,
    controllerVersion: editor ? editor.version : null,
    augloopSockets: editor && editor.hook ? editor.hook.sockets : 0,
    paneOpen: Boolean(p.pane && p.pane.promptBox),
    allowEditing: p.pane ? p.pane.allowEditing : null,
    busy: requests.has(tab.id) || uncertainTabs.has(tab.id),
    uncertain: uncertainTabs.has(tab.id),
  };
}

async function status() {
  const tabs = await chrome.tabs.query({ url: OFFICE_URLS });
  const described = [];
  for (const tab of tabs) {
    const info = await describeTab(tab);
    if (info.editorReady) described.push(info);
  }
  return { extension: VERSION, protocol: PROTOCOL, lastHostError, tabs: described };
}

async function inspectTab(tabId) {
  const tab = await chrome.tabs.get(tabId);
  let url;
  try { url = new URL(tab.url); } catch (e) { throw new Error("The tab has no valid URL."); }
  if (url.protocol !== "https:" || !/(^|\.)(sharepoint\.com|officeapps\.live\.com|cloud\.microsoft)$/.test(url.hostname))
    throw new Error("Only Office browser tabs can be inspected.");
  return { extension: VERSION, tabId, frames: await runInFrames(tabId, { action: "inspect" }) };
}

async function findTab(file, open) {
  const tabs = await chrome.tabs.query({});
  const tab = tabs.find((t) => tabMatches(t, file));
  if (tab) return tab;
  if (!open) throw new Error("Open the file in Chrome first, or use --open.");
  if (!file.webUrl) throw new Error("The file has no web URL to open.");
  return chrome.tabs.create({ url: file.webUrl, active: true });
}

async function waitForEditor(tabId, timeoutMs) {
  const end = Date.now() + timeoutMs;
  while (Date.now() < end) {
    const p = await probe(tabId);
    if (p.editors.some((e) => e.paneButton) || (p.editors.length && p.pane && p.pane.promptBox)) return p;
    await sleep(1000);
  }
  throw new Error("The Office editor did not load in time.");
}

async function waitForPane(tabId, timeoutMs) {
  const end = Date.now() + timeoutMs;
  while (Date.now() < end) {
    const p = await probe(tabId);
    if (p.pane && p.pane.promptBox) return p;
    await sleep(500);
  }
  throw new Error("The Copilot pane did not open in time.");
}

// ---- requests -----------------------------------------------------------

function finish(req, status, extra = {}) {
  if (req.done) return;
  req.done = true;
  clearTimeout(req.timer);
  requests.delete(req.tabId);
  if (req.sent && status !== "completed") uncertainTabs.add(req.tabId);
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
  };
  // A function invocation is not proof that edits were applied. Excel scripts
  // can be reads. Report only a successful client response after a write signal.
  const writeObserved = req.actions.some((a) => a.writes === true || (a.name === "slide_activity" && a.status === "success"));
  const succeeded = req.actions.some((a) => a.type === "action_result" && a.ok === true);
  const failed = req.actions.some((a) => a.type === "action_result" && a.ok === false);
  result.applied = status === "completed" && writeObserved && succeeded && !failed;
  result.appliedEvidence = "Observed write signals and client action responses, not an independent document check.";
  send({ id: req.id, done: true, result });
}

function onEvent(req, event) {
  if (req.done) return;
  if (event.type === "started") {
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
    finish(req, "completed", { text: event.text || req.text, conversationId: event.conversationId });
  } else if (event.type === "error") {
    finish(req, "error", { message: event.message });
  } else if (event.type === "approval_required") {
    finish(req, "approval_required", { approval: { kind: event.kind, text: event.text } });
  }
}

async function ask(msg) {
  const file = msg.file || {};
  const tab = await findTab(file, msg.open === true);
  if (cancelledRequests.delete(msg.id) || !port) return;
  if (requests.has(tab.id)) throw new Error("busy: another request is running in this tab.");
  if (uncertainTabs.has(tab.id)) throw new Error("The previous request ended without a final answer. Check the document, then reload the tab before retrying.");
  const req = {
    id: msg.id, tabId: tab.id, app: null, text: "", progress: [], actions: [], armedFrames: [], done: false,
  };
  requests.set(tab.id, req);
  try {
    let p = await waitForEditor(tab.id, msg.open ? 90000 : 30000);
    if (req.done) return;
    req.app = appOf(p);
    if (msg.app && req.app && msg.app !== req.app) throw new Error(`The tab is ${req.app}, not ${msg.app}.`);
    send({ id: req.id, event: { type: "tab", tabId: tab.id, app: req.app, title: tab.title } });
    if (!(p.pane && p.pane.promptBox)) {
      const editor = p.editors.find((e) => e.paneButton) || p.editors[0];
      const opened = await runIn(tab.id, editor.frameId, { action: "open_pane" });
      if (!opened.ok) throw new Error(opened.error);
      p = await waitForPane(tab.id, 30000);
      // The input can be on screen before it accepts typing.
      await sleep(1500);
      p = await waitForPane(tab.id, 5000);
    }
    if (req.done) return;
    if (msg.newChat) {
      const fresh = await runIn(tab.id, p.pane.frameId, { action: "new_chat" });
      if (!fresh.ok) throw new Error(fresh.error);
      await sleep(1500);
      p = await waitForPane(tab.id, 15000);
    }
    if (req.done) return;
    for (const editor of p.editors) {
      req.armedFrames.push(editor.frameId);
      const armed = await runIn(tab.id, editor.frameId, { action: "arm", on: true, requestId: req.id, prompt: msg.prompt });
      if (!armed.hook || !armed.hook.seen || armed.hook.revision < (req.app === "word" ? 3 : 2))
        throw new Error("The AugLoop hook is missing or out of date. Reload the Office tab.");
    }
    if (req.done) return;
    req.timer = setTimeout(() => finish(req, "timeout", { message: "Copilot did not finish in time." }),
      Math.max(10, Number(msg.timeout) || 300) * 1000);
    req.armedFrames.push(p.pane.frameId);
    await runIn(tab.id, p.pane.frameId, { action: "arm", on: true, requestId: req.id });
    if (!tabMatches(await chrome.tabs.get(tab.id), file)) throw new Error("The tab no longer matches the requested file.");
    const ready = await probe(tab.id);
    if (!ready.pane || ready.pane.frameId !== p.pane.frameId || !ready.pane.promptBox)
      throw new Error("The Copilot drawer is not visibly open. Nothing was sent.");
    if (req.done) return;
    req.sent = true; // Sending can have side effects even if the response is lost.
    const sent = await runIn(tab.id, p.pane.frameId, { action: "send_prompt", text: msg.prompt, requestId: req.id });
    if (!sent.ok) throw new Error(sent.error);
    send({ id: req.id, event: { type: "sent" } });
  } catch (e) {
    finish(req, "error", { message: String((e && e.message) || e) });
  }
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
