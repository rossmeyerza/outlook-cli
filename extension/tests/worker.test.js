const test = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");

function worker() {
  const sent = [];
  const calls = [];
  const listeners = {};
  const event = (name) => ({ addListener: (fn) => { listeners[name] = fn; } });
  const tab = { id: 7, url: "https://tenant.sharepoint.com/_layouts/15/Doc.aspx?sourcedoc=%7Btest-guid%7D", title: "Test.xlsx" };
  const chrome = {
    runtime: {
      id: "extension", getManifest: () => ({ version: "test" }),
      connectNative: () => ({ postMessage: (m) => sent.push(m), onMessage: event("native"), onDisconnect: event("disconnect") }),
      onMessage: event("message"), onStartup: event("startup"), onInstalled: event("installed"),
    },
    tabs: {
      query: async () => [tab], get: async () => tab,
      create: async () => tab, onRemoved: event("removed"), onUpdated: event("updated"),
    },
    scripting: { executeScript: async ({ args, files, target }) => {
      if (files) return [];
      const cmd = args[0];
      calls.push(cmd);
      if (cmd.action === "describe") return [
        { frameId: 1, result: { role: "editor", isTop: true, app: "excel", paneButton: true, hook: { seen: true, revision: 2 }, embeddedFrames: [{ url: "https://pane.resources.office.net/taskpane.html", visible: true }] } },
        { frameId: 2, result: { role: "pane", url: "https://pane.resources.office.net/taskpane.html", promptBox: true } },
      ];
      return [{ frameId: target.frameIds[0], result: { ok: true, hook: { seen: true, revision: 2 } } }];
    } },
    alarms: { create() {}, onAlarm: event("alarm") },
  };
  const context = vm.createContext({ chrome, console, setTimeout, clearTimeout, URL });
  const code = fs.readFileSync(path.join(__dirname, "../sw.js"), "utf8");
  vm.runInContext(code + "\n globalThis.testApi = { ask, tabMatches, requests, uncertainTabs, onEvent, finish, framePresented };", context);
  return { api: context.testApi, sent, calls, listeners, tab, chrome };
}

test("host disconnect consumes lastError, reconnects transport and never replays prompts", async () => {
  const { api, chrome, listeners, sent, calls } = worker();
  await api.ask({ id: "req", prompt: "Hello", file: { uniqueId: "test-guid" } });
  let read = false;
  Object.defineProperty(chrome.runtime, "lastError", { get() { read = true; return { message: "Native host has exited." }; } });
  const alarms = [];
  chrome.alarms.create = (...args) => alarms.push(args);
  listeners.disconnect();
  assert.equal(read, true);
  assert.equal(api.requests.size, 0);
  assert.equal(api.uncertainTabs.has(7), true);
  assert.equal(alarms[0][0], "ocli-reconnect");
  const prompts = calls.filter((c) => c.action === "send_prompt").length;
  listeners.alarm();
  assert.equal(calls.filter((c) => c.action === "send_prompt").length, prompts);
  assert.equal(sent.filter((s) => s.type === "hello").length, 2);
});

test("pane readiness follows the frame that embeds the drawer", () => {
  const { api } = worker();
  const pane = { frameId: 2, url: "pane" };
  const editor = { frameId: 1, url: "editor", embeddedFrames: [{ url: "pane", visible: false }] };
  // Live Word data: the SharePoint host reports its editor iframe as hidden.
  const host = { frameId: 0, isTop: true, embeddedFrames: [{ url: "editor", visible: false }] };
  // Closed drawer: the pane is mounted, but its parent does not show it.
  assert.equal(api.framePresented(pane, [host, editor, pane]), false);
  editor.embeddedFrames[0].visible = true;
  assert.equal(api.framePresented(pane, [host, editor, pane]), true);
  // No frame claims the pane: readiness is unknown, so it is not presented.
  assert.equal(api.framePresented(pane, [host, pane]), false);
});

test("tab matching uses the GUID, not the title", () => {
  const { api, tab } = worker();
  assert(api.tabMatches(tab, { uniqueId: "TEST-GUID" }));
  assert(!api.tabMatches(tab, { uniqueId: "other-guid", name: "Test.xlsx" }));
});

test("one request per tab, correlated controller events, final result", async () => {
  const { api, sent, listeners, calls } = worker();
  const request = { id: "req", prompt: "Hello", app: "excel", file: { uniqueId: "test-guid" } };
  await api.ask(request);
  await assert.rejects(api.ask({ ...request, id: "second" }), /busy/);
  const sendEvent = (frameId, event) => listeners.message({ type: "augloop", requestId: "req", event }, { id: "extension", tab: { id: 7 }, frameId });
  sendEvent(99, { type: "completed", text: "forged" });
  assert.equal(api.requests.size, 1);
  sendEvent(1, { type: "text", text: "Hello" });
  sendEvent(1, { type: "completed", text: "Hello world" });
  assert.equal(api.requests.size, 0);
  const result = sent.find((s) => s.done).result;
  assert.equal(result.text, "Hello world");
  assert.equal(result.applied, false);
  assert(calls.some((c) => c.action === "send_prompt"));
});

test("an invocation alone does not prove edits were applied", async () => {
  const { api, sent } = worker();
  await api.ask({ id: "req", prompt: "Hello", file: { uniqueId: "test-guid" } });
  const req = api.requests.get(7);
  api.onEvent(req, { type: "action", app: "excel", name: "executeScript", writes: true });
  api.onEvent(req, { type: "action_result", app: "excel", ok: false });
  api.onEvent(req, { type: "completed", text: "Could not edit" });
  assert.equal(sent.find((s) => s.done).result.applied, false);
});

test("timeout blocks another request until the tab is checked and reloaded", async () => {
  const { api } = worker();
  const msg = { id: "req", prompt: "Hello", file: { uniqueId: "test-guid" } };
  await api.ask(msg);
  api.finish(api.requests.get(7), "timeout");
  await assert.rejects(api.ask({ ...msg, id: "second" }), /reload/);
});

test("approval returns without issuing an approval command", async () => {
  const { api, sent, calls } = worker();
  await api.ask({ id: "req", prompt: "Hello", file: { uniqueId: "test-guid" } });
  api.onEvent(api.requests.get(7), { type: "approval_required", kind: "pending_action", text: "Approve?" });
  const result = sent.find((s) => s.done).result;
  assert.equal(result.status, "approval_required");
  assert.equal(result.approval.text, "Approve?");
  assert(!calls.some((c) => /approve|allow|keep/i.test(c.action)));
});
