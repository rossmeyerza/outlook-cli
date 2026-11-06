const test = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");
const common = require("../adapters/common.js");
const excel = require("../adapters/excel.js");
const powerpoint = require("../adapters/powerpoint.js");

function hook(pathname = "/x/xlviewerinternal.aspx", prompt = "Read the table") {
  const emitted = [];
  const listeners = [];
  class WebSocket {
    constructor(url) { this.url = url; this.listeners = {}; }
    addEventListener(type, fn) { this.listeners[type] = fn; }
    send() {}
  }
  const context = vm.createContext({
    URL, WebSocket, OcliAdapters: { common, excel, powerpoint },
    location: { origin: "https://eu.officeapps.live.com", pathname },
    addEventListener: (type, fn) => listeners.push(fn),
    postMessage: (msg) => emitted.push(msg),
  });
  vm.runInContext("globalThis.window = globalThis;", context);
  const source = fs.readFileSync(path.join(__dirname, "../hook.js"), "utf8");
  vm.runInContext(source, context);
  context.listeners = listeners;
  context.armPrompt = prompt;
  vm.runInContext(`
    listeners[0]({ source: window, data: { __ocliCopilot: "controller", kind: "arm", on: true, prompt: armPrompt } });
    globalThis.socket = new WebSocket("wss://augloop.office.com");
  `, context);
  function signal(id, query, extra = {}) {
    context.record = { ops: [{ items: [{ body: {
      H_: { T_: excel.TYPES.signal }, requestId: id, content: { query }, ...extra,
    } }] }] };
    vm.runInContext("socket.send(JSON.stringify(record))", context);
  }
  function pptSignal(id, query, extra = {}) {
    context.record = { ops: [{ items: [{ body: {
      H_: { T_: powerpoint.TYPES.signal }, queryId: id, query, ...extra,
    } }] }] };
    vm.runInContext("socket.send(JSON.stringify(record))", context);
  }
  function recv(record) {
    context.record = record;
    vm.runInContext("socket.listeners.message({ data: JSON.stringify(record) })", context);
  }
  function sending() {
    vm.runInContext('listeners[0]({ source: window, data: { __ocliCopilot: "controller", kind: "sending" } })', context);
  }
  function reply(id) {
    context.record = { ops: [{ items: [{ body: {
      H_: { T_: excel.TYPES.annotation }, requestId: id,
      content: { H_: { T_: excel.TYPES.message }, status: "completed", text: "Sample table" },
    } }] }] };
    vm.runInContext("socket.listeners.message({ data: JSON.stringify(record) })", context);
  }
  function diagnostics() {
    vm.runInContext('listeners[0]({ source: window, data: { __ocliCopilot: "controller", kind: "ping" } })', context);
    return emitted.filter((m) => m.kind === "pong").at(-1);
  }
  return { signal, pptSignal, recv, sending, reply, diagnostics,
    // Plain copies: objects from the vm context have another realm's prototypes.
    events: () => JSON.parse(JSON.stringify(emitted.filter((m) => m.kind === "event").map((m) => m.event))) };
}

test("startup signals do not bind the active request", () => {
  const h = hook();
  h.signal("startup", undefined);
  h.signal("old", "An earlier prompt");
  assert.equal(h.events().length, 0);
  h.signal("current", "Read the table");
  h.reply("old");
  h.reply("current");
  const events = h.events();
  assert.deepEqual(events[0], { type: "bound", mode: "exact" });
  assert.equal(events[1].type, "started");
  assert.equal(events[1].requestId, "current");
  assert.equal(events.length, 3);
  assert.equal(events[2].type, "completed");
});

test("schema diagnostics contain names, never frame values", () => {
  const h = hook();
  h.signal("PRIVATE_REQUEST", "PRIVATE_PROMPT");
  h.reply("PRIVATE_REQUEST");
  const result = h.diagnostics();
  assert.equal(result.revision, 4);
  // Diagnostics carry the binding mode and counts, never ids or prompt text.
  assert.equal(result.request.bound, null);
  assert.equal(result.request.signals, 1);
  assert.equal(result.request.ignored, 1);
  assert(result.schemas.some((s) => s.name === excel.TYPES.signal));
  assert(!JSON.stringify(result).includes("PRIVATE_"));
  assert(!JSON.stringify(result).includes("Sample table"));
});

test("Word OfficeAgent chat uses the shared parser and labels events as Word", () => {
  const h = hook("/we/wordeditorframe.aspx");
  h.signal("word-request", "Read the table");
  h.reply("word-request");
  const events = h.events().filter((e) => e.type !== "bound");
  assert.equal(events[0].app, "word");
  assert.equal(events[1].type, "completed");
  assert.equal(events[1].text, "Sample table");
});

test("a second matching prompt is still rejected", () => {
  const h = hook();
  h.signal("first", "Read the table");
  h.signal("second", "Read the table");
  assert.equal(h.events().at(-1).type, "error");
});

const MULTILINE = 'Edit only slide 6, "Use case 1".\n\nDirection:\n- Problem statement: asks come in by email.\n- The user:  Tarik triages.';

test("a multi-line prompt binds after the editor changes its text", () => {
  const h = hook("/pods/ppt.aspx", MULTILINE);
  // The pane turned the bullets into a list, joined paragraphs, used curly quotes and a non-breaking space.
  h.pptSignal("q1", "Edit only slide 6, \u201cUse case 1\u201d.\nDirection:\n\u2022 Problem statement: asks come in by email.\n\u2022 The user:\u00a0 Tarik triages.");
  assert.deepEqual(h.events()[0], { type: "bound", mode: "normalised" });
  assert.equal(h.events()[1].requestId, "q1");
});

test("an unmatched prompt binds only after Send, and the turn completes", () => {
  const h = hook("/pods/ppt.aspx", MULTILINE);
  h.pptSignal("before", "Text that the editor changed beyond recognition");
  assert.equal(h.events().length, 0);
  h.sending();
  h.pptSignal("q1", "Text that the editor changed beyond recognition");
  assert.deepEqual(h.events()[0], { type: "bound", mode: "after_send" });
  h.recv({ ops: [{ items: [{ body: { H_: { T_: powerpoint.TYPES.output }, queryId: "q1", responseStatus: "complete",
    finalResponse: '{"assistantResponse":"Done"}' } }] }] });
  assert(h.events().some((e) => e.type === "completed" && e.text === "Done"));
});

test("task operations join the request; a repeated prompt is an error", () => {
  const h = hook("/pods/ppt.aspx", "Fix slide 2");
  h.pptSignal("q1", "Fix slide 2");
  h.pptSignal("q2", "", { taskOperationParams: { operation: "resume" } });
  h.recv({ ops: [{ items: [{ body: { H_: { T_: powerpoint.TYPES.output }, queryId: "q2", responseStatus: "complete",
    finalResponse: '{"assistantResponse":"Resumed"}' } }] }] });
  assert(h.events().some((e) => e.type === "completed" && e.text === "Resumed"));
  h.pptSignal("q3", "Fix slide 2");
  assert.equal(h.events().at(-1).type, "error");
  assert.equal(h.diagnostics().request.repeated, 1);
  assert.equal(h.diagnostics().request.related, 1);
});

test("task status names reach the worker after binding", () => {
  const h = hook("/pods/ppt.aspx", "Fix slide 2");
  const task = { H_: { T_: common.TASK_INFO }, ownerId: "o", taskStatus: "Completed" };
  h.recv(task);
  assert.equal(h.events().length, 0); // Not bound yet.
  h.pptSignal("q1", "Fix slide 2");
  h.recv(task);
  assert.deepEqual(h.events().at(-1), { type: "task_status", value: "Completed" });
});
