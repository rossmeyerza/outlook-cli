const test = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");
const common = require("../adapters/common.js");
const excel = require("../adapters/excel.js");
const powerpoint = require("../adapters/powerpoint.js");

function hook(pathname = "/x/xlviewerinternal.aspx") {
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
  vm.runInContext(`
    listeners[0]({ source: window, data: { __ocliCopilot: "controller", kind: "arm", on: true, prompt: "Read the table" } });
    globalThis.socket = new WebSocket("wss://augloop.office.com");
  `, context);
  function signal(id, query) {
    context.record = { ops: [{ items: [{ body: {
      H_: { T_: excel.TYPES.signal }, requestId: id, content: { query },
    } }] }] };
    vm.runInContext("socket.send(JSON.stringify(record))", context);
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
  return { signal, reply, diagnostics, events: () => emitted.filter((m) => m.kind === "event").map((m) => m.event) };
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
  assert.equal(events[0].type, "started");
  assert.equal(events[0].requestId, "current");
  assert.equal(events.length, 2);
  assert.equal(events[1].type, "completed");
});

test("schema diagnostics contain names, never frame values", () => {
  const h = hook();
  h.signal("PRIVATE_REQUEST", "PRIVATE_PROMPT");
  h.reply("PRIVATE_REQUEST");
  const result = h.diagnostics();
  assert.equal(result.revision, 3);
  assert(result.schemas.some((s) => s.name === excel.TYPES.signal));
  assert(!JSON.stringify(result).includes("PRIVATE_"));
  assert(!JSON.stringify(result).includes("Sample table"));
});

test("Word OfficeAgent chat uses the shared parser and labels events as Word", () => {
  const h = hook("/we/wordeditorframe.aspx");
  h.signal("word-request", "Read the table");
  h.reply("word-request");
  assert.equal(h.events()[0].app, "word");
  assert.equal(h.events()[1].type, "completed");
  assert.equal(h.events()[1].text, "Sample table");
});

test("a second matching prompt is still rejected", () => {
  const h = hook();
  h.signal("first", "Read the table");
  h.signal("second", "Read the table");
  assert.equal(h.events().at(-1).type, "error");
});
