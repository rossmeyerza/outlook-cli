const test = require("node:test");
const assert = require("node:assert/strict");
const vm = require("node:vm");
const fs = require("node:fs");
const path = require("node:path");
const diagnostics = require("../diagnostics.js");

test("pane approval observer reports permission and never types or clicks", async () => {
  const sent = [];
  let check;
  const rect = () => ({ width: 20, height: 20 });
  const approve = {
    innerText: "Approve", getAttribute: () => null, getBoundingClientRect: rect,
    click: () => { throw new Error("Must not approve"); },
  };
  const dialog = { getBoundingClientRect: rect, querySelectorAll: () => [approve] };
  const context = vm.createContext({
    OcliDiagnostics: diagnostics,
    location: { hostname: "x.resources.office.net", pathname: "/taskpane.html", origin: "https://x.resources.office.net" },
    chrome: { runtime: { id: "test", sendMessage: async (msg) => { sent.push(msg); } } },
    document: { querySelectorAll: () => [dialog], execCommand: () => { throw new Error("Must not type"); } },
    addEventListener() {}, postMessage() {}, setTimeout,
    setInterval: (fn) => { check = fn; return 1; }, clearInterval() {},
  });
  vm.runInContext("globalThis.window = globalThis", context);
  vm.runInContext(fs.readFileSync(path.join(__dirname, "../controller.js"), "utf8"), context);
  await vm.runInContext('__ocliController.run({ action: "arm", on: true, requestId: "req" })', context);
  check();
  assert.equal(sent[0].event.type, "approval_required");
  assert.equal(sent[0].event.text, "Approve");
  const result = await vm.runInContext('__ocliController.run({ action: "send_prompt", text: "Hello", requestId: "req" })', context);
  assert.equal(result.ok, false);
  assert.match(result.error, /Nothing was approved/);
  await vm.runInContext('__ocliController.run({ action: "arm", on: false })', context);
});
