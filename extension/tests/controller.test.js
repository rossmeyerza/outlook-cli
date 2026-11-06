const test = require("node:test");
const assert = require("node:assert/strict");

global.window = { addEventListener() {} };
global.chrome = { runtime: { id: "extension" } };
const controller = require("../controller.js");

test("editor and pane frame detection", () => {
  assert.deepEqual(controller.frameRole({ hostname: "eu.officeapps.live.com", pathname: "/x/_layouts/xlviewerinternal.aspx" }), { role: "editor", app: "excel" });
  assert.deepEqual(controller.frameRole({ hostname: "eu.officeapps.live.com", pathname: "/pods/ppt.aspx" }), { role: "editor", app: "powerpoint" });
  assert.deepEqual(controller.frameRole({ hostname: "x.resources.office.net", pathname: "/thing/taskpane.html" }), { role: "pane", app: null });
  assert.equal(controller.frameRole({ hostname: "evil.test", pathname: "/pods/ppt.aspx" }).role, "other");
});

function fakeElement(label, opts = {}) {
  return {
    tagName: opts.editable ? "DIV" : "BUTTON", isContentEditable: opts.editable || false,
    getAttribute: (name) => name === "aria-label" ? label : null,
    getBoundingClientRect: () => ({ width: opts.hidden ? 0 : 10, height: 10 }),
  };
}

test("prompt selectors find only labelled editable elements", () => {
  const wrong = fakeElement("Chat input");
  const hidden = fakeElement("Chat input", { editable: true, hidden: true });
  const target = fakeElement("Chat input", { editable: true });
  const nodes = [wrong, hidden, target];
  let i = 0;
  const root = { createTreeWalker: () => ({ currentNode: nodes[0], nextNode: () => nodes[++i] }) };
  assert.equal(controller.findByLabel(controller.PROMPT_LABELS, { editable: true, root }), target);
});

test("the click allowlist cannot approve, keep, undo, or change editing", () => {
  for (const label of ["Done", "Undo", "Keep", "Allow editing", "Approve", "Confirm", "Apply"])
    assert.throws(() => controller.safeClick(fakeElement(label), [label]), /refused/);
  assert.throws(() => controller.safeClick(fakeElement("Share"), controller.CLICKABLE.send), /refused/);
});

test("the Stop control is detected for status only and is never clickable", () => {
  for (const label of ["Stop", "Stop generating", "Stop responding"]) assert.match(label, controller.STOP_LABEL);
  assert.doesNotMatch("Stopwatch", controller.STOP_LABEL);
  for (const labels of Object.values(controller.CLICKABLE))
    assert(!labels.some((label) => controller.STOP_LABEL.test(label)));
  const nodes = [fakeElement("Stop generating", { hidden: true }), fakeElement("Stop generating")];
  let i = 0;
  const root = { createTreeWalker: () => ({ currentNode: nodes[0], nextNode: () => nodes[++i] }) };
  assert.equal(controller.findByPattern(controller.STOP_LABEL, root), nodes[1]);
});

test("New chat is skipped on a fresh pane and refused while Copilot works", async () => {
  const pane = (labels) => {
    const nodes = labels.map((label) => fakeElement(label));
    nodes.forEach((n) => { n.clicked = 0; n.click = () => { n.clicked += 1; }; });
    let i = 0;
    globalThis.document = { createTreeWalker: () => ({ currentNode: nodes[0], nextNode: () => nodes[++i] }) };
    return nodes;
  };
  const run = (cmd) => globalThis.__ocliController.run(cmd);
  pane(["Close Copilot chat"]);
  assert.deepEqual(await run({ action: "new_chat" }), { ok: true, newChat: "skipped" });
  const [button] = pane(["New chat"]);
  assert.deepEqual(await run({ action: "new_chat" }), { ok: true, newChat: "clicked" });
  assert.equal(button.clicked, 1);
  pane(["Stop generating"]);
  assert.equal((await run({ action: "new_chat" })).ok, false);
});

test("a request types its prompt at most once", async () => {
  const box = fakeElement("Chat input", { editable: true });
  box.innerText = "";
  box.focus = () => {};
  const send = fakeElement("Send");
  send.disabled = false;
  send.click = () => { send.clicks = (send.clicks || 0) + 1; };
  const nodes = [box, send];
  let typed = 0;
  globalThis.document = {
    createTreeWalker: () => { let i = 0; return { currentNode: nodes[0], nextNode: () => nodes[++i] }; },
    querySelectorAll: () => [],
    execCommand: () => { typed += 1; return true; },
  };
  controller.setArmed("r1");
  assert.deepEqual(await controller.sendPrompt("Hello", "r1"), { ok: true, typed: true, clicked: true });
  const again = await controller.sendPrompt("Hello", "r1");
  assert.equal(again.ok, false);
  assert.match(again.error, /never sent twice/);
  assert.equal(typed, 1);
  assert.equal(send.clicks, 1);
  controller.setArmed(null);
});
