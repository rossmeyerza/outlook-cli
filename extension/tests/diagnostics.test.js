const test = require("node:test");
const assert = require("node:assert/strict");
const diagnostics = require("../diagnostics.js");

function button(name, options = {}) {
  return {
    tagName: "BUTTON", innerText: name, disabled: false,
    getAttribute: (key) => ({ "aria-label": name, "aria-expanded": options.expanded || null })[key] || null,
    click: () => { throw new Error("Diagnostics must never click"); },
  };
}
const visible = (el) => el.hidden !== true;

test("inspect returns UI labels and states, never input values", () => {
  const target = button("Chat with Copilot", { expanded: "false" });
  target.value = "PRIVATE PROMPT";
  const other = button("Slide title: PRIVATE DOCUMENT");
  const output = diagnostics.controls([target, other], visible);
  assert.equal(output.length, 1);
  assert.equal(output[0].expanded, "false");
  assert(!JSON.stringify(output).includes("PRIVATE"));
});

test("approval controls in a visible modal stop collection without clicking", () => {
  const approve = button("Allow once");
  const dialog = { querySelectorAll: () => [approve, button("Cancel")], innerText: "PRIVATE DOCUMENT PREVIEW" };
  const doc = { querySelectorAll: () => [dialog] };
  assert.deepEqual(diagnostics.approval(doc, visible), { kind: "permission_dialog", text: "Allow once" });
});

test("Allow editing, Keep and Undo are not treated as pending permission", () => {
  const doc = { querySelectorAll: () => [{ querySelectorAll: () => [button("Allow editing"), button("Keep"), button("Undo")] }] };
  assert.equal(diagnostics.approval(doc, visible), null);
});

test("parent-side iframe visibility rejects hidden, clipped and offscreen drawers", () => {
  const rect = { left: 800, right: 1000, top: 0, bottom: 700, width: 200, height: 700 };
  const node = { getBoundingClientRect: () => rect, getAttribute: () => null };
  const view = { innerWidth: 1200, innerHeight: 800, getComputedStyle: (el) => ({ display: el.hidden ? "none" : "block", visibility: "visible", opacity: "1" }) };
  assert.equal(diagnostics.visibleInViewport(node, view), true);
  node.parentElement = { hidden: true, getAttribute: () => null };
  assert.equal(diagnostics.visibleInViewport(node, view), false);
  node.parentElement = { getAttribute: () => null, getBoundingClientRect: () => ({ left: 0, right: 0, top: 0, bottom: 0 }) };
  view.getComputedStyle = () => ({ display: "block", visibility: "visible", opacity: "1", overflow: "hidden" });
  assert.equal(diagnostics.visibleInViewport(node, view), false);
  delete node.parentElement;
  rect.left = 1300; rect.right = 1500;
  assert.equal(diagnostics.visibleInViewport(node, view), false);
});

test("accessibility hiding does not hide visible editor pixels", () => {
  const el = { inert: true, getAttribute: () => "true", getBoundingClientRect: () => ({ left: 0, right: 500, top: 0, bottom: 500, width: 500, height: 500 }) };
  const view = { innerWidth: 1000, innerHeight: 800, getComputedStyle: () => ({ display: "block", visibility: "visible", opacity: "1" }) };
  assert.equal(diagnostics.visibleInViewport(el, view), true);
  view.getComputedStyle = () => ({ display: "none", visibility: "visible", opacity: "1" });
  assert.equal(diagnostics.visibilityReason(el, view), "display_none");
});

test("hidden permission dialogs are ignored", () => {
  const dialog = { hidden: true, querySelectorAll: () => [button("Confirm")] };
  assert.equal(diagnostics.approval({ querySelectorAll: () => [dialog] }, visible), null);
});
