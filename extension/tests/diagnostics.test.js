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

test("clipDetail reports layout facts only for a clipped drawer", () => {
  const rect = { left: 800, right: 1000, top: 0, bottom: 700, width: 200, height: 700 };
  const parent = { tagName: "DIV", innerText: "PRIVATE DOCUMENT", id: "secret-id", className: "secret-class",
    getAttribute: () => null, getBoundingClientRect: () => ({ left: 0, right: 0, top: 0, bottom: 0 }) };
  const node = { tagName: "IFRAME", getBoundingClientRect: () => rect, getAttribute: () => null, parentElement: parent };
  const style = (overflow, position) => ({ display: "block", visibility: "visible", opacity: "1", overflow, overflowX: overflow, overflowY: overflow, position, transform: "none" });
  const view = { innerWidth: 1200, innerHeight: 800, getComputedStyle: (el) => el === parent ? style("hidden", "absolute") : style("visible", "fixed") };
  assert.equal(diagnostics.visibilityReason(node, view), "clipped");
  const detail = diagnostics.clipDetail(node, view);
  assert.equal(detail.clipper.tag, "DIV");
  assert.equal(detail.clipper.depth, 1);
  assert.equal(detail.clipper.position, "absolute");
  assert.equal(detail.element.position, "fixed");
  assert.deepEqual(detail.clipper.rect, { left: 0, top: 0, right: 0, bottom: 0 });
  assert(!/PRIVATE|secret/.test(JSON.stringify(detail)));
  node.parentElement = null;
  assert.equal(diagnostics.clipDetail(node, view), null);
});

test("overflow on the document root clips at the viewport, not at the root's own box", () => {
  const doc = {};
  const rect = { left: 1528, right: 1907, top: 216, bottom: 1009, width: 379, height: 793 };
  const zero = () => ({ left: 0, right: 1908, top: 0, bottom: 0 });
  const root = { tagName: "HTML", ownerDocument: doc, getAttribute: () => null, getBoundingClientRect: zero };
  doc.documentElement = root;
  const node = { tagName: "IFRAME", ownerDocument: doc, getAttribute: () => null, getBoundingClientRect: () => rect, parentElement: root };
  const style = (overflow) => ({ display: "block", visibility: "visible", opacity: "1", overflow, overflowX: overflow, overflowY: overflow, position: "static", transform: "none" });
  const view = { innerWidth: 1908, innerHeight: 1032, getComputedStyle: (el) => el === node ? style("visible") : style("hidden") };
  assert.equal(diagnostics.visibilityReason(node, view), null);
  assert.equal(diagnostics.clipDetail(node, view), null);
  // A real container that clips the drawer is still refused.
  node.parentElement = { tagName: "DIV", ownerDocument: doc, parentElement: root, getAttribute: () => null, getBoundingClientRect: zero };
  assert.equal(diagnostics.visibilityReason(node, view), "clipped");
  assert.equal(diagnostics.clipDetail(node, view).clipper.tag, "DIV");
  // A drawer outside the window is still refused.
  node.parentElement = root;
  rect.left = 2000; rect.right = 2300;
  assert.equal(diagnostics.visibilityReason(node, view), "offscreen");
});
