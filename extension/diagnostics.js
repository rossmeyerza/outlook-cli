// Read-only UI diagnostics. Export control labels and states, never document
// text, prompt values, scripts or raw protocol frames.
(function (root) {
  "use strict";
  const APPROVAL_LABELS = new Set([
    "Allow", "Allow once", "Allow this time", "Allow access", "Approve", "Confirm", "Yes, allow", "Sign in",
  ]);

  const UI_LABEL = /^(?:(?:open |close |chat with |start )?copilot(?: chat| pane| menu| options)?|new chat|new conversation|start new chat|allow editing|edit mode|editing|editing mode|approve|confirm|allow(?: once| this time| access)?|sign in|cancel|send(?: message)?|stop(?: generating| responding)?)$/i;

  function label(el) {
    return (el.getAttribute("aria-label") || el.getAttribute("title") || el.innerText || "").trim().slice(0, 100);
  }

  function control(el) {
    return {
      tag: el.tagName, role: el.getAttribute("role"), label: label(el),
      disabled: Boolean(el.disabled) || el.getAttribute("aria-disabled") === "true",
      expanded: el.getAttribute("aria-expanded"), pressed: el.getAttribute("aria-pressed"),
      checked: el.getAttribute("aria-checked"), popup: el.getAttribute("aria-haspopup"),
    };
  }

  function controls(elements, visible) {
    const found = [];
    for (const el of elements) {
      if (!el.getAttribute || !visible(el)) continue;
      const role = el.getAttribute("role");
      if (el.tagName !== "BUTTON" && !["button", "combobox", "switch", "menuitem"].includes(role)) continue;
      const name = label(el);
      if (UI_LABEL.test(name)) found.push(control(el));
      if (found.length >= 50) break;
    }
    return found;
  }

  function approval(document, visible) {
    for (const dialog of document.querySelectorAll('[role="dialog"], [role="alertdialog"]')) {
      if (!visible(dialog)) continue;
      const labels = [...dialog.querySelectorAll('button, [role="button"]')]
        .filter(visible).map(label).filter((name) => APPROVAL_LABELS.has(name));
      if (labels.length) return { kind: "permission_dialog", text: labels.join(" / ") };
    }
    return null;
  }

  function visibilityReason(el, view) {
    const r = el.getBoundingClientRect();
    if (!(r.width > 0 && r.height > 0)) return "zero_size";
    if (r.right <= 0 || r.bottom <= 0 || r.left >= view.innerWidth || r.top >= view.innerHeight) return "offscreen";
    for (let node = el; node; node = node.parentElement || (node.getRootNode && node.getRootNode().host)) {
      // aria-hidden and inert do not make pixels disappear. Office can use
      // them for accessibility/focus management on a visible editor frame.
      const style = view.getComputedStyle(node);
      if (style.display === "none") return "display_none";
      if (["hidden", "collapse"].includes(style.visibility)) return "visibility_hidden";
      if (style.opacity !== "" && Number(style.opacity) === 0) return "transparent";
      if (node !== el && style.display !== "contents" && /hidden|clip|auto|scroll/.test(style.overflow + style.overflowX + style.overflowY)) {
        const clip = node.getBoundingClientRect();
        if (clip.right <= r.left || clip.left >= r.right || clip.bottom <= r.top || clip.top >= r.bottom) return "clipped";
      }
    }
    return null;
  }

  function visibleInViewport(el, view) {
    return visibilityReason(el, view) === null;
  }

  const api = { label, control, controls, approval, visibleInViewport, visibilityReason };
  root.OcliDiagnostics = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
