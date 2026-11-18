// Read-only UI diagnostics. Export control labels and states, never document
// text, prompt values, scripts or raw protocol frames.
(function (root) {
  "use strict";
  const APPROVAL_LABELS = new Set([
    "Allow", "Allow once", "Allow this time", "Allow access", "Approve", "Confirm", "Yes, allow", "Sign in",
  ]);

  const UI_LABEL = /^(?:(?:open |close |chat with |start )?copilot(?: chat| pane| menu| options)?|new chat|new conversation|start new chat|allow editing|edit mode|editing|editing mode|approve|confirm|allow(?: once| this time| access)?|sign in|cancel|send(?: message)?|stop(?: generating| responding)?)(?:\s*\((?:(?:alt|ctrl|control|shift|cmd|command|option|meta)\s*\+\s*)+[a-z0-9]{1,3}\))?$/i;

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

  // Overflow on the root element clips at the viewport, not at the root's own
  // box (CSS Overflow 3, viewport propagation). Excel on the web (2026-11-18) sets
  // overflow:hidden on a zero-height <html>, so its rect must not count as a
  // clip. The viewport itself is covered by the offscreen check.
  function isDocumentRoot(node) {
    return Boolean(node.ownerDocument && node === node.ownerDocument.documentElement);
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
        if (!isDocumentRoot(node) && (clip.right <= r.left || clip.left >= r.right || clip.bottom <= r.top || clip.top >= r.bottom)) return "clipped";
      }
    }
    return null;
  }

  // Read-only layout facts for a "clipped" result: which ancestor clipped the
  // element, and how the element and the ancestors in between are positioned.
  // Tag names, CSS values and rectangles only, never text, ids or class names.
  function clipDetail(el, view) {
    const box = (r) => ({ left: Math.round(r.left), top: Math.round(r.top), right: Math.round(r.right), bottom: Math.round(r.bottom) });
    const r = el.getBoundingClientRect();
    const positioned = [];
    let depth = 0;
    for (let node = el; node; node = node.parentElement || (node.getRootNode && node.getRootNode().host), depth += 1) {
      const style = view.getComputedStyle(node);
      if (node !== el && style.display !== "contents" && /hidden|clip|auto|scroll/.test(style.overflow + style.overflowX + style.overflowY)) {
        const clip = node.getBoundingClientRect();
        if (!isDocumentRoot(node) && (clip.right <= r.left || clip.left >= r.right || clip.bottom <= r.top || clip.top >= r.bottom)) {
          return {
            element: { position: view.getComputedStyle(el).position, rect: box(r) },
            clipper: { tag: node.tagName, depth, position: style.position, overflowX: style.overflowX, overflowY: style.overflowY,
              transform: Boolean(style.transform && style.transform !== "none"), rect: box(clip) },
            positionedBetween: positioned,
            viewport: { width: view.innerWidth, height: view.innerHeight },
          };
        }
      }
      if (node !== el && style.position && style.position !== "static") positioned.push({ depth, position: style.position });
    }
    return null;
  }

  function visibleInViewport(el, view) {
    return visibilityReason(el, view) === null;
  }

  const api = { label, control, controls, approval, visibleInViewport, visibilityReason, clipDetail };
  root.OcliDiagnostics = api;
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
