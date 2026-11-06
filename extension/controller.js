// Isolated world, all Office frames. Opens the Copilot pane, types and sends
// prompts, and relays hook events to the service worker while a request runs.
// Rules: find controls by aria-label, keep references in memory, never write to
// the DOM, and click only the controls in CLICKABLE.
(() => {
  "use strict";

  const VERSION = 6;
  const childUrls = new WeakMap();
  const diagnostics = globalThis.OcliDiagnostics || (typeof require === "function" ? require("./diagnostics.js") : null);
  const CHANNEL = "__ocliCopilot";
  const CLICKABLE = {
    openPane: ["Chat with Copilot"],
    send: ["Send"],
    newChat: ["New chat", "New conversation", "Start new chat"],
  };
  const PROMPT_LABELS = ["Chat input", "Type your message"];
  // Read only, never clicked: a visible Stop control means Copilot is still working.
  const STOP_LABEL = /^stop\b/i;
  // Labels that must never be clicked, even if a page uses them for an allowed name.
  const FORBIDDEN = /\b(done|undo|keep|allow|approve|accept|confirm|apply|discard|delete|share)\b/i;

  function frameRole(loc) {
    const host = loc.hostname;
    const path = loc.pathname.toLowerCase();
    if (/\.officeapps\.live\.com$/.test(host)) {
      if (path.includes("xlviewerinternal.aspx")) return { role: "editor", app: "excel" };
      if (path.includes("/pods/ppt.aspx") || path.includes("powerpointframe.aspx")) return { role: "editor", app: "powerpoint" };
      if (path.includes("wordeditorframe.aspx")) return { role: "editor", app: "word" };
      return { role: "office", app: null };
    }
    if (/\.resources\.office\.net$/.test(host) && path.includes("taskpane")) return { role: "pane", app: null };
    if (/\.sharepoint\.com$/.test(host)) return { role: "host", app: null };
    return { role: "other", app: null };
  }

  function* allElements(root) {
    const walker = root.createTreeWalker ? root.createTreeWalker(root, 1) : document.createTreeWalker(root, 1);
    let node = walker.currentNode;
    while (node) {
      yield node;
      if (node.shadowRoot) yield* allElements(node.shadowRoot);
      node = walker.nextNode();
    }
  }

  function visible(el) {
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0;
  }

  function labelOf(el) {
    return (el.getAttribute("aria-label") || el.getAttribute("title") || "").trim();
  }

  function findByLabel(labels, { editable = false, root = document } = {}) {
    for (const el of allElements(root)) {
      if (!el.getAttribute) continue;
      if (editable && !el.isContentEditable) continue;
      if (!editable && el.tagName !== "BUTTON" && el.getAttribute("role") !== "button") continue;
      if (labels.includes(labelOf(el)) && visible(el)) return el;
    }
    return null;
  }

  function findByPattern(pattern, root = document) {
    for (const el of allElements(root)) {
      if (!el.getAttribute || (el.tagName !== "BUTTON" && el.getAttribute("role") !== "button")) continue;
      if (pattern.test(labelOf(el)) && visible(el)) return el;
    }
    return null;
  }

  function enabled(el) {
    return !el.disabled && el.getAttribute("aria-disabled") !== "true";
  }

  function safeClick(el, allowed) {
    const label = labelOf(el);
    if (!allowed.includes(label) || FORBIDDEN.test(label)) throw new Error(`refused to click "${label}"`);
    // One activation only. A synthetic pointer sequence plus click can trigger
    // both pointer and click handlers in the Office pane.
    el.click();
  }

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  async function waitFor(fn, timeoutMs, stepMs = 200) {
    const end = Date.now() + timeoutMs;
    for (;;) {
      const value = fn();
      if (value || Date.now() > end) return value;
      await sleep(stepMs);
    }
  }

  // Hook state, from the main world.
  const hook = { seen: false, armed: false, sockets: 0, revision: null, schemas: [], request: null };
  let armedRequest = null;
  let approvalTimer = null;

  function relay(event) {
    if (!armedRequest) return;
    try {
      chrome.runtime.sendMessage({ type: "augloop", requestId: armedRequest, event }).catch(() => {});
    } catch (e) {
      // The extension was reloaded. A new controller takes over.
    }
  }

  function watchApprovals() {
    clearInterval(approvalTimer);
    approvalTimer = null;
    if (!armedRequest || !["pane", "editor"].includes(frameRole(location).role)) return;
    approvalTimer = setInterval(() => {
      if (globalThis.__ocliController !== self || !alive()) {
        clearInterval(approvalTimer);
        return;
      }
      const pending = diagnostics.approval(document, visible);
      if (pending) {
        relay({ type: "approval_required", ...pending });
        clearInterval(approvalTimer);
      }
    }, 250);
  }

  function toHook(kind, extra = {}) {
    window.postMessage({ [CHANNEL]: "controller", kind, ...extra }, location.origin);
  }

  async function pingHook() {
    toHook("ping");
    await sleep(150);
    return { ...hook };
  }

  function allowEditingState() {
    for (const el of allElements(document)) {
      if (!el.getAttribute) continue;
      const label = labelOf(el);
      if (/allow editing|edit mode/i.test(label)) {
        return { label, checked: el.getAttribute("aria-checked") || el.getAttribute("aria-pressed") || null };
      }
    }
    return null;
  }

  async function embeddedFrames() {
    const elements = [...allElements(document)].filter((el) => el.tagName === "IFRAME");
    for (const el of elements) {
      try { el.contentWindow.postMessage({ [CHANNEL]: "frame_probe" }, "*"); } catch (e) { /* detached frame */ }
    }
    if (elements.length) await sleep(100);
    const frames = [];
    for (const el of allElements(document)) {
      if (el.tagName !== "IFRAME") continue;
      let url;
      try { url = new URL(childUrls.get(el) || el.src); } catch (e) { continue; }
      if (!/\.resources\.office\.net$|\.officeapps\.live\.com$/.test(url.hostname)) continue;
      const reason = diagnostics.visibilityReason(el, window);
      frames.push({ url: url.origin + url.pathname, visible: reason === null, visibilityReason: reason });
    }
    return frames;
  }

  async function describe() {
    const { role, app } = frameRole(location);
    const info = { version: VERSION, role, app, url: location.origin + location.pathname,
      isTop: window === window.top, embeddedFrames: await embeddedFrames() };
    if (role === "editor") {
      info.hook = await pingHook();
      info.paneButton = Boolean(findByLabel(CLICKABLE.openPane));
    } else if (role === "pane") {
      info.promptBox = Boolean(findByLabel(PROMPT_LABELS, { editable: true }));
      info.sendButton = Boolean(findByLabel(CLICKABLE.send));
      info.newChatButton = Boolean(findByLabel(CLICKABLE.newChat));
      info.stopButton = Boolean(findByPattern(STOP_LABEL));
      info.allowEditing = allowEditingState();
    }
    return info;
  }

  async function inspect() {
    const info = await describe();
    info.controls = diagnostics.controls(allElements(document), visible);
    if (info.role === "pane") {
      info.approval = diagnostics.approval(document, visible);
      const box = findByLabel(PROMPT_LABELS, { editable: true });
      info.draftPresent = Boolean(box && (box.innerText || "").trim());
    }
    return info;
  }

  async function openPane() {
    const button = findByLabel(CLICKABLE.openPane);
    if (!button) return { ok: false, error: "The Copilot button is not on the page." };
    safeClick(button, CLICKABLE.openPane);
    return { ok: true };
  }

  async function newChat() {
    const button = findByLabel(CLICKABLE.newChat);
    if (!button) return { ok: false, error: "The pane has no New chat button." };
    safeClick(button, CLICKABLE.newChat);
    return { ok: true };
  }

  async function sendPrompt(text, requestId) {
    if (!requestId || armedRequest !== requestId) return { ok: false, error: "Request was cancelled before typing." };
    const pending = diagnostics.approval(document, visible);
    if (pending) {
      relay({ type: "approval_required", ...pending });
      return { ok: false, error: "A permission dialog needs user input. Nothing was approved." };
    }
    const box = findByLabel(PROMPT_LABELS, { editable: true });
    if (!box) return { ok: false, error: "The Copilot prompt box is not open." };
    if ((box.innerText || "").trim()) return { ok: false, error: "The prompt box has an unsent draft. Clear it first." };
    if (findByPattern(STOP_LABEL)) return { ok: false, error: "Copilot is already running in this pane." };
    box.focus();
    if (!document.execCommand("insertText", false, text)) return { ok: false, error: "Could not type the prompt." };
    const send = await waitFor(() => {
      const el = findByLabel(CLICKABLE.send);
      return el && enabled(el) ? el : null;
    }, 5000);
    if (!send) return { ok: false, error: "The Send button did not become active." };
    if (armedRequest !== requestId) return { ok: false, error: "Request was cancelled before sending. The prompt remains unsent." };
    const approval = diagnostics.approval(document, visible);
    if (approval) {
      relay({ type: "approval_required", ...approval });
      return { ok: false, error: "Permission dialog appeared before sending. Nothing was approved." };
    }
    safeClick(send, CLICKABLE.send);
    return { ok: true };
  }

  function readPane() {
    const text = (document.body && document.body.innerText) || "";
    return { ok: true, text: text.slice(-4000) };
  }

  async function run(cmd) {
    switch (cmd.action) {
      case "describe": return describe();
      case "inspect": return inspect();
      case "open_pane": return openPane();
      case "new_chat": return newChat();
      case "send_prompt": return sendPrompt(String(cmd.text || ""), cmd.requestId);
      case "read_pane": return readPane();
      case "sending":
        // Tells the hook that Send is about to be clicked for the armed request.
        if (!cmd.requestId || armedRequest !== cmd.requestId) return { ok: false, error: "The frame is not armed for this request." };
        toHook("sending");
        await sleep(50);
        return { ok: true };
      case "arm":
        armedRequest = cmd.on ? cmd.requestId : null;
        watchApprovals();
        toHook("arm", { on: Boolean(cmd.on), prompt: cmd.prompt || "" });
        await sleep(100);
        return { ok: true, hook: { ...hook } };
      default: return { ok: false, error: `unknown action ${cmd.action}` };
    }
  }

  const alive = () => {
    try { return Boolean(chrome.runtime && chrome.runtime.id); } catch (e) { return false; }
  };
  const existing = globalThis.__ocliController;
  if (existing && existing.version === VERSION && existing.alive()) return;
  const self = { version: VERSION, alive, describe, run };
  globalThis.__ocliController = self;

  window.addEventListener("message", (ev) => {
    const msg = ev.data;
    if (globalThis.__ocliController !== self) return;
    if (msg && msg[CHANNEL] === "frame_probe" && ev.source === window.parent && window !== window.parent) {
      window.parent.postMessage({ [CHANNEL]: "frame_identity", url: location.origin + location.pathname }, ev.origin);
      return;
    }
    if (msg && msg[CHANNEL] === "frame_identity") {
      let url;
      try { url = new URL(msg.url); } catch (e) { return; }
      if (url.origin !== ev.origin || !/\.resources\.office\.net$|\.officeapps\.live\.com$/.test(url.hostname)) return;
      for (const el of allElements(document)) {
        if (el.tagName === "IFRAME" && el.contentWindow === ev.source) childUrls.set(el, url.origin + url.pathname);
      }
      return;
    }
    if (ev.source !== window || !msg || msg[CHANNEL] !== "hook") return;
    if (msg.kind === "pong") {
      Object.assign(hook, { seen: true, armed: msg.armed, sockets: msg.sockets,
        revision: msg.revision || null, schemas: msg.schemas || [], request: msg.request || null });
    } else if (msg.kind === "event" && armedRequest && msg.event && typeof msg.event.type === "string") {
      relay(msg.event);
    }
  });

  if (typeof module !== "undefined" && module.exports) {
    module.exports = { frameRole, findByLabel, findByPattern, safeClick, CLICKABLE, PROMPT_LABELS, FORBIDDEN, STOP_LABEL };
  }
})();
