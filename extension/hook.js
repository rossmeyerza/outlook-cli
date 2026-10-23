// Main world, Office editor frames. Watches AugLoop WebSocket frames and passes
// them to the app adapters. Only adapter events leave this script, and only
// while the controller has armed it for a request. Raw frames never leave.
(() => {
  "use strict";
  if (window.__ocliCopilotHook) return;

  const CHANNEL = "__ocliCopilot";
  const AUGLOOP_HOST = /(^|\.)(augloop\.office\.com|augloop\.svc\.cloud\.microsoft)$/i;
  const RELEVANT = /OfficeAgent|PptAgent|ExecutionCorrelated/;
  const adapters = window.OcliAdapters || {};
  const watched = new WeakSet();
  const pptState = {};
  let armed = false;
  let agentRequestId = null;
  let expectedPrompt = "";
  let sockets = 0;
  const schemas = new Map();

  function observeSchemas(direction, record) {
    const bodies = adapters.common.bodies(record);
    const content = bodies.map((body) => body.content).filter((value) => value && typeof value === "object" && !Array.isArray(value));
    for (const body of bodies.concat(content)) {
      const name = adapters.common.typeOf(body) || body.annotationType;
      if (typeof name !== "string" || !/^AugLoop_[A-Za-z0-9_]{1,160}$/.test(name)) continue;
      const key = direction + ":" + name;
      if (!schemas.has(key) && schemas.size >= 128) continue;
      const sample = schemas.get(key) || { name, direction, count: 0, keys: [] };
      sample.count += 1;
      sample.keys = Object.keys(body).filter((k) => /^[A-Za-z_][A-Za-z0-9_]{0,60}$/.test(k)).sort().slice(0, 30);
      schemas.set(key, sample);
    }
  }

  const post = (payload) => {
    try {
      window.postMessage({ [CHANNEL]: "hook", ...payload }, location.origin);
    } catch (e) {
      // The page can block postMessage. The controller then reports no events.
    }
  };

  function handle(direction, data) {
    if (typeof data !== "string" || data.length > 8 * 1024 * 1024 || !data.includes("AugLoop_")) return;
    for (const raw of adapters.common.parseFrame(data)) {
      // Schema names and property names only. Never retain values or raw frames.
      observeSchemas(direction, raw);
      if (!armed || !RELEVANT.test(data)) continue;
      const record = filterRecord(direction, raw);
      if (!record) continue;
      let events = [];
      try {
        events = events.concat(adapters.excel.adapt(direction, record,
          /wordeditorframe\.aspx/i.test(location.pathname) ? "word" : "excel"));
        events = events.concat(adapters.powerpoint.adapt(direction, record, pptState));
      } catch (e) {
        events.push({ type: "note", message: "adapter failed: " + String(e && e.message) });
      }
      for (const event of events) post({ kind: "event", event });
    }
  }

  function filterRecord(direction, record) {
    const signals = adapters.common.bodies(record).filter((b) =>
      [adapters.excel.TYPES.signal, adapters.powerpoint.TYPES.signal].includes(adapters.common.typeOf(b)));
    if (direction === "send" && signals.length) {
      // OfficeAgent also sends non-prompt signals while the pane starts. Only
      // bind to the exact prompt this request placed in the input control.
      const signal = signals.find((b) => {
        const query = b.query || (b.content && b.content.query);
        return typeof query === "string" && query.replace(/\r\n/g, "\n").trim() === expectedPrompt;
      });
      if (!signal) return null;
      const id = signal.requestId || signal.queryId;
      if (agentRequestId && agentRequestId !== id) {
        post({ kind: "event", event: { type: "error", message: "Another prompt was sent in this pane. Do not retry automatically." } });
        return null;
      }
      agentRequestId = id;
    }
    if (!agentRequestId) return null;
    if (!record.ops) return record;
    return { ...record, ops: record.ops.map((op) => ({ ...op, items: (op.items || []).filter((item) => {
      const body = item.body || {};
      const id = body.requestId || body.queryId;
      return !id || id === agentRequestId;
    }) })) };
  }

  function watch(ws) {
    if (watched.has(ws)) return;
    let url = "";
    try { url = String(ws.url || ""); } catch (e) { return; }
    const parsed = new URL(url);
    if (parsed.protocol !== "wss:" || !AUGLOOP_HOST.test(parsed.hostname)) return;
    watched.add(ws);
    sockets += 1;
    ws.addEventListener("message", (ev) => handle("recv", ev.data));
  }

  const NativeWebSocket = window.WebSocket;
  const nativeSend = NativeWebSocket.prototype.send;

  // A socket opened before this script ran is picked up at its next send.
  NativeWebSocket.prototype.send = function (data) {
    try {
      watch(this);
      if (watched.has(this)) handle("send", data);
    } catch (e) {
      // Never break the editor.
    }
    return nativeSend.apply(this, arguments);
  };

  function HookedWebSocket(url, protocols) {
    const ws = protocols === undefined ? new NativeWebSocket(url) : new NativeWebSocket(url, protocols);
    try { watch(ws); } catch (e) { /* Never break the editor. */ }
    return ws;
  }
  HookedWebSocket.prototype = NativeWebSocket.prototype;
  for (const key of ["CONNECTING", "OPEN", "CLOSING", "CLOSED"]) HookedWebSocket[key] = NativeWebSocket[key];
  window.WebSocket = HookedWebSocket;

  window.addEventListener("message", (ev) => {
    const msg = ev.data;
    if (ev.source !== window || !msg || msg[CHANNEL] !== "controller") return;
    if (msg.kind === "arm") {
      armed = msg.on === true;
      agentRequestId = null;
      expectedPrompt = typeof msg.prompt === "string" ? msg.prompt.replace(/\r\n/g, "\n").trim() : "";
      pptState.raw = "";
      pptState.emitted = "";
    }
    if (msg.kind === "arm" || msg.kind === "ping") post({ kind: "pong", armed, sockets,
      revision: 3, schemas: [...schemas.values()] });
  });

  window.__ocliCopilotHook = { version: 3 };
  post({ kind: "pong", armed, sockets, revision: 3, schemas: [] });
})();
