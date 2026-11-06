// Main world, Office editor frames. Watches AugLoop WebSocket frames and passes
// them to the app adapters. Only adapter events leave this script, and only
// while the controller has armed it for a request. Raw frames never leave.
(() => {
  "use strict";
  if (window.__ocliCopilotHook) return;

  const REVISION = 4;
  const CHANNEL = "__ocliCopilot";
  const AUGLOOP_HOST = /(^|\.)(augloop\.office\.com|augloop\.svc\.cloud\.microsoft)$/i;
  const RELEVANT = /OfficeAgent|PptAgent|ExecutionCorrelated|Core_TaskInformation/;
  // After the controller clicks Send, the first prompt signal is this request,
  // even if the editor changed the typed text. Only for a short time.
  const FALLBACK_WINDOW_MS = 60000;
  const adapters = window.OcliAdapters || {};
  const watched = new WeakSet();
  const pptState = {};
  let armed = false;
  let binding = null; // { id, query, mode }. The query is normalised and stays in the page.
  let related = new Set();
  let expectedRaw = "";
  let expected = "";
  let sendingAt = 0;
  let stats = null;
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
        events = events.concat(adapters.common.taskEvents(direction, record));
      } catch (e) {
        events.push({ type: "note", message: "adapter failed: " + String(e && e.message) });
      }
      for (const event of events) post({ kind: "event", event });
    }
  }

  const isSignal = (b) => [adapters.excel.TYPES.signal, adapters.powerpoint.TYPES.signal]
    .includes(adapters.common.typeOf(b));
  const signalId = (b) => b.requestId || b.queryId || null;

  function signalQuery(b) {
    const query = b.query || (b.content && b.content.query);
    return typeof query === "string" ? query : "";
  }

  // A signal that operates on a running task (resume, cancel) rather than a new prompt.
  function isOperation(b) {
    const params = b.taskOperationParams || (b.content && b.content.taskOperationParams);
    if (typeof params === "string") return params.trim() !== "";
    return Boolean(params && typeof params === "object"
      && Object.values(params).some((v) => v !== null && v !== undefined && v !== "" && v !== false));
  }

  function bind(id, query, mode) {
    binding = { id, query, mode };
    post({ kind: "event", event: { type: "bound", mode } });
  }

  // Returns an error message, or null. Only the exact prompt of this request,
  // or the first prompt after Send, binds the request.
  function classify(signal) {
    stats.signals += 1;
    const id = signalId(signal);
    const raw = signalQuery(signal);
    const query = adapters.common.normalisePrompt(raw);
    const exact = Boolean(raw) && raw.replace(/\r\n/g, "\n").trim() === expectedRaw;
    const match = exact || (Boolean(query) && query === expected);
    if (!id) return null;
    if (!binding) {
      if (match) bind(id, query, exact ? "exact" : "normalised");
      else if (query && sendingAt && Date.now() - sendingAt < FALLBACK_WINDOW_MS && !isOperation(signal)) bind(id, query, "after_send");
      else stats.ignored += 1;
      return null;
    }
    if (id === binding.id) return null;
    if (!query || isOperation(signal)) {
      related.add(id);
      stats.related += 1;
      return null;
    }
    if (match || query === binding.query) {
      stats.repeated += 1;
      return "Another prompt was sent in this pane. Do not retry automatically.";
    }
    stats.ignored += 1;
    return null;
  }

  function filterRecord(direction, record) {
    const signals = adapters.common.bodies(record).filter(isSignal);
    if (direction === "send" && signals.length) {
      for (const signal of signals) {
        const error = classify(signal);
        if (error) {
          post({ kind: "event", event: { type: "error", message: error } });
          return null;
        }
      }
    }
    if (!binding) return null;
    if (!record.ops) return isSignal(record) && signalId(record) !== binding.id ? null : record;
    return { ...record, ops: record.ops.map((op) => ({ ...op, items: (op.items || []).filter((item) => {
      const body = item.body || {};
      if (isSignal(body)) return signalId(body) === binding.id;
      const id = body.requestId || body.queryId;
      return !id || id === binding.id || related.has(id);
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
      binding = null;
      related = new Set();
      sendingAt = 0;
      stats = { signals: 0, ignored: 0, related: 0, repeated: 0 };
      expectedRaw = typeof msg.prompt === "string" ? msg.prompt.replace(/\r\n/g, "\n").trim() : "";
      expected = adapters.common.normalisePrompt(expectedRaw);
      for (const key of Object.keys(pptState)) delete pptState[key];
    }
    if (msg.kind === "sending" && armed) sendingAt = Date.now();
    if (msg.kind === "arm" || msg.kind === "ping") post({ kind: "pong", armed, sockets, revision: REVISION,
      schemas: [...schemas.values()],
      // Counts and the binding mode only. Never ids or prompt text.
      request: stats ? { bound: binding ? binding.mode : null, ...stats } : null });
  });

  window.__ocliCopilotHook = { version: REVISION };
  post({ kind: "pong", armed, sockets, revision: REVISION, schemas: [], request: null });
})();
