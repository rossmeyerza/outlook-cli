// Shared helpers for the AugLoop adapters. Pure functions: they run in the page
// (main world) and in node for tests. They never return raw frames.
(function (root) {
  "use strict";

  const RECORD_SEPARATOR = "\x1e";

  // An AugLoop WebSocket frame can hold several JSON records.
  function parseFrame(data) {
    if (typeof data !== "string" || data.length < 2) return [];
    const records = [];
    for (const part of data.split(RECORD_SEPARATOR)) {
      const text = part.trim();
      if (!text || text[0] !== "{") continue;
      try {
        records.push(JSON.parse(text));
      } catch (e) {
        // Keepalive frames and partial records are not JSON.
      }
    }
    return records;
  }

  function typeOf(obj) {
    return (obj && obj.H_ && obj.H_.T_) || "";
  }

  // Yield every annotation body in a record: ops[].items[].body, and the record itself.
  function bodies(record) {
    const out = [];
    for (const op of record.ops || []) {
      for (const item of op.items || []) {
        if (item && item.body) out.push(item.body);
      }
    }
    out.push(record);
    return out;
  }

  function parseJson(text) {
    if (typeof text !== "string") return null;
    try {
      return JSON.parse(text);
    } catch (e) {
      return null;
    }
  }

  // Thought headers look like "<|im_sep|>**Weighing the options**".
  function cleanHeader(text) {
    return String(text || "")
      .replace(/<\|im_sep\|>/g, "")
      .replace(/\*\*/g, "")
      .trim();
  }

  // Decode a JSON string literal that may still be incomplete.
  function decodePartialJsonString(raw) {
    if (!raw.startsWith('"')) return raw;
    let body = raw.endsWith('"') && raw.length > 1 && !raw.endsWith('\\"') ? raw : raw + '"';
    for (let cut = 0; cut < 8; cut++) {
      const value = parseJson(body);
      if (typeof value === "string") return value;
      // Drop a cut-off escape sequence and try again.
      body = body.slice(0, -2) + '"';
    }
    return raw.slice(1);
  }

  // The pane is a rich text editor. Typed newlines, list markers, spaces and
  // quotes can come back changed in the signal query. Compare a normal form.
  const QUOTES = { "\u2018": "'", "\u2019": "'", "\u201a": "'", "\u201b": "'",
    "\u201c": '"', "\u201d": '"', "\u201e": '"', "\u201f": '"' };
  const LIST_MARKER = /^(?:[-*+\u2022\u2023\u2043\u2219\u25aa\u25cf\u25e6]|\d{1,3}[.)])\s+/;

  function normalisePrompt(text) {
    if (typeof text !== "string") return "";
    return text
      .replace(/\r\n?/g, "\n")
      .replace(/[\u2018-\u201f]/g, (c) => QUOTES[c] || c)
      .replace(/[\u00a0\u2000-\u200a\u202f\u205f\u3000]/g, " ")
      .replace(/[\u200b-\u200d\ufeff]/g, "")
      .split("\n")
      .map((line) => line.trim().replace(LIST_MARKER, ""))
      .filter(Boolean)
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
  }

  // A short enum-like value (a status name). Anything else is not exported.
  function token(value) {
    if (typeof value === "number" && Number.isFinite(value)) return String(value);
    if (typeof value !== "string") return null;
    const text = value.trim();
    return /^[A-Za-z][A-Za-z0-9_ -]{0,39}$/.test(text) ? text : null;
  }

  const TASK_INFO = "AugLoop_Core_TaskInformationAnnotation";

  // Task state from the shared AugLoop core schema. Status names only.
  function taskEvents(direction, record) {
    if (direction !== "recv") return [];
    const events = [];
    for (const body of bodies(record)) {
      if (typeOf(body) !== TASK_INFO) continue;
      events.push({ type: "task_status", value: token(body.taskStatus) || typeof body.taskStatus });
    }
    return events;
  }

  const api = { parseFrame, typeOf, bodies, parseJson, cleanHeader, decodePartialJsonString,
    normalisePrompt, token, taskEvents, TASK_INFO };
  root.OcliAdapters = Object.assign(root.OcliAdapters || {}, { common: api });
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
