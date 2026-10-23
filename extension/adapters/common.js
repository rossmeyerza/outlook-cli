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

  const api = { parseFrame, typeOf, bodies, parseJson, cleanHeader, decodePartialJsonString };
  root.OcliAdapters = Object.assign(root.OcliAdapters || {}, { common: api });
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
