// Shared Excel/Word Copilot (OfficeAgent) adapter. Word chat was observed using
// these same schemas. App-specific action payloads never leave the page.
(function (root) {
  "use strict";

  const common = root.OcliAdapters ? root.OcliAdapters.common : require("./common.js");
  const { typeOf, bodies, parseJson } = common;

  const T = {
    signal: "AugLoop_OfficeAgent_OfficeAgentSignal",
    annotation: "AugLoop_OfficeAgent_OfficeAgentAnnotation",
    progress: "AugLoop_OfficeAgent_OfficeAgentProgress",
    message: "AugLoop_OfficeAgent_OfficeAgentAssistantMessage",
    invoke: "AugLoop_OfficeAgent_OfficeAgentInvokeFunctionAnnotation",
    invokeResult: "AugLoop_OfficeAgent_OfficeAgentInvokeFunctionResultResponse",
    // Not seen in a recording yet. The names come from the activated annotation list.
    pendingAction: "AugLoop_OfficeAgent_OfficeAgentPendingActionAnnotation",
  };

  const WRITE_PATTERN = /\.(values|formulas|numberFormat)\s*=|\.format\.|\.(insert|delete|clear|add)\s*\(|\.name\s*=/;

  // A short, content-free summary of an executeScript call.
  function summariseScript(script) {
    const ranges = [];
    const re = /getRange\(\s*["'`]([^"'`]{1,40})["'`]\s*\)/g;
    let m;
    while ((m = re.exec(script)) && ranges.length < 10) {
      if (!ranges.includes(m[1])) ranges.push(m[1]);
    }
    return { scriptLength: script.length, ranges, writes: WRITE_PATTERN.test(script) };
  }

  function fromAnnotation(body, events) {
    const content = body.content || {};
    const kind = typeOf(content);
    const conversationId = body.conversationId || undefined;
    if (kind === T.progress) {
      const info = parseJson(content.text);
      if (info && info.headerText) events.push({ type: "progress", text: String(info.headerText) });
    } else if (kind === T.message) {
      const text = typeof content.text === "string" ? content.text : "";
      if (content.status === "completed") {
        events.push({ type: "completed", text, conversationId });
      } else if (content.status === "failed" || content.status === "error") {
        events.push({ type: "error", message: text || "Office Copilot reported an error.", conversationId });
      } else if (text) {
        events.push({ type: "text", text });
      }
    }
  }

  function fromInvoke(body, events) {
    const args = parseJson(body.argumentsJson) || {};
    const action = { type: "action", app: "excel", namespace: body.namespace, name: body.functionName };
    if (typeof args.script === "string") Object.assign(action, summariseScript(args.script));
    events.push(action);
  }

  function fromInvokeResult(record, events) {
    const items = (record.content && record.content.items) || [];
    for (const item of items) {
      const payload = parseJson(item.content);
      const resultType = payload && payload.scriptResponse && payload.scriptResponse.resultType;
      if (resultType) events.push({ type: "action_result", app: "excel", ok: resultType === "success" });
    }
  }

  function adapt(direction, record, app = "excel") {
    if (!["excel", "word"].includes(app)) return [];
    const events = [];
    for (const body of bodies(record)) {
      const kind = typeOf(body);
      if (direction === "send" && kind === T.signal) {
        events.push({ type: "started", app: "excel", requestId: body.requestId, conversationId: body.conversationId || undefined });
      } else if (direction === "recv" && kind === T.annotation) {
        fromAnnotation(body, events);
      } else if (direction === "recv" && kind === T.invoke) {
        fromInvoke(body, events);
      } else if (direction === "recv" && kind === T.pendingAction) {
        events.push({ type: "approval_required", app: "excel", kind: "pending_action", text: String(body.text || body.title || "") });
      } else if (direction === "send" && kind === "AugLoop_Session_Protocol_ExecutionCorrelatedClientResponse"
        && typeOf(body.content) === T.invokeResult) {
        fromInvokeResult(body, events);
      }
    }
    if (app === "word") {
      for (const event of events) {
        if (event.app) event.app = "word";
        // Excel range/write heuristics have no meaning for Word actions.
        delete event.ranges;
        delete event.writes;
      }
    }
    return events;
  }

  const api = { adapt, summariseScript, TYPES: T };
  root.OcliAdapters = Object.assign(root.OcliAdapters || {}, { excel: api });
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
