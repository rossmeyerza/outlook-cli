// PowerPoint Copilot (PptAgentExperimental) adapter: AugLoop records in, bridge events out.
// PowerPoint sends the whole document model over AugLoop. Only names, ids and
// answer text leave this adapter, never slide content or edit payloads.
(function (root) {
  "use strict";

  const common = root.OcliAdapters ? root.OcliAdapters.common : require("./common.js");
  const { typeOf, bodies, parseJson, cleanHeader, decodePartialJsonString } = common;

  const T = {
    signal: "AugLoop_PptAgentExperimental_PptAgentExperimentalSignal",
    output: "AugLoop_PptAgentExperimental_PptAgentExperimentalOutputAnnotation",
    actionResponse: "AugLoop_PptAgentExperimental_CopilotActionResponseObject",
    correlated: "AugLoop_Session_Protocol_ExecutionCorrelatedClientResponse",
  };

  function conversationIdOf(body) {
    if (body.sydneyConversationId) return body.sydneyConversationId;
    const conversation = parseJson(body.conversation);
    return (conversation && conversation.conversationId) || undefined;
  }

  // Response chunks can be pieces of a JSON string literal. Decode what we have
  // so far and return only the new text.
  function responseDelta(state, chunk) {
    const previous = state.raw || "";
    state.raw = chunk.startsWith(previous) && previous ? chunk : previous + chunk;
    const decoded = decodePartialJsonString(state.raw);
    const emitted = state.emitted || "";
    state.emitted = decoded;
    return decoded.startsWith(emitted) ? decoded.slice(emitted.length) : "";
  }

  function fromChunk(body, state, events) {
    const chunk = typeof body.chunkContent === "string" ? body.chunkContent : "";
    switch (body.chunkType) {
      case "thoughts_hardcoded":
      case "thoughts_generated": {
        const text = cleanHeader(chunk);
        if (text) events.push({ type: "progress", text });
        break;
      }
      case "response": {
        const text = responseDelta(state, chunk);
        if (text) events.push({ type: "text", text });
        break;
      }
      case "slide_activity": {
        const info = parseJson(chunk);
        if (info && info.phase === "completed") {
          const targets = (info.targets || []).map((t) => ({ kind: t.kind, slideId: t.slideId, index: t.index }));
          events.push({ type: "action", app: "powerpoint", name: "slide_activity", status: info.terminalStatus, targets });
        }
        break;
      }
      case "slide_update":
      case undefined:
        break;
      default:
        // Names only, for the approval work in phase 4. No content.
        events.push({ type: "note", app: "powerpoint", chunkType: String(body.chunkType) });
    }
  }

  function fromOutput(body, state, events) {
    if (body.chunkType !== undefined) fromChunk(body, state, events);
    if (body.responseStatus === "complete") {
      const final = parseJson(body.finalResponse) || {};
      const text = typeof final.assistantResponse === "string" ? final.assistantResponse : (state.emitted || "");
      events.push({ type: "completed", text, conversationId: conversationIdOf(body) });
    } else if (body.responseStatus === "error" || body.responseStatus === "failed") {
      events.push({ type: "error", message: "PowerPoint Copilot reported an error.", conversationId: conversationIdOf(body) });
    }
  }

  function adapt(direction, record, state) {
    state = state || {};
    const events = [];
    for (const body of bodies(record)) {
      const kind = typeOf(body);
      if (direction === "send" && kind === T.signal) {
        state.raw = "";
        state.emitted = "";
        events.push({ type: "started", app: "powerpoint", requestId: body.queryId, conversationId: conversationIdOf(body) });
      } else if (direction === "recv" && kind === T.output) {
        fromOutput(body, state, events);
      } else if (direction === "send" && kind === T.correlated && typeOf(body.content) === T.actionResponse) {
        const response = body.content.actionResponse || {};
        events.push({ type: "action_result", app: "powerpoint", ok: response.isSuccess === true, status: response.statusCode });
      }
    }
    return events;
  }

  const api = { adapt, responseDelta, TYPES: T };
  root.OcliAdapters = Object.assign(root.OcliAdapters || {}, { powerpoint: api });
  if (typeof module !== "undefined" && module.exports) module.exports = api;
})(typeof globalThis !== "undefined" ? globalThis : this);
