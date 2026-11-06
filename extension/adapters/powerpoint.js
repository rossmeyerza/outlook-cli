// PowerPoint Copilot (PptAgentExperimental) adapter: AugLoop records in, bridge events out.
// PowerPoint sends the whole document model over AugLoop. Only names, ids and
// answer text leave this adapter, never slide content or edit payloads.
(function (root) {
  "use strict";

  const common = root.OcliAdapters ? root.OcliAdapters.common : require("./common.js");
  const { typeOf, bodies, parseJson, cleanHeader, decodePartialJsonString, token } = common;

  const T = {
    signal: "AugLoop_PptAgentExperimental_PptAgentExperimentalSignal",
    output: "AugLoop_PptAgentExperimental_PptAgentExperimentalOutputAnnotation",
    // Seen 2026-11-06: one per edit batch, each followed by one client action response.
    outputV1: "AugLoop_PptAgentExperimental_PptAgentExperimentalOutputAnnotationV1",
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

  const DONE = /^(complete|completed|success|succeeded|done|finished)$/i;
  const FAILED = /^(error|failed|failure|cancelled|canceled|aborted|timeout|timedout)$/i;
  const RUNNING = /^(in[ _-]?progress|streaming|running|pending|started|partial|thinking|processing)$/i;

  // Status names are exported once per change, so a live check shows the real values.
  function noteStatus(state, source, value, events) {
    const status = token(value);
    if (!status || state[source] === status) return status;
    state[source] = status;
    events.push({ type: "status", source, value: status });
    return status;
  }

  function fromOutput(body, state, events) {
    if (body.chunkType !== undefined) fromChunk(body, state, events);
    const status = noteStatus(state, "output", body.responseStatus, events);
    const final = parseJson(body.finalResponse);
    const finalText = final && typeof final.assistantResponse === "string" ? final.assistantResponse : null;
    const conversationId = conversationIdOf(body);
    if (status && FAILED.test(status)) {
      events.push({ type: "error", message: `PowerPoint Copilot reported ${status}.`, conversationId });
    } else if ((status && DONE.test(status)) || (finalText !== null && !(status && RUNNING.test(status)))) {
      // A final answer object is the end of the turn, whatever the status name is.
      events.push({ type: "completed", text: finalText !== null ? finalText : (state.emitted || ""), conversationId });
    }
  }

  // Edit batches. Export that a batch arrived, never its payload.
  function fromOutputV1(body, state, events) {
    const status = noteStatus(state, "output_v1", body.responseStatus, events);
    if (body.pptEditsPayload) events.push({ type: "action", app: "powerpoint", name: "ppt_edits", status });
  }

  function adapt(direction, record, state) {
    state = state || {};
    const events = [];
    for (const body of bodies(record)) {
      const kind = typeOf(body);
      if (direction === "send" && kind === T.signal) {
        state.raw = "";
        state.emitted = "";
        state.output = null;
        state.output_v1 = null;
        events.push({ type: "started", app: "powerpoint", requestId: body.queryId, conversationId: conversationIdOf(body) });
      } else if (direction === "recv" && kind === T.output) {
        fromOutput(body, state, events);
      } else if (direction === "recv" && kind === T.outputV1) {
        fromOutputV1(body, state, events);
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
