const test = require("node:test");
const assert = require("node:assert/strict");
const common = require("../adapters/common.js");
const excel = require("../adapters/excel.js");
const ppt = require("../adapters/powerpoint.js");

const tagged = (type, extra = {}) => ({ H_: { T_: type }, ...extra });
const wrap = (body) => ({ ops: [{ items: [{ body }] }] });
const EX = excel.TYPES;
const PP = ppt.TYPES;

test("frame parsing: multiple records, keepalive and non-JSON", () => {
  assert.deepEqual(common.parseFrame('~\x1e{"a":1}\x1e{"b":2}'), [{ a: 1 }, { b: 2 }]);
  assert.deepEqual(common.parseFrame(new ArrayBuffer(5)), []);
  assert.deepEqual(common.parseFrame("{invalid"), []);
});

test("Excel signal: ids only, no document context or prompt", () => {
  const input = wrap(tagged(EX.signal, { requestId: "r", conversationId: "c", content: { query: "PRIVATE" }, model: "PRIVATE" }));
  assert.deepEqual(excel.adapt("send", input), [{ type: "started", app: "excel", requestId: "r", conversationId: "c" }]);
});

test("Excel assistant chunks and final full text are distinct", () => {
  const body = tagged(EX.annotation, { requestId: "r", content: tagged(EX.message, { text: "Hello", status: "working" }) });
  assert.deepEqual(excel.adapt("recv", wrap(body)), [{ type: "text", text: "Hello" }]);
  body.content.status = "completed";
  body.content.text = "Hello world";
  assert.deepEqual(excel.adapt("recv", wrap(body)), [{ type: "completed", text: "Hello world", conversationId: undefined }]);
});

test("Excel progress exports the header, not chain-of-thought", () => {
  const body = tagged(EX.annotation, { content: tagged(EX.progress, { text: JSON.stringify({ headerText: "Working", text: "PRIVATE" }) }) });
  assert.deepEqual(excel.adapt("recv", wrap(body)), [{ type: "progress", text: "Working" }]);
});

test("Excel script summary exports no code or cell values", () => {
  const body = tagged(EX.invoke, { namespace: "excel", functionName: "executeScript", argumentsJson: JSON.stringify({ script: 'sheet.getRange("A1:C4").values = [["PRIVATE"]]' }) });
  const events = excel.adapt("recv", wrap(body));
  assert.equal(events[0].name, "executeScript");
  assert.deepEqual(events[0].ranges, ["A1:C4"]);
  assert.equal(events[0].writes, true);
  assert(!JSON.stringify(events).includes("PRIVATE"));
  assert.equal(excel.summariseScript('s.getRange("A1").load("values")').writes, false);
});

test("Excel pending action returns approval and never an action", () => {
  const events = excel.adapt("recv", wrap(tagged(EX.pendingAction, { text: "Approve?", script: "PRIVATE" })));
  assert.deepEqual(events, [{ type: "approval_required", app: "excel", kind: "pending_action", text: "Approve?" }]);
});

test("Excel invocation result exports only success", () => {
  const body = tagged("AugLoop_Session_Protocol_ExecutionCorrelatedClientResponse", {
    content: tagged(EX.invokeResult, { items: [{ content: '{"scriptResponse":{"resultType":"success","result":"PRIVATE"}}' }] }),
  });
  assert.deepEqual(excel.adapt("send", body), [{ type: "action_result", app: "excel", ok: true }]);
});

test("PowerPoint edit batches export a name and status, not the payload", () => {
  const raw = wrap(tagged(PP.outputV1, { pptEditsPayload: "PRIVATE", basePresentation: "PRIVATE", responseStatus: "complete" }));
  const events = ppt.adapt("recv", raw, {});
  assert.deepEqual(events, [
    { type: "status", source: "output_v1", value: "complete" },
    { type: "action", app: "powerpoint", name: "ppt_edits", status: "complete" },
  ]);
  assert(!JSON.stringify(events).includes("PRIVATE"));
  assert.deepEqual(ppt.adapt("recv", wrap(tagged(PP.output, { basePresentation: "PRIVATE" })), {}), []);
  assert.deepEqual(ppt.adapt("send", wrap(tagged("AugLoop_Powerpoint_PowerPointShape", { text: "PRIVATE" })), {}), []);
});

test("PowerPoint response chunks decode partial JSON strings", () => {
  const state = {};
  const chunk = (text) => wrap(tagged(PP.output, { chunkType: "response", chunkContent: text }));
  assert.deepEqual(ppt.adapt("recv", chunk('"Hello'), state), [{ type: "text", text: "Hello" }]);
  assert.deepEqual(ppt.adapt("recv", chunk(' world"'), state), [{ type: "text", text: " world" }]);
});

test("PowerPoint plain chunks and full response", () => {
  const state = {};
  assert.equal(ppt.responseDelta(state, "Hello"), "Hello");
  assert.equal(ppt.responseDelta(state, " world"), " world");
  const final = wrap(tagged(PP.output, { responseStatus: "complete", queryId: "r", sydneyConversationId: "c", finalResponse: '{"assistantResponse":"Hello world","thoughts":"PRIVATE"}', basePresentation: "PRIVATE" }));
  assert.deepEqual(ppt.adapt("recv", final, state), [
    { type: "status", source: "output", value: "complete" },
    { type: "completed", text: "Hello world", conversationId: "c" },
  ]);
});

test("PowerPoint slide activity exports only target ids", () => {
  const raw = wrap(tagged(PP.output, { chunkType: "slide_activity", chunkContent: JSON.stringify({ phase: "completed", terminalStatus: "success", targets: [{ kind: "insertion", slideId: 5, index: 1, title: "PRIVATE" }] }) }));
  const events = ppt.adapt("recv", raw, {});
  assert.deepEqual(events, [{ type: "action", app: "powerpoint", name: "slide_activity", status: "success", targets: [{ kind: "insertion", slideId: 5, index: 1 }] }]);
});

test("Word actions contain only names and no arguments or Excel write claims", () => {
  const raw = wrap(tagged(EX.invoke, { namespace: "word", functionName: "readDocument", argumentsJson: '{"text":"PRIVATE","script":"PRIVATE"}' }));
  const events = excel.adapt("recv", raw, "word");
  assert.equal(events[0].app, "word");
  assert.equal(events[0].namespace, "word");
  assert.equal(events[0].name, "readDocument");
  assert(!JSON.stringify(events).includes("PRIVATE"));
  assert.equal(events[0].writes, undefined);
  assert.equal(events[0].ranges, undefined);
});

test("PowerPoint action responses export no document data", () => {
  const response = tagged(PP.correlated, { content: tagged(PP.actionResponse, { actionResponse: { isSuccess: true, statusCode: "CompletedStatusCode", errorDetails: "PRIVATE" } }) });
  assert.deepEqual(ppt.adapt("send", response, {}), [{ type: "action_result", app: "powerpoint", ok: true, status: "CompletedStatusCode" }]);
});

test("PowerPoint completion: final answer object, status names and failures", () => {
  const out = (extra) => wrap(tagged(PP.output, { queryId: "r", ...extra }));
  const types = (events) => events.map((e) => e.type);
  // Unknown status name, but a final answer object: the turn is over.
  assert.deepEqual(types(ppt.adapt("recv", out({ responseStatus: "Finalised", finalResponse: '{"assistantResponse":"A"}' }), {})), ["status", "completed"]);
  // No status at all, final answer present.
  assert.deepEqual(types(ppt.adapt("recv", out({ finalResponse: '{"assistantResponse":"A"}' }), {})), ["completed"]);
  // Still running: a partial final object is not the end.
  assert.deepEqual(types(ppt.adapt("recv", out({ responseStatus: "inProgress", finalResponse: '{"assistantResponse":"A"}' }), {})), ["status"]);
  // Other completion names.
  assert.deepEqual(types(ppt.adapt("recv", out({ responseStatus: "Completed", finalResponse: "" }), {})), ["status", "completed"]);
  assert.deepEqual(types(ppt.adapt("recv", out({ responseStatus: "failed" }), {})), ["status", "error"]);
  // A status is exported once per change, and long or odd values never.
  const state = {};
  ppt.adapt("recv", out({ responseStatus: "streaming" }), state);
  assert.deepEqual(ppt.adapt("recv", out({ responseStatus: "streaming" }), state), []);
  assert.deepEqual(ppt.adapt("recv", out({ responseStatus: "PRIVATE text with: punctuation!" }), {}), []);
});

test("prompt normal form ignores editor changes, not wording", () => {
  const a = common.normalisePrompt('Edit slide 6.\r\n\r\n- One  "two"\n1. Three');
  assert.equal(a, 'Edit slide 6. One "two" Three');
  assert.equal(common.normalisePrompt("Edit slide 6.\n\u2022 One\u00a0 \u201ctwo\u201d\n\u2022 Three"), a);
  assert.notEqual(common.normalisePrompt("Edit slide 7. One two Three"), a);
  assert.equal(common.normalisePrompt(undefined), "");
});

test("task status exports a status name only", () => {
  const record = tagged(common.TASK_INFO, { ownerId: "PRIVATE", taskStatus: "Running" });
  assert.deepEqual(common.taskEvents("recv", record), [{ type: "task_status", value: "Running" }]);
  assert.deepEqual(common.taskEvents("send", record), []);
  assert.deepEqual(common.taskEvents("recv", tagged(common.TASK_INFO, { taskStatus: { PRIVATE: 1 } })), [{ type: "task_status", value: "object" }]);
});
