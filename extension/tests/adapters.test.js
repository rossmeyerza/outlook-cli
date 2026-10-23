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

test("PowerPoint ignores document and edit payloads", () => {
  const raw = wrap(tagged(PP.output + "V1", { pptEditsPayload: "PRIVATE", basePresentation: "PRIVATE" }));
  assert.deepEqual(ppt.adapt("recv", raw, {}), []);
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
  assert.deepEqual(ppt.adapt("recv", final, state), [{ type: "completed", text: "Hello world", conversationId: "c" }]);
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
