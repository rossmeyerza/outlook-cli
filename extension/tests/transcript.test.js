const assert = require("node:assert/strict");
const { readFileSync } = require("node:fs");
const { join } = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const source = readFileSync(join(__dirname, "../transcript/capture.js"), "utf8");
const context = { URL };

vm.createContext(context);
vm.runInContext(
  `${source}\nthis.__testExports = { createExport, formatVttTime, escapeVtt, isTeamsUrl };`,
  context,
);

const { createExport, escapeVtt, formatVttTime, isTeamsUrl } = context.__testExports;
const sample = {
  ok: true,
  entries: [
    {
      timestamp: "0:03",
      seconds: 3,
      speaker: "A & B",
      text: "One < two",
    },
    {
      timestamp: "0:05",
      seconds: 5,
      speaker: "C",
      text: "Done.",
    },
  ],
};

test("plain-text export contains timestamped entries", () => {
  const output = createExport(sample, "txt");
  assert.equal(output.mime, "text/plain");
  assert.equal(
    output.content,
    "[0:03] A & B: One < two\n\n[0:05] C: Done.\n",
  );
});

test("JSON export preserves the transcript data", () => {
  const output = createExport(sample, "json");
  assert.equal(output.mime, "application/json");
  assert.deepEqual(JSON.parse(output.content), sample);
});

test("WebVTT export uses cue times and escapes markup", () => {
  const output = createExport(sample, "vtt");
  assert.equal(output.mime, "text/vtt");
  assert.match(output.content, /00:00:03\.000 --> 00:00:05\.000/);
  assert.match(output.content, /<v A &amp; B>One &lt; two/);
});

test("time and VTT escaping helpers are stable", () => {
  assert.equal(formatVttTime(3661.25), "01:01:01.250");
  assert.equal(escapeVtt("<&>"), "&lt;&amp;&gt;");
});

test("both Teams hosts are accepted, but lookalike hosts are rejected", () => {
  assert.equal(isTeamsUrl("https://teams.microsoft.com/v2/"), true);
  assert.equal(isTeamsUrl("https://teams.cloud.microsoft/"), true);
  assert.equal(isTeamsUrl("https://teams.cloud.microsoft.evil.test/"), false);
  assert.equal(isTeamsUrl("http://teams.cloud.microsoft/"), false);
  assert.equal(isTeamsUrl("not a URL"), false);
});
