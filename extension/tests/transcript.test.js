const assert = require("node:assert/strict");
const { readFileSync } = require("node:fs");
const { join } = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const source = readFileSync(join(__dirname, "../transcript/capture.js"), "utf8");
const context = { URL };

vm.createContext(context);
vm.runInContext(
  `${source}\nthis.__testExports = { captureActiveTab, createExport, formatVttTime, escapeVtt, isTeamsUrl };`,
  context,
);

const { captureActiveTab, createExport, escapeVtt, formatVttTime, isTeamsUrl } = context.__testExports;
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

function stubCapture(download) {
  context.chrome = {
    tabs: { query: async () => [{ id: 42, url: "https://teams.microsoft.com/v2/" }] },
    scripting: { executeScript: async () => [{ result: sample }] },
    downloads: { download },
  };
}

test("copy returns the selected export format and never requests a download", async () => {
  stubCapture(() => assert.fail("Copy must not request a download"));
  for (const format of ["txt", "vtt", "json"]) {
    const response = await captureActiveTab(format, "copy");
    assert.equal(response.ok, true);
    assert.equal(response.count, sample.entries.length);
    assert.equal(response.content, createExport(sample, format).content);
    assert.equal(response.filename, undefined);
  }
});

test("download is still the default and requests a Save As dialog", async () => {
  for (const destination of [undefined, "download"]) {
    const downloads = [];
    stubCapture(async (request) => { downloads.push(request); });
    const response = await captureActiveTab("txt", destination);
    assert.equal(response.ok, true);
    assert.equal(downloads.length, 1);
    const [download] = downloads;
    assert.equal(download.saveAs, true);
    assert.equal(response.filename, download.filename);
    assert.match(download.filename, /\.txt$/);
    assert.equal(
      download.url,
      `data:text/plain;charset=utf-8,${encodeURIComponent(createExport(sample, "txt").content)}`,
    );
  }
});

const popupSource = readFileSync(join(__dirname, "../transcript/popup.js"), "utf8");

function loadPopup(sendMessage, writeText) {
  const elements = Object.fromEntries(
    ["capture", "copy", "format", "status", "diagnostics", "diagnostics-output"].map((id) => [id, {
      dataset: {},
      addEventListener(type, listener) { this[type] = listener; },
    }]),
  );
  elements.format.value = "vtt";
  vm.runInNewContext(popupSource, {
    document: { querySelector: (selector) => elements[selector.slice(1)] },
    chrome: { runtime: { sendMessage } },
    navigator: { clipboard: { writeText } },
  });
  return elements;
}

test("copy button writes the formatted transcript and confirms success", async () => {
  const content = createExport(sample, "vtt").content;
  const written = [];
  const popup = loadPopup(async (message) => {
    assert.equal(message.type, "CAPTURE_TEAMS_TRANSCRIPT");
    assert.equal(message.format, "vtt");
    assert.equal(message.destination, "copy");
    return { ok: true, count: 2, content };
  }, async (value) => { written.push(value); });

  await popup.copy.click();
  assert.deepEqual(written, [content]);
  assert.equal(popup.status.dataset.state, "success");
  assert.equal(popup.copy.disabled, false);
});

test("clipboard failure shows an error and lets the user retry or download", async () => {
  const popup = loadPopup(
    async () => ({ ok: true, count: 2, content: "Transcript" }),
    async () => { throw new Error("Clipboard unavailable"); },
  );

  await popup.copy.click();
  assert.equal(popup.status.dataset.state, "error");
  assert.match(popup.status.textContent, /Could not copy/);
  assert.equal(popup.copy.disabled, false);
  assert.equal(popup.capture.disabled, false);
  assert.equal(popup.format.disabled, false);
});
