"use strict";

async function captureActiveTab(format, destination) {
  const [tab] = await chrome.tabs.query({ active: true, lastFocusedWindow: true });

  if (!tab?.id || !isTeamsUrl(tab.url)) {
    throw new Error("Make the Microsoft Teams transcript tab active, then try again.");
  }

  const executions = await chrome.scripting.executeScript({
    target: { tabId: tab.id, allFrames: true },
    func: captureTranscriptFromPage,
    args: [{
      settleDelayMs: 220,
      maximumDurationMs: 120000,
      origins: [new URL(tab.url).origin, "https://wppcloud-my.sharepoint.com"],
    }],
  });

  const successfulExecutions = executions
    .filter((execution) => execution?.result?.ok)
    .sort((left, right) => right.result.entries.length - left.result.entries.length);
  const result = successfulExecutions[0]?.result;

  if (!result?.ok) {
    const frameDiagnostics = executions.map((execution) => ({
      frameId: execution.frameId,
      documentId: execution.documentId,
      error: execution.result?.error,
      diagnostics: execution.result?.diagnostics,
    }));
    const mostRelevantFailure = executions.find((execution) => {
      return execution.result?.diagnostics?.transcriptWrapper;
    }) || executions.find((execution) => execution.result?.diagnostics);

    return {
      ok: false,
      error: mostRelevantFailure?.result?.error || "Teams did not return transcript entries.",
      diagnostics: { frames: frameDiagnostics },
    };
  }

  const exportFormat = ["txt", "json", "vtt"].includes(format) ? format : "txt";
  const output = createExport(result, exportFormat);
  const count = result.entries.length;

  if (destination === "copy") {
    return { ok: true, count, content: output.content };
  }

  const filename = await downloadExport(output, exportFormat);
  return { ok: true, count, filename };
}

async function downloadExport(output, exportFormat) {
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  const filename = `teams-transcript-${stamp}.${exportFormat}`;
  const url = `data:${output.mime};charset=utf-8,${encodeURIComponent(output.content)}`;

  await chrome.downloads.download({
    url,
    filename,
    saveAs: true,
  });

  return filename;
}

function isTeamsUrl(value) {
  try {
    const url = new URL(value);
    return url.protocol === "https:"
      && ["teams.microsoft.com", "teams.cloud.microsoft"].includes(url.hostname);
  } catch {
    return false;
  }
}

function createExport(result, format) {
  if (format === "json") {
    return {
      mime: "application/json",
      content: `${JSON.stringify(result, null, 2)}\n`,
    };
  }

  if (format === "vtt") {
    const cues = result.entries.map((entry, index, entries) => {
      const start = entry.seconds;
      const followingStart = entries[index + 1]?.seconds;
      const end = Number.isFinite(followingStart) && followingStart > start
        ? followingStart
        : start + 2;
      const speaker = escapeVtt(entry.speaker || "Unknown speaker");
      const text = escapeVtt(entry.text);
      return `${formatVttTime(start)} --> ${formatVttTime(end)}\n<v ${speaker}>${text}`;
    });

    return {
      mime: "text/vtt",
      content: `WEBVTT\n\n${cues.join("\n\n")}\n`,
    };
  }

  const lines = result.entries.map(
    (entry) => `[${entry.timestamp}] ${entry.speaker}: ${entry.text}`,
  );
  return { mime: "text/plain", content: `${lines.join("\n\n")}\n` };
}

function formatVttTime(totalSeconds) {
  const safeSeconds = Math.max(0, Number(totalSeconds) || 0);
  const hours = Math.floor(safeSeconds / 3600);
  const minutes = Math.floor((safeSeconds % 3600) / 60);
  const seconds = Math.floor(safeSeconds % 60);
  const milliseconds = Math.round((safeSeconds % 1) * 1000);
  return [hours, minutes, seconds]
    .map((part) => String(part).padStart(2, "0"))
    .join(":") + `.${String(milliseconds).padStart(3, "0")}`;
}

function escapeVtt(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

async function captureTranscriptFromPage(options) {
  if (!options.origins.includes(location.origin)) return null;

  const timestampPattern = /^\d{1,2}:\d{2}(?::\d{2})?$/;
  const wait = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds));
  const nextFrame = () => new Promise((resolve) => requestAnimationFrame(resolve));
  const normalise = (value) => String(value || "").replace(/\s+/g, " ").trim();

  function isVisible(element) {
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== "none"
      && style.visibility !== "hidden"
      && Number(style.opacity || 1) !== 0
      && rect.width > 0
      && rect.height > 0;
  }

  function documentRoots(startRoot = document) {
    const roots = [];
    const pending = [startRoot];
    const visited = new Set();

    while (pending.length > 0) {
      const root = pending.shift();
      if (!root || visited.has(root)) continue;
      visited.add(root);
      roots.push(root);

      for (const element of root.querySelectorAll?.("*") || []) {
        let shadowRoot = element.shadowRoot;
        if (!shadowRoot && chrome.dom?.openOrClosedShadowRoot) {
          try {
            shadowRoot = chrome.dom.openOrClosedShadowRoot(element);
          } catch {
            shadowRoot = null;
          }
        }
        if (shadowRoot) pending.push(shadowRoot);
      }
    }

    return roots;
  }

  function deepQuerySelectorAll(selector) {
    return documentRoots().flatMap((root) => [...root.querySelectorAll(selector)]);
  }

  function deepQuerySelector(selector) {
    for (const root of documentRoots()) {
      const match = root.querySelector(selector);
      if (match) return match;
    }
    return null;
  }

  function timestampTextNodes(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    const nodes = [];

    while (walker.nextNode()) {
      const node = walker.currentNode;
      const value = normalise(node.textContent);
      if (!timestampPattern.test(value) || !node.parentElement || !isVisible(node.parentElement)) {
        continue;
      }

      const range = document.createRange();
      range.selectNodeContents(node);
      const rect = range.getBoundingClientRect();
      if (rect.width > 0 && rect.height > 0) {
        nodes.push({ node, timestamp: value, rect });
      }
    }

    return nodes;
  }

  function timestampCount(root) {
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let count = 0;
    while (walker.nextNode()) {
      if (timestampPattern.test(normalise(walker.currentNode.textContent))) {
        count += 1;
      }
    }
    return count;
  }

  const transcriptWrapper = deepQuerySelector('[data-testid="transcript-list-wrapper"]');
  const transcriptTab = deepQuerySelectorAll('[role="tab"]').find((element) => {
    return element.getAttribute("aria-selected") === "true"
      && /^transcript$/i.test(normalise(element.textContent));
  });

  if (!transcriptTab && !transcriptWrapper) {
    return {
      ok: false,
      error: "This frame does not contain the selected Transcript tab or transcript list.",
      diagnostics: {
        url: `${location.origin}${location.pathname}`,
        rootCount: documentRoots().length,
        closedShadowRootApi: Boolean(chrome.dom?.openOrClosedShadowRoot),
      },
    };
  }

  function elementSummary(element) {
    if (!element) return null;
    const rect = element.getBoundingClientRect();
    const style = getComputedStyle(element);
    return {
      tag: element.tagName.toLowerCase(),
      id: element.id || "",
      className: typeof element.className === "string" ? element.className.slice(0, 180) : "",
      role: element.getAttribute("role") || "",
      dataTid: element.getAttribute("data-tid") || element.getAttribute("data-testid") || "",
      box: {
        x: Math.round(rect.x),
        y: Math.round(rect.y),
        width: Math.round(rect.width),
        height: Math.round(rect.height),
      },
      overflowY: style.overflowY,
      clientHeight: element.clientHeight,
      scrollHeight: element.scrollHeight,
    };
  }

  const wrapperRect = transcriptWrapper?.getBoundingClientRect();
  const tabRect = transcriptTab?.getBoundingClientRect() || {
    bottom: wrapperRect?.top || 0,
    left: wrapperRect?.left || 0,
  };
  const timestampSearchRoot = transcriptWrapper || document.body;
  const allTimestampNodes = timestampTextNodes(timestampSearchRoot);
  const likelyTimestampNodes = allTimestampNodes.filter(({ rect }) => {
    return rect.top >= tabRect.bottom - 10
      && rect.bottom <= window.innerHeight + 5
      && rect.left >= Math.max(0, tabRect.left - 500);
  });

  const candidateStats = new Map();
  for (const match of likelyTimestampNodes) {
    let element = match.node.parentElement;
    let depth = 0;

    while (element && element !== document.body && element !== document.documentElement && depth < 20) {
      const rect = element.getBoundingClientRect();
      const canScroll = element.scrollHeight > element.clientHeight + 8
        && element.clientHeight >= 120
        && rect.width >= 280
        && rect.height >= 120;

      if (canScroll) {
        const existing = candidateStats.get(element) || {
          element,
          hits: 0,
          proximity: 0,
        };
        existing.hits += 1;
        existing.proximity += Math.max(1, 20 - depth);
        candidateStats.set(element, existing);
      }

      element = element.parentElement;
      depth += 1;
    }
  }

  const rankedCandidates = [...candidateStats.values()].map((candidate) => {
    const rect = candidate.element.getBoundingClientRect();
    const style = getComputedStyle(candidate.element);
    const attributes = [
      candidate.element.id,
      typeof candidate.element.className === "string" ? candidate.element.className : "",
      candidate.element.getAttribute("aria-label"),
      candidate.element.getAttribute("data-tid"),
      candidate.element.getAttribute("data-testid"),
    ].join(" ");
    let score = (candidate.hits * 500) + (candidate.proximity * 20);
    if (/transcript/i.test(attributes)) score += 300;
    if (/auto|scroll/i.test(style.overflowY)) score += 150;
    if (transcriptTab && candidate.element.contains(transcriptTab)) score -= 400;
    score -= Math.max(0, rect.width * rect.height / 100000);
    return { ...candidate, score };
  }).sort((left, right) => right.score - left.score);

  const explicitCandidates = [];
  if (transcriptWrapper) {
    explicitCandidates.push(transcriptWrapper, ...transcriptWrapper.querySelectorAll("*"));
    let ancestor = transcriptWrapper.parentElement;
    while (ancestor && ancestor !== document.body && ancestor !== document.documentElement) {
      explicitCandidates.push(ancestor);
      ancestor = ancestor.parentElement;
    }
  }

  const explicitScrollContainers = [...new Set(explicitCandidates)].filter((element) => {
    const rect = element.getBoundingClientRect();
    return isVisible(element)
      && element.scrollHeight > element.clientHeight + 8
      && element.clientHeight >= 120
      && rect.width >= 280;
  }).sort((left, right) => {
    const leftIsWrapper = left === transcriptWrapper ? 1 : 0;
    const rightIsWrapper = right === transcriptWrapper ? 1 : 0;
    if (leftIsWrapper !== rightIsWrapper) return rightIsWrapper - leftIsWrapper;

    const leftTimes = timestampTextNodes(left).length;
    const rightTimes = timestampTextNodes(right).length;
    if (leftTimes !== rightTimes) return rightTimes - leftTimes;

    const leftArea = left.clientWidth * left.clientHeight;
    const rightArea = right.clientWidth * right.clientHeight;
    return leftArea - rightArea;
  });

  const container = explicitScrollContainers[0] || rankedCandidates[0]?.element;
  if (!container) {
    return {
      ok: false,
      error: "The extension could not find the scrollable transcript list. Keep the transcript visible and try again.",
      diagnostics: {
        url: `${location.origin}${location.pathname}`,
        rootCount: documentRoots().length,
        closedShadowRootApi: Boolean(chrome.dom?.openOrClosedShadowRoot),
        tab: elementSummary(transcriptTab),
        transcriptWrapper: elementSummary(transcriptWrapper),
        explicitScrollContainers: explicitScrollContainers.slice(0, 8).map(elementSummary),
        allVisibleTimestampCount: allTimestampNodes.length,
        likelyTranscriptTimestampCount: likelyTimestampNodes.length,
        candidateCount: rankedCandidates.length,
        candidates: rankedCandidates.slice(0, 8).map((candidate) => ({
          ...elementSummary(candidate.element),
          hits: candidate.hits,
          proximity: candidate.proximity,
          score: Math.round(candidate.score),
        })),
      },
    };
  }

  function findRow(timestampNode) {
    let row = timestampNode.parentElement;
    let current = row;
    const containerRect = container.getBoundingClientRect();

    while (current?.parentElement && current.parentElement !== container) {
      const parent = current.parentElement;
      const rect = parent.getBoundingClientRect();
      const count = timestampCount(parent);

      if (count !== 1 || rect.height > Math.max(600, container.clientHeight * 0.8)) {
        break;
      }

      if (rect.width >= containerRect.width * 0.35) {
        row = parent;
      }
      current = parent;
    }

    return row;
  }

  function timestampToSeconds(timestamp) {
    return timestamp
      .split(":")
      .map(Number)
      .reduce((total, part) => (total * 60) + part, 0);
  }

  function parseRow(row, timestamp) {
    const rawText = String(row.innerText || "").trim();
    const lines = rawText
      .split(/\n+/)
      .map(normalise)
      .filter(Boolean)
      .filter((line) => !/^(play|more options)$/i.test(line));

    let speaker = "";
    let textParts = [];
    const exactIndex = lines.findIndex((line) => line === timestamp);

    if (exactIndex >= 0) {
      speaker = lines.slice(0, exactIndex).at(-1) || "";
      textParts = lines.slice(exactIndex + 1);
    } else {
      const lineIndex = lines.findIndex((line) => line.includes(timestamp));
      if (lineIndex >= 0) {
        const line = lines[lineIndex];
        const offset = line.indexOf(timestamp);
        speaker = normalise(line.slice(0, offset)) || lines.slice(0, lineIndex).at(-1) || "";
        const remainder = normalise(line.slice(offset + timestamp.length));
        textParts = [...(remainder ? [remainder] : []), ...lines.slice(lineIndex + 1)];
      }
    }

    const text = normalise(textParts.join(" "));
    if (!text) return null;

    return {
      timestamp,
      seconds: timestampToSeconds(timestamp),
      speaker: normalise(speaker) || "Unknown speaker",
      text,
      rawText,
    };
  }

  const entries = new Map();
  let firstSeen = 0;
  let mutationCount = 0;

  function collectVisibleRows() {
    const found = timestampTextNodes(container)
      .sort((left, right) => left.rect.top - right.rect.top);

    for (const match of found) {
      const row = findRow(match.node);
      const entry = parseRow(row, match.timestamp);
      if (!entry) continue;

      const key = [entry.timestamp, entry.speaker, entry.text].join("\u241f");
      if (!entries.has(key)) {
        entries.set(key, { ...entry, firstSeen: firstSeen++ });
      }
    }
  }

  const observer = new MutationObserver(() => {
    mutationCount += 1;
    collectVisibleRows();
  });
  observer.observe(container, { childList: true, subtree: true, characterData: true });

  const originalScrollTop = container.scrollTop;
  const startedAt = Date.now();
  let iterations = 0;
  let stableEndPasses = 0;

  async function settle() {
    await nextFrame();
    await wait(options.settleDelayMs);
    await nextFrame();
  }

  try {
    container.scrollTo({ top: 0, behavior: "auto" });
    await settle();

    while (iterations < 1000 && Date.now() - startedAt < options.maximumDurationMs) {
      iterations += 1;
      const countBefore = entries.size;
      collectVisibleRows();

      const maximum = Math.max(0, container.scrollHeight - container.clientHeight);
      const current = container.scrollTop;
      const atEnd = current >= maximum - 2;

      if (atEnd) {
        await settle();
        collectVisibleRows();
        const newMaximum = Math.max(0, container.scrollHeight - container.clientHeight);
        const stillAtEnd = container.scrollTop >= newMaximum - 2;
        stableEndPasses = stillAtEnd && entries.size === countBefore
          ? stableEndPasses + 1
          : 0;
        if (stableEndPasses >= 3) break;
      } else {
        stableEndPasses = 0;
      }

      const step = Math.max(180, Math.floor(container.clientHeight * 0.65));
      const next = Math.min(Math.max(0, container.scrollHeight - container.clientHeight), current + step);
      if (next <= current + 1 && atEnd) continue;
      container.scrollTo({ top: next, behavior: "auto" });
      await settle();
    }
  } finally {
    observer.disconnect();
    container.scrollTo({ top: originalScrollTop, behavior: "auto" });
  }

  const orderedEntries = [...entries.values()]
    .sort((left, right) => left.seconds - right.seconds || left.firstSeen - right.firstSeen)
    .map(({ firstSeen: _firstSeen, ...entry }) => entry);

  if (orderedEntries.length === 0) {
    return {
      ok: false,
      error: "The transcript list was found, but no timestamped transcript entries were captured.",
    };
  }

  return {
    ok: true,
    capturedAt: new Date().toISOString(),
    pageTitle: document.title,
    source: `${location.origin}${location.pathname}`,
    entries: orderedEntries,
    diagnostics: {
      iterations,
      mutationCount,
      durationMs: Date.now() - startedAt,
      container: {
        tag: container.tagName.toLowerCase(),
        className: typeof container.className === "string" ? container.className.slice(0, 240) : "",
        dataTid: container.getAttribute("data-tid") || container.getAttribute("data-testid") || "",
        clientHeight: container.clientHeight,
        scrollHeight: container.scrollHeight,
      },
    },
  };
}
