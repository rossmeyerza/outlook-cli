const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');
const diagnostics = require('../diagnostics.js');

test('parent learns navigated frame URL even when src stays about:blank', async () => {
  let message;
  const child = { postMessage() {} };
  const frame = {
    tagName: 'IFRAME', src: 'about:blank', contentWindow: child,
    getAttribute: () => null,
    getBoundingClientRect: () => ({ left: 0, right: 900, top: 0, bottom: 700, width: 900, height: 700 }),
  };
  const doc = { createTreeWalker: () => { let done = false; return { currentNode: frame, nextNode: () => { if (!done) done = true; return null; } }; } };
  const context = vm.createContext({
    document: doc, OcliDiagnostics: diagnostics, URL, setTimeout,
    location: { hostname: 'tenant.sharepoint.com', origin: 'https://tenant.sharepoint.com', pathname: '/Doc.aspx' },
    chrome: { runtime: { id: 'test' } },
    innerWidth: 1200, innerHeight: 800,
    getComputedStyle: () => ({ display: 'block', visibility: 'visible', opacity: '1' }),
    addEventListener: (_, fn) => { message = fn; },
  });
  vm.runInContext('globalThis.window = globalThis; window.top = window;', context);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../controller.js'), 'utf8'), context);
  message({ source: child, origin: 'https://eu.officeapps.live.com', data: {
    __ocliCopilot: 'frame_identity', url: 'https://eu.officeapps.live.com/we/wordeditorframe.aspx?secret=omitted',
  } });
  const info = await vm.runInContext('__ocliController.run({action:"describe"})', context);
  assert.equal(info.embeddedFrames[0].url, 'https://eu.officeapps.live.com/we/wordeditorframe.aspx');
  assert.equal(info.embeddedFrames[0].visible, true);
});
