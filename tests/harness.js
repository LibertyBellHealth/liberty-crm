'use strict';
// Loads the real app.js into a jsdom window so its functions can be exercised in Node.
// Nothing is copied or re-implemented here — the tests call the SAME code that ships.
const { JSDOM } = require('jsdom');
const fs = require('fs');
const path = require('path');

let _win = null;

function loadApp() {
  if (_win) return _win;
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  // app.js's last statement calls initMSAL()/loadCarriers()/loadSettingsExtras(), which reach for
  // real page elements. Those are wrapped in the file's own try/catch, but the settings renderers
  // are not, so give them the handful of containers they look up. Anything missing is guarded by
  // `if(!el)return` inside app.js itself.
  const initDom = `
    <div class="sidebar"><div class="logo">Liberty Bell Health</div></div>
    <div id="authScreen"></div><div id="authScreenMsg"></div><div id="mainApp"></div>
    <div id="clientTableBody"></div><div id="clientDocsSection"></div>
    <div id="clientAuditSection"></div><div id="reportCards"></div>
    <div id="agentSettingsList"></div><div id="leadSourceSettingsList"></div>
    <div id="customMedSettingsList"></div><div id="renewalSettingsList"></div>
    <div id="planTypeSettingsList"></div><div id="projectCodeSettingsList"></div>
    <input id="settingsCrmName"><div id="toastContainer"></div>
    <div id="confirmModal"><div id="confirmTitle"></div><div id="confirmMessage"></div>
      <button id="confirmOkBtn"></button><button id="confirmCancelBtn"></button>
      <button id="confirmExtraBtn"></button></div>`;
  const dom = new JSDOM('<!doctype html><html><head></head><body>' + initDom + '</body></html>', {
    url: 'http://localhost/',
    runScripts: 'outside-only',
    pretendToBeVisual: true,
  });
  const { window } = dom;
  // CDN globals that index.html loads before app.js. Not present in jsdom; stub what startup touches.
  // A no-op MSAL so the file's final initMSAL() call doesn't log a constructor error on every run.
  window.msal = window.msal || { PublicClientApplication: function () {
    return { initialize: () => new Promise(() => {}), getAllAccounts: () => [] };
  } };
  window.appInsights = window.appInsights || {
    trackEvent() {}, trackException() {}, trackPageView() {}, trackTrace() {},
  };
  window.ZIP_COUNTIES = window.ZIP_COUNTIES || {};
  window.ROUTING_LOOKUP = window.ROUTING_LOOKUP || {};
  // No test should ever reach the network. Anything that tries is a bug in the test.
  window.fetch = function () { throw new Error('harness: unexpected network call'); };

  // Evaluate in global scope and let ANY top-level exception fail the suite loudly — a half-loaded
  // app.js still hoists its function declarations, so tests would keep passing while whole regions
  // of the file (every `var X = [...]` after the throw) were silently undefined.
  window.eval(src);

  // Sentinel from the LAST few lines of app.js. If evaluation ever stops early without throwing,
  // this is undefined and the whole suite fails instead of quietly testing a truncated app.
  if (typeof window._settingsRenewals === 'undefined') {
    throw new Error('harness: app.js did not evaluate to completion — late declarations are missing');
  }
  _win = window;
  return window;
}

// Replacing one of app.js's own globals (to observe a call, or to keep a test off the network) has
// to be undone, or the next test calls the stub instead of the real function and passes for the
// wrong reason — or, worse, fails for one. Stubs registered here are restored by resetStorage().
const _stubs = [];
function stub(win, map) {
  Object.keys(map).forEach(function (name) {
    _stubs.push({ win, name, had: name in win, orig: win[name] });
    win[name] = map[name];
  });
}
function restoreStubs() {
  while (_stubs.length) {
    const s = _stubs.pop();
    if (s.had) s.win[s.name] = s.orig; else delete s.win[s.name];
  }
}

// Fresh localStorage per test group. loadApp() caches ONE window, so module-scope state otherwise
// leaks between tests and can make a later test pass for the wrong reason.
function resetStorage(win) {
  restoreStubs();
  win.localStorage.clear();
  try { win.eval('clients = []; carriers = []; _recentRecords = []; _todos = []; _clientDocs = [];'); } catch (e) {}
}

// Realm-safe deep compare. Values returned from app.js live in the jsdom realm, so
// assert.deepStrictEqual fails on prototype identity against a Node-realm literal.
function jsonEqual(a, b) { return JSON.stringify(a) === JSON.stringify(b); }

// Click the real confirm dialog's OK button. Tests that go through showConfirm should use this
// rather than stubbing the dialog out — the gap between "state read at click time" and "state when
// OK is pressed" is exactly where this class of bug lives, and a stub closes that gap invisibly.
function confirmOk(win) {
  const btn = win.document.getElementById('confirmOkBtn');
  if (!btn || win.document.getElementById('confirmModal').style.display !== 'flex') {
    throw new Error('confirmOk: no confirm dialog is open');
  }
  btn.dispatchEvent(new win.Event('click'));
}

module.exports = { loadApp, resetStorage, jsonEqual, confirmOk, stub };
