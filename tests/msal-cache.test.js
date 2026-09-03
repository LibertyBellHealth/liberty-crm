'use strict';
// Tokens must not outlive the PHI they unlock. clearCRMStorage wipes the cached roster on
// sign-out and on tab close but never touches msal.*, so tokens in localStorage survived a
// browser restart and could fetch everything back.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('MSAL keeps its tokens in sessionStorage, not on disk', () => {
  const w = loadApp();
  let captured = null;
  const origMsal = w.msal;
  w.msal = { PublicClientApplication: function (cfg) {
    captured = cfg;
    return { initialize: () => new Promise(() => {}), getAllAccounts: () => [] };
  } };
  try { w.initMSAL(); } finally { w.msal = origMsal; }

  assert.ok(captured, 'initMSAL did not construct an MSAL instance');
  assert.strictEqual(captured.cache.cacheLocation, 'sessionStorage',
    'tokens in localStorage outlive the tab-close PHI wipe');
  // Needed for the redirect flow to survive the round trip when the cache is per-tab.
  assert.strictEqual(captured.cache.storeAuthStateInCookie, true);
});
