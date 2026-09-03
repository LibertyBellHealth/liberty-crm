'use strict';
// MSAL token cache location — and the tradeoff behind it, recorded so it is not re-litigated.
//
// THE GAP (real, still open): clearCRMStorage wipes crm_* / crmCarriers / lch_* on sign-out and on
// tab close, but never touches msal.*. So closing the tab clears the cached roster while leaving
// the access and refresh tokens that fetch it back, on disk, across a browser restart. The
// 45-minute idle timeout has the same hole — it signs this tab out, not the stored tokens.
//
// WHAT WAS TRIED: 2026-09-03, sessionStorage (PR #18). It was deployed and REVERTED the same day
// (c932e1c) because sign-in did not hold up. Home Care's initMSAL carries a matching warning:
// iOS Safari's Intelligent Tracking Prevention wipes sessionStorage mid-redirect during the OAuth
// dance. Do not re-propose sessionStorage without a way to test the real redirect on iOS.
//
// WHERE THE FIX ACTUALLY LIVES: token lifetime is an identity-provider concern. An Azure AD
// Conditional Access sign-in-frequency policy expires the tokens themselves; nothing in this file
// can, because MSAL must be able to read its own cache to use it.
//
// This test pins the CURRENT, deliberate configuration so the revert cannot be undone by accident.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

function msalConfig(w) {
  let captured = null;
  const orig = w.msal;
  w.msal = { PublicClientApplication: function (cfg) {
    captured = cfg;
    return { initialize: () => new Promise(() => {}), getAllAccounts: () => [] };
  } };
  try { w.initMSAL(); } finally { w.msal = orig; }
  return captured;
}

test('MSAL uses localStorage — deliberate, after sessionStorage was tried and reverted', () => {
  const cfg = msalConfig(loadApp());
  assert.ok(cfg, 'initMSAL did not construct an MSAL instance');
  assert.strictEqual(cfg.cache.cacheLocation, 'localStorage',
    'sessionStorage was deployed on 2026-09-03 and reverted the same day — see the note at the top of this file');
});

test('auth state still rides in a cookie, which the redirect depends on', () => {
  const cfg = msalConfig(loadApp());
  // Required either way: iOS Safari ITP can clear web storage mid-redirect, and this is what
  // lets the OAuth round trip still validate.
  assert.strictEqual(cfg.cache.storeAuthStateInCookie, true);
});
