'use strict';
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp } = require('./harness');

test('app.js loads and evaluates to completion', () => {
  const w = loadApp();
  assert.strictEqual(typeof w.escHtml, 'function');
  assert.strictEqual(typeof w.saveClient, 'function');
  assert.ok(Array.isArray(w._settingsRenewals));
});

test('escHtml neutralises the four HTML-significant characters', () => {
  const w = loadApp();
  assert.strictEqual(w.escHtml('<script>"x"&\'y\''), '&lt;script&gt;&quot;x&quot;&amp;&#39;y&#39;');
});
