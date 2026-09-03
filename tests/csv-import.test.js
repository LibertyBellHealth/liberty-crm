'use strict';
// The CSV import is where data from outside the app arrives in bulk — a carrier's export, a
// spreadsheet someone was emailed — and it writes straight to client records.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, jsonEqual } = require('./harness');

test('a quoted comma does not shift every column after it', () => {
  const w = loadApp();
  const csv = 'First Name,Last Name,Address,Phone\n' +
              'Ada,Lovelace,"123 Main St, Apt 4",555-0100\n';
  const out = w.parseCSV(csv);

  // jsonEqual, not deepStrictEqual: values come back in the jsdom realm.
  assert.ok(jsonEqual(out.headers, ['First Name', 'Last Name', 'Address', 'Phone']), 'headers: ' + JSON.stringify(out.headers));
  assert.strictEqual(out.rows.length, 1);
  // Split on every comma, the address becomes "123 Main St" and Apt 4 lands in Phone —
  // silently writing the wrong value into a patient's phone number.
  assert.strictEqual(out.rows[0]['Address'], '123 Main St, Apt 4');
  assert.strictEqual(out.rows[0]['Phone'], '555-0100');
});

test('escaped quotes, embedded newlines and CRLF survive the parse', () => {
  const w = loadApp();
  const csv = 'Name,Notes\r\n' +
              '"Ada","She said ""hello"""\r\n' +
              '"Bo","line one\nline two"\r\n';
  const out = w.parseCSV(csv);

  assert.strictEqual(out.rows.length, 2, 'a newline inside a quoted field split the row');
  assert.strictEqual(out.rows[0]['Notes'], 'She said "hello"');
  assert.strictEqual(out.rows[1]['Notes'], 'line one\nline two');
  assert.strictEqual(out.rows[0]['Name'], 'Ada', 'CRLF left a stray carriage return');
});

test('a byte-order mark does not become part of the first header name', () => {
  const w = loadApp();
  const out = w.parseCSV('﻿First Name,Last Name\nAda,Lovelace\n');
  assert.strictEqual(out.headers[0], 'First Name');
  assert.strictEqual(out.rows[0]['First Name'], 'Ada');
});

test('a blank trailing line does not become an empty client row', () => {
  const w = loadApp();
  const out = w.parseCSV('First Name\nAda\n\n');
  assert.strictEqual(out.rows.length, 1);
});

// ── The column-mapping table renders the CSV's own header names. They are attacker-supplied
// text on a file the agent was sent, and they were going into markup unescaped.
test('a CSV header name cannot inject markup into the mapping table', () => {
  const w = loadApp();
  resetStorage(w);
  let body = w.document.getElementById('mappingBody');
  if (!body) {
    const d = w.document.createElement('div');
    d.innerHTML = '<table><tbody id="mappingBody"></tbody></table><div id="mappingSection"></div>';
    w.document.body.appendChild(d);
    body = w.document.getElementById('mappingBody');
  }

  w.renderCsvMapping(['First Name', '<img src=x onerror="window.__pwned=true">']);

  assert.strictEqual(body.querySelectorAll('img').length, 0, 'the header name became a real element');
  const opts = Array.from(body.querySelectorAll('option')).map(o => o.textContent);
  assert.ok(opts.includes('<img src=x onerror="window.__pwned=true">'),
    'the header should still be offered verbatim as text');
});
