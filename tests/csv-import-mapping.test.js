'use strict';
// The importer guessed each field's column by taking the label's first four letters and selecting
// EVERY header containing them — so the LAST match won. "Date of Birth" reduced to "date", which
// matches "Lead Date", "Effective Date" and "Application Date"; whichever came last in the file
// became the patient's date of birth. Across a thousand imported records that is a scrambled
// dataset with no error raised anywhere, and after import a wrong DOB is indistinguishable from
// one the patient gave us.
//
// Two more defects sat alongside it: the mapping table the operator saw was a DUPLICATE set of
// element ids (getElementById always returned the hidden Import-view copy, so the Settings
// importer ran on mappings nobody could see), and an SSN column mapped to f_ssnLast4, which
// clientToDbRow has no key for — so it was silently discarded.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, confirmOk, stub } = require('./harness');

function app() {
  const w = loadApp();
  resetStorage(w);
  if (!w.document.getElementById('mappingBody')) {
    w.document.body.insertAdjacentHTML('beforeend',
      '<div id="mappingSection"><table><tbody id="mappingBody"></tbody></table>' +
      '<div id="importPreview"></div><span id="importStatus"></span></div>');
  }
  w.document.getElementById('mappingBody').innerHTML = '';
  w.document.getElementById('importStatus').innerHTML = '';
  w.IMPORT_BATCH_SIZE = 500;
  w.eval('csvHeaders=[];csvData=[];_importFailedRows=[];');
  // The import refreshes the roster when it finishes. That is correct behaviour but it shares the
  // fetch stub, so leave it out of these tests rather than teaching every stub to answer a GET.
  stub(w, { loadClients: () => {} });
  return w;
}
const load = (w, headers, rows) =>
  w.eval('csvHeaders=' + JSON.stringify(headers) + ';csvData=' + JSON.stringify(rows) + ';');
const mapped = (w, key) => {
  const sel = w.document.getElementById('map_' + key);
  return sel ? sel.value : null;
};
const settle = () => new Promise((r) => setTimeout(r, 0));

test('a later "Application Date" column does not become the date of birth', () => {
  const w = app();
  const headers = ['First Name', 'Date of Birth', 'Lead Date', 'Application Date'];
  load(w, headers, [{ 'First Name': 'Ada', 'Date of Birth': '1815-12-10',
                      'Lead Date': '2026-01-02', 'Application Date': '2026-03-04' }]);
  w.renderCsvMapping(headers);

  assert.strictEqual(mapped(w, 'f_dob'), 'Date of Birth',
    'the exact "Date of Birth" header must win over the other columns containing "date"');
});

test('an exact header match beats a partial one regardless of column order', () => {
  const w = app();
  const headers = ['Primary Phone', 'Phone'];
  load(w, headers, [{ 'Primary Phone': '555-0001', Phone: '555-0002' }]);
  w.renderCsvMapping(headers);
  assert.strictEqual(mapped(w, 'f_phone'), 'Phone');
});

test('a header that matches nothing is left unmapped rather than guessed at', () => {
  const w = app();
  const headers = ['First Name', 'Widget Code'];
  load(w, headers, [{ 'First Name': 'Ada', 'Widget Code': 'XYZ' }]);
  w.renderCsvMapping(headers);
  assert.strictEqual(mapped(w, 'f_notes'), '', 'an unrelated column must not be guessed into Notes');
  assert.strictEqual(mapped(w, 'f_firstName'), 'First Name');
});

test('a "Status" column is not mistaken for the "State" field', () => {
  const w = app();
  // The old rule took the label's first four letters — "State" -> "stat" — and matched any header
  // CONTAINING them, so a "Status" column was written into every patient's state of residence.
  const headers = ['First Name', 'Status'];
  load(w, headers, [{ 'First Name': 'Ada', Status: 'Active' }]);
  w.renderCsvMapping(headers);
  assert.strictEqual(mapped(w, 'f_resSt'), '',
    '"Status" must not be guessed into the State field');
});

test('when two headers match equally well, the first one wins', () => {
  const w = app();
  const headers = ['Notes A', 'Notes B'];
  load(w, headers, [{ 'Notes A': 'first', 'Notes B': 'second' }]);
  w.renderCsvMapping(headers);
  assert.strictEqual(mapped(w, 'f_notes'), 'Notes A',
    'a tie must resolve to the earlier column, not whichever happens to come last');
});

test('an imported SSN column reaches the database row instead of being dropped', () => {
  const w = app();
  const headers = ['First Name', 'SSN'];
  load(w, headers, [{ 'First Name': 'Ada', SSN: '123-45-6789' }]);
  w.renderCsvMapping(headers);

  const data = w._importRowToData(w.csvData[0]);
  const row = w.clientToDbRow(data);
  assert.strictEqual(row.ssn, '123-45-6789',
    'f_ssnLast4 has no clientToDbRow mapping — a mapped SSN column was silently discarded');
});

test('the preview shows what will actually be written, row by row', () => {
  const w = app();
  const headers = ['First Name', 'Date of Birth', 'Application Date'];
  load(w, headers, [{ 'First Name': 'Ada', 'Date of Birth': '1815-12-10', 'Application Date': '2026-03-04' }]);
  w.renderCsvMapping(headers);

  const text = w.document.getElementById('importPreview').textContent;
  assert.ok(text.includes('1815-12-10'), 'preview should show the real DOB, got: ' + text);
  assert.ok(!text.includes('2026-03-04'), 'the application date is not mapped and must not appear');
});

test('nothing is written until the operator confirms the mapping', async () => {
  const w = app();
  const headers = ['First Name'];
  load(w, headers, [{ 'First Name': 'Ada' }]);
  w.renderCsvMapping(headers);
  let calls = 0;
  stub(w, { fetch: (url) => {
    if (String(url).endsWith('/health-clients/bulk')) calls++;
    return Promise.resolve({ ok: true, json: () => Promise.resolve({ inserted: 1, ids: [1] }) });
  } });

  w.importClients();
  // Must settle first: the batch chain starts from Promise.resolve(), so the first POST is a
  // microtask away whether or not the confirm gate exists. Asserting synchronously tested the
  // scheduling, not the gate.
  await settle(); await settle();
  assert.strictEqual(calls, 0, 'the import must not write before the confirm is accepted');

  confirmOk(w);
  await settle();
  assert.strictEqual(calls, 1, 'accepting the confirm should send the batch');
});

test('rows go to the transactional bulk endpoint, in batches', async () => {
  const w = app();
  const headers = ['First Name'];
  load(w, headers, [{ 'First Name': 'A' }, { 'First Name': 'B' }, { 'First Name': 'C' }]);
  w.renderCsvMapping(headers);
  w.IMPORT_BATCH_SIZE = 2;
  const seen = [];
  stub(w, { fetch: (url, opt) => {
    const n = opt && opt.body ? JSON.parse(opt.body).length : 0;
    if (String(url).endsWith('/health-clients/bulk')) seen.push({ url: String(url), rows: n });
    return Promise.resolve({ ok: true, json: () => Promise.resolve({ inserted: n, ids: [] }) });
  } });

  w.importClients(); confirmOk(w);
  await settle(); await settle(); await settle();

  assert.strictEqual(seen.length, 2, 'three rows at a batch size of two should be two batches');
  assert.ok(seen[0].url.endsWith('/health-clients/bulk'), 'got ' + seen[0].url);
  assert.deepStrictEqual(seen.map((s) => s.rows), [2, 1]);
});

test('a rolled-back batch names the file lines that did not import, and can be retried', async () => {
  const w = app();
  const headers = ['First Name'];
  load(w, headers, [{ 'First Name': 'A' }, { 'First Name': 'B' }, { 'First Name': 'C' }]);
  w.renderCsvMapping(headers);
  w.IMPORT_BATCH_SIZE = 2;
  let call = 0;
  stub(w, { fetch: () => {
    call++;
    // First batch (rows 1-2 → file lines 2-3) rolls back; the second succeeds.
    if (call === 1) return Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'Server error. Please try again.' }) });
    return Promise.resolve({ ok: true, json: () => Promise.resolve({ inserted: 1, ids: [9] }) });
  } });

  w.importClients(); confirmOk(w);
  await settle(); await settle(); await settle();

  const status = w.document.getElementById('importStatus').textContent;
  assert.ok(status.includes('lines 2') && status.includes('3'),
    'the failed rows must be identified by file line, got: ' + status);
  assert.ok(status.includes('Imported 1 of 3'), 'got: ' + status);
  assert.strictEqual(w._importFailedRows.length, 2,
    'exactly the rolled-back rows should be held for retry');
  assert.strictEqual(w._importFailedRows[0]['First Name'], 'A');
});

// The import itself had the defect it exists to prevent: the mapping was read from the DOM afresh
// for every batch, and the row set was read inside the confirm callback rather than captured. A
// mapping edited part-way through a long import would silently apply to the remaining batches.
test('a mapping changed mid-import does not affect the batches still to go', async () => {
  const w = app();
  const headers = ['First Name', 'Other'];
  load(w, headers, [{ 'First Name': 'A', Other: 'wrong-1' }, { 'First Name': 'B', Other: 'wrong-2' }]);
  w.renderCsvMapping(headers);
  w.IMPORT_BATCH_SIZE = 1;
  const sent = [];
  stub(w, { fetch: (url, opt) => {
    if (String(url).endsWith('/health-clients/bulk')) {
      sent.push(JSON.parse(opt.body));
      // Between batch one and batch two, the mapping on screen changes.
      w.document.getElementById('map_f_firstName').value = 'Other';
    }
    return Promise.resolve({ ok: true, json: () => Promise.resolve({ inserted: 1, ids: [1] }) });
  } });

  w.importClients(); confirmOk(w);
  await settle(); await settle(); await settle();

  assert.strictEqual(sent.length, 2, 'expected two batches, got ' + sent.length);
  assert.strictEqual(sent[1][0].first_name, 'B',
    'the second batch must use the mapping that was confirmed, not the one now on screen');
});

test('the audit entry records what actually imported, not what was attempted', async () => {
  const w = app();
  const headers = ['First Name'];
  load(w, headers, [{ 'First Name': 'A' }, { 'First Name': 'B' }]);
  w.renderCsvMapping(headers);
  const logged = [];
  stub(w, {
    logActivity: (type, text) => { logged.push(text); },
    fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'Server error. Please try again.' }) }),
  });

  w.importClients(); confirmOk(w);
  await settle(); await settle(); await settle();

  assert.strictEqual(logged.length, 1, 'expected exactly one audit entry, got ' + logged.length);
  assert.ok(logged[0].startsWith('0 of 2'),
    'nothing committed, so the audit row must not claim 2 were imported — got: ' + logged[0]);
});

test('retrying failed rows reuses the confirmed mapping, not whatever is on screen', async () => {
  const w = app();
  const headers = ['First Name', 'Other'];
  load(w, headers, [{ 'First Name': 'A', Other: 'wrong' }]);
  w.renderCsvMapping(headers);
  let call = 0;
  const sent = [];
  stub(w, { fetch: (url, opt) => {
    if (!String(url).endsWith('/health-clients/bulk')) return Promise.resolve({ ok: true, json: () => Promise.resolve({}) });
    call++;
    sent.push(JSON.parse(opt.body));
    if (call === 1) return Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'boom' }) });
    return Promise.resolve({ ok: true, json: () => Promise.resolve({ inserted: 1, ids: [7] }) });
  } });

  w.importClients(); confirmOk(w);
  await settle(); await settle();
  assert.strictEqual(w._importFailedRows.length, 1);

  w.document.getElementById('map_f_firstName').value = 'Other';   // operator fiddles after the failure
  w.retryFailedImportRows();
  await settle(); await settle();

  assert.strictEqual(sent.length, 2, 'the retry should have re-sent the failed row');
  assert.strictEqual(sent[1][0].first_name, 'A',
    'the retry must use the mapping the rows were confirmed under');
});

// ── Found by the independent review, written before the fixes ────────────────────────────────

test('a "Carrier" column can be imported at all', () => {
  const w = app();
  const headers = ['First Name', 'Carrier'];
  load(w, headers, [{ 'First Name': 'Ada', Carrier: 'BCBS' }]);
  w.renderCsvMapping(headers);
  assert.strictEqual(mapped(w, 'f_planCarrier'), 'Carrier',
    'Carrier has no destination field, so the column is silently dropped');
  const row = w.clientToDbRow(w._importRowToData(w.csvData[0]));
  assert.strictEqual(row.plan_carrier, 'BCBS');
});

test('a qualified header still auto-maps', () => {
  const w = app();
  const headers = ['Client First', 'Client Last', 'Member Zip'];
  load(w, headers, [{ 'Client First': 'Ada', 'Client Last': 'Lovelace', 'Member Zip': '48201' }]);
  w.renderCsvMapping(headers);
  assert.strictEqual(mapped(w, 'f_firstName'), 'Client First');
  assert.strictEqual(mapped(w, 'f_lastName'), 'Client Last');
  assert.strictEqual(mapped(w, 'f_resZip'), 'Member Zip');
});

test('the confirmation names the CSV columns that will be ignored', () => {
  const w = app();
  const headers = ['First Name', 'Widget Code', 'Internal Ref'];
  load(w, headers, [{ 'First Name': 'Ada', 'Widget Code': 'X', 'Internal Ref': 'Y' }]);
  w.renderCsvMapping(headers);
  let msg = '';
  stub(w, { showConfirm: (m) => { msg = String(m); } });
  w.importClients();
  assert.match(msg, /Widget Code/, 'an ignored column must be named: ' + msg);
  assert.match(msg, /Internal Ref/);
});

test('changing a mapping clears its "guessed" marker', () => {
  const w = app();
  const headers = ['Primary Phone', 'Other'];
  load(w, headers, [{ 'Primary Phone': '555', Other: 'x' }]);
  w.renderCsvMapping(headers);
  const sel = w.document.getElementById('map_f_phone');
  assert.match(sel.parentNode.textContent, /guessed/, 'precondition: it was a guess');
  sel.value = 'Other';
  sel.dispatchEvent(new w.Event('change'));
  assert.ok(!/guessed/.test(sel.parentNode.textContent),
    'the operator chose this column — it is no longer a guess');
});

test('picking a new file clears a previous failure and its retry rows', async () => {
  const w = app();
  const headers = ['First Name'];
  load(w, headers, [{ 'First Name': 'A' }]);
  w.renderCsvMapping(headers);
  stub(w, { fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'boom' }) }) });
  w.importClients(); confirmOk(w);
  await settle(); await settle();
  assert.strictEqual(w._importFailedRows.length, 1, 'precondition: a failed batch is held');

  // Operator picks a different file.
  w.handleCSV({ target: { files: [new w.Blob(['Last Name\nLovelace\n'], { type: 'text/csv' })] } });
  await new Promise((r) => setTimeout(r, 30));

  assert.strictEqual(w._importFailedRows.length, 0,
    "the previous file's rows are still queued behind a live Retry button");
  assert.strictEqual(w.document.getElementById('importStatus').textContent, '',
    "the previous file's failure banner is still on screen");
});

test('a rolled-back batch tells the operator nothing landed', async () => {
  const w = app();
  const headers = ['First Name'];
  load(w, headers, [{ 'First Name': 'A' }]);
  w.renderCsvMapping(headers);
  stub(w, { fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({
    error: 'Server error. Please try again.', inserted: 0,
    detail: 'The whole batch was rolled back — no clients were imported.' }) }) });
  w.importClients(); confirmOk(w);
  await settle(); await settle();
  assert.match(w.document.getElementById('importStatus').textContent, /rolled back/i,
    'the one sentence saying nothing partially landed is dropped: ' + w.document.getElementById('importStatus').textContent);
});
