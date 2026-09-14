'use strict';
// Found by importing a 20-row test file on the dev backend. An export's "04/04/1953" was stored as
// written; <input type="date"> can't show it, so the first ordinary save of that patient sent
// dob:'' and erased it. Running the same file twice also doubled every patient.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, confirmOk, stub, formDom } = require('./harness');

const settle = () => new Promise((r) => setTimeout(r, 0));

function app(roster) {
  const w = loadApp();
  resetStorage(w);
  if (!w.document.getElementById('mappingBody')) {
    w.document.body.insertAdjacentHTML('beforeend',
      '<div id="mappingSection"><table><tbody id="mappingBody"></tbody></table>' +
      '<div id="importPreview"></div><span id="importStatus"></span></div>');
  }
  w.document.getElementById('mappingBody').innerHTML = '';
  w.document.getElementById('importStatus').textContent = '';
  w.IMPORT_BATCH_SIZE = 500;
  w.eval('csvHeaders=[];csvData=[];_importFailedRows=[];');
  const sent = [];
  let rosterCalls = 0;
  stub(w, {
    loadClients: () => {},
    toast: () => {},
    _fetchImportRoster: () => { rosterCalls++; return roster instanceof Error ? Promise.reject(roster) : Promise.resolve(roster || []); },
    fetch: (url, opt) => {
      if (String(url).endsWith('/health-clients/bulk')) {
        const rows = JSON.parse(opt.body);
        sent.push(rows);
        return Promise.resolve({ ok: true, json: () => Promise.resolve({ inserted: rows.length, ids: [] }) });
      }
      throw new Error('unexpected fetch ' + url);
    },
  });
  return { w, sent, rosterCalls: () => rosterCalls };
}
function load(w, headers, rows) {
  w.eval('csvHeaders=' + JSON.stringify(headers) + ';csvData=' + JSON.stringify(rows) + ';');
  w.renderCsvMapping(headers);
}
const status = (w) => w.document.getElementById('importStatus').textContent;
const message = (w) => w.document.getElementById('confirmMessage').textContent;
const clickExtra = (w) => w.document.getElementById('confirmExtraBtn').dispatchEvent(new w.Event('click'));

test('export date formats are read as real dates, and non-dates are refused', () => {
  const w = loadApp();
  const cases = {
    '04/04/1953': '1953-04-04', '4/4/1953': '1953-04-04', '1953-04-04': '1953-04-04',
    '1953/4/4': '1953-04-04', '1953-04-04T00:00:00': '1953-04-04', '4/4/1953 12:00:00 AM': '1953-04-04',
    '': '', '  ': '',
    '13/45/1953': null, '02/30/1990': null, '04/04/53': null, 'April 4 1953': null, '04/04/1853': null,
  };
  Object.keys(cases).forEach((input) => {
    assert.strictEqual(w._isoDate(input), cases[input], JSON.stringify(input));
  });
});

test('an imported MM/DD/YYYY birth date is sent as YYYY-MM-DD', async () => {
  const { w, sent } = app();
  load(w, ['First Name', 'Last Name', 'Date of Birth', 'Health Effective'],
    [{ 'First Name': 'Ada', 'Last Name': 'Test', 'Date of Birth': '04/04/1953', 'Health Effective': '3/1/2026' }]);

  w.importClients(); await settle(); confirmOk(w); await settle();

  assert.strictEqual(sent.length, 1);
  assert.strictEqual(sent[0][0].dob, '1953-04-04');
  assert.strictEqual(sent[0][0].health_effective, '2026-03-01');
});

test('a patient saved with an MM/DD/YYYY birth date keeps it through an unrelated save', () => {
  const w = loadApp();
  resetStorage(w);
  formDom(w);
  w.setFormData({ f_firstName: 'Ada', f_dob: '04/04/1953', f_healthEffective: '03/01/2026' });

  const data = w.getFormData();
  assert.strictEqual(data.f_dob, '1953-04-04', 'the date box showed blank, so saving erased the DOB');
  assert.strictEqual(data.f_healthEffective, '2026-03-01');
});

test('a file with an unreadable date imports nothing and names the line', async () => {
  const { w, sent, rosterCalls } = app();
  load(w, ['First Name', 'Date of Birth'],
    [{ 'First Name': 'A', 'Date of Birth': '01/02/1950' }, { 'First Name': 'B', 'Date of Birth': '13/45/1953' }]);

  w.importClients(); await settle();

  assert.strictEqual(w.document.getElementById('confirmModal').style.display === 'flex', false, 'no confirm should open');
  assert.strictEqual(sent.length, 0);
  assert.strictEqual(rosterCalls(), 0);
  assert.match(status(w), /Date of Birth on line 3/);
  assert.match(status(w), /13\/45\/1953/);
});

test('rows already in the CRM, or repeated in the file, are skipped and named', async () => {
  const { w, sent } = app([{ f_firstName: 'Ada', f_lastName: 'Test', f_dob: '1953-04-04' }]);
  load(w, ['First Name', 'Last Name', 'Date of Birth'], [
    { 'First Name': 'ADA', 'Last Name': 'test', 'Date of Birth': '04/04/1953' },   // line 2: in CRM
    { 'First Name': 'Bo', 'Last Name': 'Test', 'Date of Birth': '01/02/1960' },    // line 3: new
    { 'First Name': 'Bo', 'Last Name': 'Test', 'Date of Birth': '1960-01-02' },    // line 4: repeat of 3
    { 'First Name': 'Ada', 'Last Name': 'Test', 'Date of Birth': '05/05/1970' },   // line 5: same name, other DOB
  ]);

  w.importClients(); await settle();
  assert.match(message(w), /Already in the CRM.*line 2\b/);
  assert.match(message(w), /Repeated earlier in this file.*line 4\b/);
  confirmOk(w); await settle();

  assert.strictEqual(sent.length, 1);
  assert.deepStrictEqual(sent[0].map((r) => r.first_name + ' ' + r.dob), ['Bo 1960-01-02', 'Ada 1970-05-05']);
});

test('"Import all anyway" still imports every row when the operator chooses it', async () => {
  const { w, sent } = app([{ f_firstName: 'Ada', f_lastName: 'Test', f_dob: '04/04/1953' }]);
  load(w, ['First Name', 'Last Name', 'Date of Birth'],
    [{ 'First Name': 'Ada', 'Last Name': 'Test', 'Date of Birth': '1953-04-04' }]);

  w.importClients(); await settle();
  clickExtra(w); await settle();

  assert.strictEqual(sent.length, 1);
  assert.strictEqual(sent[0].length, 1);
});

test('if the current roster cannot be read, nothing is imported', async () => {
  const { w, sent } = app(new Error('HTTP 503'));
  load(w, ['First Name'], [{ 'First Name': 'A' }]);

  w.importClients(); await settle(); await settle();

  assert.strictEqual(w.document.getElementById('confirmModal').style.display === 'flex', false);
  assert.strictEqual(sent.length, 0);
  assert.match(status(w), /Nothing was imported.*HTTP 503/);
});

test('a failed batch names the file lines of its rows, even with skipped rows before it', async () => {
  const { w } = app([{ f_firstName: 'Ada', f_lastName: 'Test', f_dob: '' }]);
  load(w, ['First Name', 'Last Name'],
    [{ 'First Name': 'Ada', 'Last Name': 'Test' }, { 'First Name': 'Bo', 'Last Name': 'Test' }]);
  stub(w, { fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'boom' }) }) });

  w.importClients(); await settle(); confirmOk(w); await settle(); await settle();

  assert.match(status(w), /NOT imported: line 3 /, status(w));
});
