'use strict';
// Opening a record fills the form from the LIST row first so it isn't blank while the detail fetch
// is in flight. That list row deliberately omits ssn, card, routing and account — the sensitive
// columns are only returned by GET /health-clients/{id}. _fullRecordFailed covered the case where
// that fetch FAILS, but nothing covered the window while it is still in flight: a save there wrote
// the list row's blanks over real stored values, and went out with no expected_version because
// _rowVersion had not arrived either, so it skipped the lost-update check as well.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, stub, formDom } = require('./harness');

const tick = () => new Promise((r) => setTimeout(r, 0));

function open(w) {
  formDom(w);
  if (!w.document.getElementById('formTitle')) {
    // editClient inserts the documents and audit sections before .form-actions inside the form
    // card, and re-routes the hash — which fires routeFromHash, i.e. the Back-button path this
    // whole guard exists for. Give it the real shape so none of that is stubbed away.
    w.document.body.insertAdjacentHTML('beforeend',
      '<div id="formTitle"></div><button id="deleteBtn"></button><button id="deleteBtn2"></button>' +
      '<div id="viewForm"><div class="form-card"><div class="form-actions"></div></div></div>');
  }
  ['clientDocsSection', 'clientAuditSection', 'clientTodoSection'].forEach((id) => {
    const n = w.document.getElementById(id);
    if (n && n.parentNode) n.parentNode.removeChild(n);
  });
  w.document.getElementById('f_firstName').value = 'Ada';
  const saves = [], toasts = [];
  let release, fail;
  stub(w, {
    clients: undefined,
    showView: () => {},
    aiTrack: () => {},
    addAuditEntry: () => {},
    trackRecentRecord: () => {},
    loadCarriersToSelect: () => {},
    loadClients: () => {},
    saveClientAPI: (data, id) => { saves.push(id); return Promise.resolve({ row_version: 'bbbbbbbbbbbbbbbb' }); },
    toast: (m) => toasts.push(String(m)),
    // Only the DETAIL fetch is held. editClient also kicks off the documents and audit loads, and
    // handing all three the same held promise meant `release` pointed at whichever ran last — the
    // detail fetch never resolved and the test passed for the wrong reason.
    fetch: (url) => {
      if (String(url).indexOf('/health-clients/') >= 0) {
        return new Promise((res) => {
          release = () => res({ ok: true, json: () => Promise.resolve({ id: 7, first_name: 'Ada', row_version_hex: 'aaaaaaaaaaaaaaaa' }) });
          fail = () => res({ ok: false, status: 500, json: () => Promise.resolve({}) });
        });
      }
      return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve([]) });
    },
  });
  w.eval('clients=[{_id:7,f_firstName:"Ada",f_lastName:"Lovelace"}];');
  return { saves, toasts, release: () => release(), fail: () => fail() };
}

test('saving while the full record is still loading is refused', async () => {
  const w = loadApp();
  resetStorage(w);
  const h = open(w);

  w.editClient(7);          // detail fetch is in flight
  w.saveClient();
  await tick();

  assert.deepStrictEqual(h.saves, [],
    'the save went out while the form still held the list row, which has no SSN, card or bank values');
  assert.ok(h.toasts.some((t) => /loading/i.test(t)), 'the operator should be told why: ' + JSON.stringify(h.toasts));
});

test('saving works once the full record has arrived', async () => {
  const w = loadApp();
  resetStorage(w);
  const h = open(w);

  w.editClient(7);
  h.release();
  await tick(); await tick();
  w.saveClient();
  await tick();

  assert.strictEqual(h.saves.length, 1, 'the block must lift once the real values are on screen');
});

test('a failed detail fetch still blocks the save', async () => {
  const w = loadApp();
  resetStorage(w);
  const h = open(w);

  w.editClient(7);
  h.fail();
  await tick(); await tick();
  w.saveClient();
  await tick();

  assert.deepStrictEqual(h.saves, [], 'saving after a failed load would write blanks over stored SSN');
});

test('starting a new record clears the loading block', async () => {
  const w = loadApp();
  resetStorage(w);
  const h = open(w);

  w.editClient(7);           // leaves _fullRecordLoading set, fetch never resolves
  w.startNewApp('health');   // operator abandons it and starts a fresh record
  w.document.getElementById('f_firstName').value = 'Grace';
  w.saveClient();
  await tick();

  assert.strictEqual(h.saves.length, 1,
    'a brand-new record has nothing to load and must not inherit the previous block');
});
