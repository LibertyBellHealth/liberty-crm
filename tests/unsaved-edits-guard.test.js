'use strict';
// Found on the dev backend: the top search bar, the Recent list and Back/Forward all called
// editClient directly, skipping the unsaved-changes prompt that the sidebar gets through showView.
// Typed edits were discarded with no warning. Opening a record also opened it twice (editClient
// sets the hash, hashchange reopened it), logging two ePHI accesses.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, stub, formDom } = require('./harness');

const settle = () => new Promise((r) => setTimeout(r, 0));

function setup(w) {
  formDom(w);
  if (!w.document.getElementById('formTitle')) {
    w.document.body.insertAdjacentHTML('beforeend',
      '<div id="formTitle"></div><button id="deleteBtn"></button><button id="deleteBtn2"></button>' +
      '<div id="viewForm"><div class="form-card"><div class="form-actions"></div></div></div>');
  }
  w.document.getElementById('viewForm').style.display = 'block';
  const opened = [], audits = [];
  stub(w, {
    showView: () => {},
    aiTrack: () => {},
    addAuditEntry: (name, action) => audits.push(action),
    trackRecentRecord: () => {},
    loadCarriersToSelect: () => {},
    loadClients: () => {},
    toast: () => {},
    fetch: (url) => {
      const m = String(url).match(/\/health-clients\/(\w+)$/);
      if (m) opened.push(m[1]);
      return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(m ? { id: +m[1], row_version_hex: 'aa' } : []) });
    },
  });
  w.eval('clients=[{_id:4,f_firstName:"Alpha",f_lastName:"Test"},{_id:5,f_firstName:"Bravo",f_lastName:"Test"}];');
  return { opened, audits };
}
const modalOpen = (w) => w.document.getElementById('confirmModal').style.display === 'flex';
const click = (w, id) => w.document.getElementById(id).dispatchEvent(new w.Event('click'));

async function openAlphaAndType(w) {
  w.editClient(4);
  await settle(); await settle();
  w.document.getElementById('f_mi').value = 'Q';
  w.markFormDirty();
}

test('opening another record with unsaved edits asks first, and Stay keeps them', async () => {
  const w = loadApp();
  resetStorage(w);
  const h = setup(w);
  await openAlphaAndType(w);
  h.opened.length = 0;

  w.editClient(5);   // the search bar and Recent list call this directly

  assert.ok(modalOpen(w), 'no unsaved-changes prompt');
  assert.strictEqual(String(w.editingId), '4');
  assert.deepStrictEqual(h.opened, [], 'the other record was fetched before the operator chose');
  click(w, 'confirmCancelBtn');
  assert.strictEqual(w.document.getElementById('f_mi').value, 'Q');
  assert.strictEqual(w._formDirty, true);
  assert.strictEqual(String(w.editingId), '4');
});

test('Discard & Leave opens the other record', async () => {
  const w = loadApp();
  resetStorage(w);
  const h = setup(w);
  await openAlphaAndType(w);

  w.editClient(5);
  click(w, 'confirmOkBtn');
  await settle();

  assert.strictEqual(String(w.editingId), '5');
  assert.strictEqual(w.document.getElementById('f_firstName').value, 'Bravo');
  assert.ok(h.opened.includes('5'));
});

test('Back/Forward with unsaved edits asks first, and Stay puts the URL back', async () => {
  const w = loadApp();
  resetStorage(w);
  setup(w);
  await openAlphaAndType(w);

  w.history.pushState(null, '', '#/client/5');
  w.routeFromHash();

  assert.ok(modalOpen(w), 'Back/Forward skipped the unsaved-changes prompt');
  click(w, 'confirmCancelBtn');
  assert.strictEqual(w.location.hash, '#/client/4');
  assert.strictEqual(String(w.editingId), '4');
});

test('opening a record logs one access and fetches it once', async () => {
  const w = loadApp();
  resetStorage(w);
  const h = setup(w);
  w.eval('editingId=null;_formDirty=false;');
  w.history.replaceState(null, '', '#');

  w.editClient(4);
  // editClient sets the hash, which fires hashchange → routeFromHash.
  await new Promise((r) => setTimeout(r, 50));

  assert.deepStrictEqual(h.opened, ['4']);
  assert.strictEqual(h.audits.filter((a) => a === 'Client record opened').length, 1);
});
