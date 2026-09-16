'use strict';
// The Access History was found by patient NAME. Two patients with the same name shared one history,
// and renaming a patient orphaned theirs — both much worse once ~1,000 records are imported. Every
// audit write now carries the patient id, and the history is searched by it.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, stub, formDom } = require('./harness');

const settle = () => new Promise((r) => setTimeout(r, 0));

function audited(w) {
  const posted = [];
  stub(w, { _postAuditRecord: (body) => posted.push(body) });
  return posted;
}

test('an audit write carries the patient id and the name at the time', () => {
  const w = loadApp();
  resetStorage(w);
  const posted = audited(w);

  w.addAuditEntry('Ada Test', 'Client record opened', 7);

  assert.strictEqual(posted.length, 1);
  assert.strictEqual(posted[0].client_id, 7);
  assert.strictEqual(posted[0].client_name, 'Ada Test',
    'the name is still recorded: it is what the record was called when this happened');
});

test('a global event (export, sign-out) has no patient id', () => {
  const w = loadApp();
  resetStorage(w);
  const posted = audited(w);

  w.logActivity('export', 'Exported 12 client records');
  w.addAuditEntry('Ada Test', 'Something with no id');

  assert.strictEqual(posted[0].client_id, undefined);
  assert.strictEqual(posted[1].client_id, '', 'an absent id must not become a patient id');
});

test('the history request asks by id, and still sends the name for older rows', async () => {
  const w = loadApp();
  resetStorage(w);
  let body = null;
  stub(w, { fetch: (url, opt) => {
    if (String(url).endsWith('/audit/search')) body = JSON.parse(opt.body);
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve([]) });
  } });
  w.document.body.insertAdjacentHTML('beforeend', '<div id="clientAuditSection"></div>');

  w.loadClientAudit('Ada Test', 7);
  await settle();

  assert.strictEqual(body.client_id, 7);
  assert.strictEqual(body.client, 'Ada Test');
  assert.strictEqual(body.scope, 'health');
});

test('opening a record files the access under that record', async () => {
  const w = loadApp();
  resetStorage(w);
  formDom(w);
  if (!w.document.getElementById('formTitle')) {
    w.document.body.insertAdjacentHTML('beforeend',
      '<div id="formTitle"></div><button id="deleteBtn"></button><button id="deleteBtn2"></button>' +
      '<div id="viewForm"><div class="form-card"><div class="form-actions"></div></div></div>');
  }
  const posted = audited(w);
  stub(w, {
    showView: () => {}, aiTrack: () => {}, trackRecentRecord: () => {}, loadCarriersToSelect: () => {},
    loadClients: () => {}, toast: () => {},
    fetch: () => Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve([]) }),
  });
  w.eval('clients=[{_id:4,f_firstName:"Alpha",f_lastName:"Test"}];editingId=null;_formDirty=false;');

  w.editClient(4);
  await settle();

  const open = posted.filter((p) => p.action === 'Client record opened');
  assert.strictEqual(open.length, 1);
  assert.strictEqual(String(open[0].client_id), '4');
});

test('a brand-new patient is filed under the id the server just assigned', async () => {
  const w = loadApp();
  resetStorage(w);
  formDom(w);
  const posted = audited(w);
  stub(w, {
    showView: () => {}, aiTrack: () => {}, loadClients: () => {}, toast: () => {}, clearFormDirty: () => {},
    saveClientAPI: () => Promise.resolve({ id: 42 }),   // create returns {id} with no row_version
  });
  w.eval('editingId=null;_fullRecordFailed=false;_fullRecordLoading=false;');
  w.document.getElementById('f_firstName').value = 'Newly';
  w.document.getElementById('f_lastName').value = 'Created';

  w.saveClient();
  await settle(); await settle();

  const created = posted.filter((p) => p.action === 'Client record created');
  assert.strictEqual(created.length, 1, JSON.stringify(posted));
  assert.strictEqual(created[0].client_id, 42,
    'a create has no editingId, so the id has to come from the save response');
});

test('deleting files the row under the deleted id, not just the name', async () => {
  const w = loadApp();
  resetStorage(w);
  const posted = audited(w);
  stub(w, {
    deleteClientAPI: () => Promise.resolve({}), loadClients: () => {}, showView: () => {},
    aiTrack: () => {}, toast: () => {}, currentUserEmail: () => 'x@y.z',
  });
  w.eval('clients=[{_id:4,f_firstName:"Alpha",f_lastName:"Test"}];editingId=4;');

  w.deleteClient();
  w.document.getElementById('confirmOkBtn').dispatchEvent(new w.Event('click'));
  await settle();

  const del = posted.filter((p) => /DELETED/.test(p.action));
  assert.strictEqual(del.length, 1);
  assert.strictEqual(String(del[0].client_id), '4');
});
