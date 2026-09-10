'use strict';
// The same bug family as late-writes.test.js, but landing somewhere worse. The document and audit
// SECTIONS are rebuilt per record, so a late write there rendered wrong and stopped. The f_* form
// fields are STATIC elements in index.html, cleared and reused for every patient — and getFormData
// reads them straight back on save. A late write into one of those does not merely display the
// wrong thing: it puts one patient's data into another patient's record and persists it.
//
// Two paths did this. The ZIP -> city/state/county lookup took no record id at all, so nothing
// even had the information needed to guard it. And saveClient's success continuation re-read
// editingId to decide whether to clear the GLOBAL unsaved-changes flag.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, stub, formDom } = require('./harness');

const tick = () => new Promise((r) => setTimeout(r, 0));
const ZIPPO = {
  places: [{ 'place name': 'Detroit', 'state abbreviation': 'MI', latitude: '42.3', longitude: '-83.0' }],
};

function form(w) {
  formDom(w);
  ['f_resZip', 'f_resCity', 'f_resSt', 'f_firstName', 'f_lastName'].forEach((id) => {
    const el = w.document.getElementById(id);
    if (el) el.value = '';
  });
  const sel = w.document.getElementById('f_resCounty');
  if (sel) sel.innerHTML = '';
  return (id) => w.document.getElementById(id);
}
// A zip the bundled dataset does not know, so the lookup takes its fetch path.
const zipInput = (w, zip) => { const el = w.document.getElementById('f_resZip'); el.value = zip; return el; };

test("a ZIP lookup landing after you switch patients does not fill the other patient's address", async () => {
  const w = loadApp();
  resetStorage(w);
  const el = form(w);
  w.eval('editingId="A";ZIP_COUNTIES={};');
  let release;
  stub(w, { fetch: () => new Promise((res) => { release = () => res({ ok: true, json: () => Promise.resolve(ZIPPO) }); }) });

  w.lookupZip(zipInput(w, '48201'), 'res', 'A');   // started while A is open
  w.eval('editingId="B";');                         // operator opens another patient
  release();
  await tick(); await tick();

  assert.strictEqual(el('f_resCity').value, '',
    "A's city was written into the form now showing B — getFormData would save it onto B");
  assert.strictEqual(el('f_resSt').value, '');
});

test('the same lookup still fills the address when you have not moved on', async () => {
  const w = loadApp();
  resetStorage(w);
  const el = form(w);
  w.eval('editingId="A";ZIP_COUNTIES={};');
  let release;
  stub(w, { fetch: () => new Promise((res) => { release = () => res({ ok: true, json: () => Promise.resolve(ZIPPO) }); }) });

  w.lookupZip(zipInput(w, '48201'), 'res', 'A');
  release();
  await tick(); await tick();

  assert.strictEqual(el('f_resCity').value, 'Detroit', 'the guard must not break the feature');
  assert.strictEqual(el('f_resSt').value, 'MI');
});

// The zippopotam leg and the FCC leg are separate fetches. Holding BOTH means the county code
// never runs at all and the test passes whether or not the guard is there — so answer the FCC
// call immediately and hold only the first leg.
function zipStub(w) {
  let release;
  stub(w, { fetch: (url) => {
    if (String(url).indexOf('geo.fcc.gov') >= 0) {
      return Promise.resolve({ ok: true, json: () => Promise.resolve({ results: [{ county_name: 'Wayne' }] }) });
    }
    return new Promise((res) => { release = () => res({ ok: true, json: () => Promise.resolve(ZIPPO) }); });
  } });
  return () => release();
}

test("a county restore landing late does not repopulate the other patient's county list", async () => {
  const w = loadApp();
  resetStorage(w);
  const el = form(w);
  w.eval('editingId="A";ZIP_COUNTIES={};');
  const release = zipStub(w);

  w.restoreCounty('48201', 'res', 'Wayne', 'A');
  w.eval('editingId="B";');
  release();
  await tick(); await tick(); await tick();

  assert.ok(!/Wayne/.test(el('f_resCounty').innerHTML),
    "A's county options were written into the county field B is looking at");
});

test('a county restore for the record still on screen does populate it', async () => {
  const w = loadApp();
  resetStorage(w);
  const el = form(w);
  w.eval('editingId="A";ZIP_COUNTIES={};');
  const release = zipStub(w);

  w.restoreCounty('48201', 'res', 'Wayne', 'A');
  release();
  await tick(); await tick(); await tick();

  assert.ok(/Wayne/.test(el('f_resCounty').innerHTML),
    'the guard must not break the feature it protects');
});

test('opening a patient threads their id into the county lookup', async () => {
  const w = loadApp();
  resetStorage(w);
  const el = form(w);
  w.eval('editingId="A";ZIP_COUNTIES={};');
  const release = zipStub(w);

  // Through setFormData, the way editClient reaches it. If the id is not passed down, the guard
  // compares the open record against undefined, fails, and the county silently never fills.
  w.setFormData({ f_firstName: 'Ada', f_resZip: '48201', f_resCounty: 'Wayne' });
  release();
  await tick(); await tick(); await tick();

  assert.ok(/Wayne/.test(el('f_resCounty').innerHTML),
    "the open patient's county list did not populate — the record id is not reaching restoreCounty");
});

// ---- saveClient's continuation ----

function saveHarness(w, resolveWith) {
  formDom(w);
  w.document.getElementById('f_firstName').value = 'Ada';
  const toasts = [], views = [];
  let release;
  stub(w, {
    saveClientAPI: () => new Promise((res) => { release = () => res(resolveWith); }),
    loadClients: () => {},
    showView: (v) => views.push(String(v)),
    addAuditEntry: () => {},
    aiTrack: () => {},
    toast: (msg) => toasts.push(String(msg)),
  });
  return { toasts, views, release: () => release() };
}

test("a save resolving after you switch does not clear the new patient's unsaved-changes flag", async () => {
  const w = loadApp();
  resetStorage(w);
  const h = saveHarness(w, { row_version: 'aaaaaaaaaaaaaaaa' });
  w.eval('editingId="A";');

  w.saveClient();                 // save started for A
  w.eval('editingId="B";');       // operator opens B and starts typing
  w.markFormDirty();
  h.release();
  await tick(); await tick();

  assert.strictEqual(w._formDirty, true,
    "A's save cleared the global dirty flag, so B's unsaved edits would be discarded without warning");
});

test("a save resolving after you switch does not stamp its row version onto the next patient", async () => {
  const w = loadApp();
  resetStorage(w);
  const h = saveHarness(w, { row_version: 'aaaaaaaaaaaaaaaa' });
  w.eval('editingId="A";_rowVersion=null;');

  w.saveClient();
  w.eval('editingId="B";');
  h.release();
  await tick(); await tick();

  assert.strictEqual(w._rowVersion, null,
    "A's concurrency token became B's expected_version, which the server would reject as a conflict");
});

test('a late save still reports that it succeeded, naming the patient', async () => {
  const w = loadApp();
  resetStorage(w);
  const h = saveHarness(w, { row_version: 'aaaaaaaaaaaaaaaa' });
  w.eval('editingId="A";');

  w.saveClient();
  w.eval('editingId="B";');
  h.release();
  await tick(); await tick();

  assert.strictEqual(h.toasts.length, 1, 'the success report must never be suppressed: ' + JSON.stringify(h.toasts));
  assert.ok(/Ada/.test(h.toasts[0]), 'it should name who was saved, got: ' + h.toasts[0]);
});

test('a late save does not navigate you off the record you have since opened', async () => {
  const w = loadApp();
  resetStorage(w);
  const h = saveHarness(w, { row_version: 'aaaaaaaaaaaaaaaa' });
  w.eval('editingId="A";');

  w.saveClient();
  w.eval('editingId="B";');   // operator is now editing B
  h.release();
  await tick(); await tick();

  assert.deepStrictEqual(h.views, [],
    "A's save sent the operator back to the client list, mid-edit on B — navigation is a write too");
});

// The discovery paste opens a BLANK record and then populates it 80ms later, to let the view and
// the starter rows render first. The whole timer body writes into the static f_* fields. 80ms is
// long enough to reach an existing patient by Back button or deep link — and the paste would then
// land in THEIR form and mark it dirty, ready to be saved over them.
test('a discovery paste does not land in a patient you opened while it was pending', async () => {
  const w = loadApp();
  resetStorage(w);
  const el = form(w);
  const toasts = [];
  stub(w, {
    startNewApp: () => { w.eval('editingId=null;'); },
    closeDiscoveryPasteModal: () => {},
    markFormDirty: () => {},
    updateMemberCount: () => {},
    toast: (m) => toasts.push(String(m)),
  });
  if (!w.document.getElementById('discoveryPasteInput')) {
    w.document.body.insertAdjacentHTML('beforeend', '<textarea id="discoveryPasteInput"></textarea>');
  }
  w.document.getElementById('discoveryPasteInput').value = 'Name- Ada Lovelace\nDOB- 1815-12-10';

  w.importDiscoveryPaste();          // blank record opened, paste scheduled
  w.eval('editingId="B";');          // operator navigates to a real patient
  await new Promise((r) => setTimeout(r, 120));

  assert.strictEqual(el('f_firstName').value, '',
    "the pasted intake data was written into patient B's form and would save onto them");
  assert.ok(toasts.some((t) => /different record/i.test(t)),
    'discarding the paste silently is its own failure — say so: ' + JSON.stringify(toasts));
});
