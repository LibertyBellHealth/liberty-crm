'use strict';
// Regression tests for the 2026-09-03 Health CRM audit. Each test here was written to FAIL
// against the code as it stood, and each was mutation-checked afterwards by re-breaking the fix.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, confirmOk, stub } = require('./harness');

const tick = () => new Promise(r => setTimeout(r, 0));
function toasts(w) {
  return Array.from(w.document.querySelectorAll('#toastContainer .toast')).map(t => t.textContent);
}

// ── A document's filename is data from outside the app: it arrives on a file the agent was
// sent, and the backend only strips path separators and whitespace from it. It reaches the
// delete button as a JavaScript string literal inside an onclick attribute, and
// encodeURIComponent leaves ' . ( ) untouched — enough to close the literal and call something.
test('a document filename cannot execute as JavaScript from the delete button', () => {
  const w = loadApp();
  resetStorage(w);
  w.eval('window.__pwned=false; window.__pwn=function(){window.__pwned=true;return "";};');
  w.__delArgs = null;
  stub(w, { deleteClientDoc: (a, b) => { w.__delArgs = [a, b]; } });

  // encodeURIComponent passes ' . ( ) through untouched, so this name closes the string
  // literal, calls __pwn(), and balances the call's parentheses again. Any zero-argument
  // function in scope could stand in its place — deleteClient() among them.
  const evil = "x'.concat(__pwn())).concat('y.pdf";
  w.renderClientDocs('42', [{ name: evil, size: 2048, url: 'https://example.invalid/x.pdf' }]);

  const btn = w.document.querySelector('#clientDocsSection button');
  assert.ok(btn, 'expected a delete button for the document');
  // Exercise whichever mechanism ships — an inline attribute or a real listener.
  const onclick = btn.getAttribute('onclick');
  // The tail of an injected expression may throw once the payload has already run; the throw is
  // not the point, __pwned is.
  try {
    if (onclick) w.eval(onclick);
    else btn.dispatchEvent(new w.Event('click'));
  } catch (e) { /* see above */ }

  assert.strictEqual(w.__pwned, false, 'the filename executed as JavaScript');
  assert.ok(w.__delArgs, 'the delete handler did not run at all');
  assert.strictEqual(decodeURIComponent(w.__delArgs[1]), evil, 'the filename did not round-trip');
});

// ── Clearing the status line and the file input is the only "upload finished" signal the agent
// gets. fetch() resolves on 500 and 413 just as it does on 201, so that signal must not be
// driven by the request completing — it has to be driven by the request succeeding.
test('a failed document upload is reported as a failure', async () => {
  const w = loadApp();
  resetStorage(w);
  w.renderClientDocs('42', []);
  const input = w.document.getElementById('docFileInput');
  Object.defineProperty(input, 'files', { value: [new w.File(['x'], 'card.pdf')], configurable: true });
  w.__loadCalled = false;
  stub(w, {
    loadClientDocs: () => { w.__loadCalled = true; },
    fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'boom' }) }),
  });

  w.uploadClientDoc('42');
  await tick(); await tick();

  const status = w.document.getElementById('docUploadStatus').textContent;
  assert.match(status + ' ' + toasts(w).join(' '), /fail/i, 'a 500 upload left no failure message');
  assert.strictEqual(w.__loadCalled, false, 'a failed upload still refreshed the list as if it worked');
});

// ── Same shape on delete: the list is refreshed either way, so a refused delete looks like a
// successful one that simply left the file in place.
test('a failed document delete is reported as a failure', async () => {
  const w = loadApp();
  resetStorage(w);
  w.__loadCalled = false;
  stub(w, {
    loadClientDocs: () => { w.__loadCalled = true; },
    fetch: () => Promise.resolve({ ok: false, status: 403, json: () => Promise.resolve({ error: 'nope' }) }),
  });

  w.deleteClientDoc('42', encodeURIComponent('card.pdf'));
  confirmOk(w);
  await tick(); await tick();

  assert.match(toasts(w).join(' '), /fail|could not/i, 'a 403 delete left no failure message');
});

// ── The roster is the app's whole view of who exists. A load that fails must not look like a
// load that returned nothing, and must not leave the previous roster on screen unlabelled.
test('a failed client load is surfaced, not swallowed', async () => {
  const w = loadApp();
  resetStorage(w);
  w.eval('clients=[{_id:1,f_firstName:"Ada",f_lastName:"L"}];');
  stub(w, { fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'db down' }) }) });

  w.loadClients();
  await tick(); await tick();

  const said = toasts(w).join(' ');
  assert.match(said, /could not|fail/i, 'a 500 client load said nothing at all');
  // Not just "some message": the server's own reason has to survive. Without the status check the
  // error body parses as JSON, `data.map` throws, and the agent is told "data.map is not a
  // function" — which reads as a bug in the page rather than an unavailable server.
  assert.match(said, /db down|500/i, 'the message did not name the actual failure');
  assert.strictEqual(w.clients.length, 1, 'a failed load discarded the roster already in memory');
});

// ── "No PHI in localStorage, ever." crm_recent and crm_todos were both fixed for this; the
// audit-log mirror kept doing it, 200 entries deep, and nothing ever read it back.
test('the audit trail keeps patient names out of localStorage', () => {
  const w = loadApp();
  resetStorage(w);
  w.eval('_apiToken="fake";');
  w.__posted = null;
  stub(w, { _postAuditRecord: (b) => { w.__posted = b; } });

  w.addAuditEntry('Jane Patient', 'Client record opened');

  const stored = w.localStorage.getItem('crm_audit') || '';
  assert.doesNotMatch(stored, /Jane|Patient/, 'a patient name was written to localStorage');
  assert.ok(w.__posted && /Jane Patient/.test(w.__posted.client_name),
    'the name must still reach the server-side audit trail');
});

// ── Stopping the writes does nothing for the browsers that already have months of names sitting
// in them. The crm_recent fix migrated on load for exactly this reason.
test('legacy audit entries already in localStorage are purged on load', () => {
  const w = loadApp();
  resetStorage(w);
  w.localStorage.setItem('crm_audit', JSON.stringify([{ client: 'Jane Patient', action: 'opened' }]));

  w.purgeLegacyAuditLog();

  assert.doesNotMatch(w.localStorage.getItem('crm_audit') || '', /Jane|Patient/);
});

// ── The confirm dialog is not modal to the browser. `hashchange` fires on the Back button with
// no click on the page at all, and routeFromHash → editClient reassigns editingId behind the
// open dialog. Anything that reads app state INSIDE the callback is reading it at OK-press time,
// not at the moment the user was shown what they were agreeing to.
test('deleting a client deletes the record the dialog named', () => {
  const w = loadApp();
  resetStorage(w);
  const deleted = [], audited = [];
  w.eval('clients=[{_id:"A",f_firstName:"Ada",f_lastName:"Lovelace"},' +
         '{_id:"B",f_firstName:"Bo",f_lastName:"Mira"}]; editingId="B";');
  stub(w, {
    deleteClientAPI: (id) => { deleted.push(String(id)); return Promise.resolve(); },
    addAuditEntry: (name) => { audited.push(name); },
    loadClients: () => {}, showView: () => {}, aiTrack: () => {},
  });

  w.deleteClient();                 // dialog says "Delete Bo Mira?"
  w.eval('editingId="A";');         // Back button re-routes the hash to the other client
  confirmOk(w);

  assert.deepStrictEqual(deleted, ['B'], 'deleted a different client than the dialog named');
  assert.deepStrictEqual(audited, ['Bo Mira'], 'the audit row named a different client than was deleted');
});

// ── Same shape, by index instead of id: the settings lists are spliced at a position captured
// before the dialog opened.
test('removing a settings entry removes the one that was named', () => {
  const w = loadApp();
  resetStorage(w);
  stub(w, { fetch: () => Promise.resolve({ ok: true, json: () => Promise.resolve({}) }) });
  w.eval('_settingsAgents=["Ann","Ben","Cal"];');

  w.removeAgentSetting(2);          // dialog opened for "Cal"
  w.eval('_settingsAgents.unshift("Zed");');  // the list shifts underneath it
  confirmOk(w);

  assert.ok(!w._settingsAgents.includes('Cal'), 'the named entry survived');
  assert.ok(w._settingsAgents.includes('Ben'), 'a different entry was removed instead');
});
