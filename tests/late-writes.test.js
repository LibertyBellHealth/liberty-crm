'use strict';
// The bug family: a global says which record is open (here `editingId`), an async gap sits between
// deciding and writing, and the DOM the write lands in is REUSED between records. So a response
// that started for client A arrives after you have navigated to client B, and A's data renders
// under B's name.
//
// editClient already guards the main form fetch this way. The two per-client SECTIONS did not.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, stub } = require('./harness');

const tick = () => new Promise((r) => setTimeout(r, 0));

function sections(w) {
  if (!w.document.getElementById('clientDocsSection')) {
    const d = w.document.createElement('div');
    d.innerHTML = '<div id="clientDocsSection"></div><div id="clientAuditSection"></div>';
    w.document.body.appendChild(d);
  }
  return {
    docs: w.document.getElementById('clientDocsSection'),
    audit: w.document.getElementById('clientAuditSection'),
  };
}

test("another client's documents do not render under the record now on screen", async () => {
  const w = loadApp();
  resetStorage(w);
  const el = sections(w);
  w.eval('editingId="A";');
  let release;
  stub(w, { fetch: () => new Promise((res) => { release = () => res({
    ok: true, status: 200,
    json: () => Promise.resolve([{ name: 'Ada-Lovelace-SSN-card.pdf', size: 100, url: 'https://x/y' }]),
  }); }) });

  w.loadClientDocs('A');            // starts loading A's documents
  w.eval('editingId="B";');         // the agent opens a different client
  release();                        // A's response arrives late
  await tick(); await tick();

  assert.ok(!/Ada-Lovelace/.test(el.docs.textContent),
    "client A's document list rendered while client B was on screen");
});

test("another client's access history does not render under the record now on screen", async () => {
  const w = loadApp();
  resetStorage(w);
  const el = sections(w);
  w.eval('editingId="A";');
  let release;
  stub(w, { fetch: () => new Promise((res) => { release = () => res({
    ok: true, status: 200,
    json: () => Promise.resolve([{ client_name: 'Ada Lovelace', action: 'Client record opened', who: 'x@y', ts: '2026-09-03' }]),
  }); }) });

  w.loadClientAudit('Ada Lovelace', 'A');
  w.eval('editingId="B";');
  release();
  await tick(); await tick();

  assert.ok(!/Ada Lovelace|record opened/.test(el.audit.textContent),
    "client A's access history rendered while client B was on screen");
});

test('a failed load also stays off the record now on screen', async () => {
  const w = loadApp();
  resetStorage(w);
  const el = sections(w);
  w.eval('editingId="A";');
  let release;
  // A REJECTION, not a 500: an HTTP error still resolves and json() parses the error body, so a
  // 500 exercises the success path. Only a network failure reaches the catch — which is the
  // branch that writes the "Could not load documents" markup into the shared section.
  stub(w, { fetch: () => new Promise((res, rej) => { release = () => rej(new Error('network down')); }) });

  w.loadClientDocs('A');
  w.eval('editingId="B";');
  // Whatever is on screen now belongs to B. The late response must not touch it AT ALL — asserting
  // on the absence of one error string passed even before the fix, because the failure path
  // produced different text.
  el.docs.innerHTML = '<p id="belongs-to-B">B\u2019s section</p>';
  release();
  await tick(); await tick();

  assert.ok(w.document.getElementById('belongs-to-B'),
    "a late response from client A overwrote the section belonging to client B");
});

test('the normal case still renders — the guard must not break loading', async () => {
  const w = loadApp();
  resetStorage(w);
  const el = sections(w);
  w.eval('editingId="A";');
  stub(w, { fetch: () => Promise.resolve({
    ok: true, status: 200,
    json: () => Promise.resolve([{ name: 'card.pdf', size: 100, url: 'https://x/y' }]),
  }) });

  w.loadClientDocs('A');
  await tick(); await tick();

  assert.match(el.docs.textContent, /card\.pdf/, 'the document list stopped rendering at all');
});

test('an upload finishing after you move on writes to the toast, not the other client', async () => {
  const w = loadApp();
  resetStorage(w);
  sections(w);
  w.eval('editingId="A";');
  w.renderClientDocs('A', []);      // builds #docFileInput and #docUploadStatus for client A
  const status = w.document.getElementById('docUploadStatus');
  const input = w.document.getElementById('docFileInput');
  Object.defineProperty(input, 'files', { value: [new w.File(['x'], 'card.pdf')], configurable: true });

  let release;
  stub(w, {
    loadClientDocs: () => {},
    fetch: () => new Promise((res) => { release = () => res({
      ok: false, status: 500, json: () => Promise.resolve({ error: 'boom' }),
    }); }),
  });

  w.uploadClientDoc('A');
  w.eval('editingId="B";');         // the agent opens a different client mid-upload
  status.textContent = "B's status line";
  release();
  await tick(); await tick(); await tick();

  assert.strictEqual(status.textContent, "B's status line",
    "client A's upload result was written into the status line client B is looking at");
  // The toast is global and must still fire — a failed upload is not something to hide because
  // the agent navigated away.
  const toasts = Array.from(w.document.querySelectorAll('#toastContainer .toast')).map((t) => t.textContent);
  assert.match(toasts.join(' '), /upload failed/i,
    'the agent was never told the upload failed');
});
