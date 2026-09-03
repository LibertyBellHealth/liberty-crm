'use strict';
// The client's status is packed into the notes column as a "[STATUS:X]" prefix — a deliberate
// design that avoids a schema change. The form unpacks it. Nothing else did, so every other
// reader of f_notes saw the marker.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, stub, formDom } = require('./harness');

function captureExport(w) {
  const captured = { rows: null, name: null };
  stub(w, {
    logActivity: () => {},
    XLSX: {
      utils: {
        aoa_to_sheet: (rows) => { captured.rows = rows; return {}; },
        book_new: () => ({}), book_append_sheet: () => {},
      },
      writeFile: (wb, name) => { captured.name = name; },
    },
  });
  return captured;
}

test('an exported Notes column carries the note, not the status marker', () => {
  const w = loadApp();
  resetStorage(w);
  const cap = captureExport(w);
  w.eval('clients=[{_id:1,f_firstName:"Ada",f_lastName:"Lovelace",' +
         'f_notes:"[STATUS:Pending]\\nPrefers morning calls."}];');

  w.exportFullBackup();

  const header = cap.rows[0], row = cap.rows[1];
  const notes = row[header.indexOf('Notes')];
  assert.strictEqual(notes, 'Prefers morning calls.',
    'the exported note still carried the internal status marker: ' + JSON.stringify(notes));
});

test('the advanced-search Notes column is clean too', () => {
  const w = loadApp();
  resetStorage(w);
  const c = { f_notes: '[STATUS:Terminated]\nMoved out of state.' };
  assert.strictEqual(w.reportFieldValue(c, 'notes'), 'Moved out of state.');
});

test('a note that merely mentions the marker later on is left alone', () => {
  const w = loadApp();
  const c = { f_notes: '[STATUS:Active]\nClient asked what [STATUS:Pending] means.' };
  assert.strictEqual(w.reportFieldValue(c, 'notes'),
    'Client asked what [STATUS:Pending] means.');
});

test('a client with no marker at all reads as Active with its note intact', () => {
  const w = loadApp();
  const c = { f_notes: 'Imported from spreadsheet.' };
  assert.strictEqual(w.reportFieldValue(c, 'notes'), 'Imported from spreadsheet.');
  assert.strictEqual(w.clientStatus(c), 'Active');
});

test('the form still shows the status and the clean note', () => {
  const w = loadApp();
  resetStorage(w);
  formDom(w);   // the real form's controls, from index.html

  w.setFormData({ f_notes: '[STATUS:Pending]\nPrefers morning calls.' });

  assert.strictEqual(w.document.getElementById('f_status').value, 'Pending');
  assert.strictEqual(w.document.getElementById('f_notes').value, 'Prefers morning calls.');
});
