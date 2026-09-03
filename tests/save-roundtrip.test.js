'use strict';
// Opening a client and saving it without editing anything must not change the record. This is the
// guard the homecare_client_id bug slipped past: a column that defaults to a real value instead of
// being omitted overwrites whatever the server holds, on every single save.
//
// Modelled on a replay of the four real production records through the same path; this version
// uses synthetic values so it carries no patient data.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, formDom } = require('./harness');

// A value for every field on the form, distinctive enough that a mix-up is visible.
function syntheticClient(w) {
  const c = {};
  w.FIELDS.forEach((f, i) => {
    const el = w.document.getElementById('f_' + f);
    if (!el) return;                       // no input — getFormData can never produce it
    if (el.type === 'checkbox') c['f_' + f] = true;
    else if (el.type === 'date') c['f_' + f] = '2026-0' + ((i % 9) + 1) + '-1' + (i % 9);
    else if (el.tagName === 'SELECT') { const o = el.querySelector('option'); c['f_' + f] = o ? o.value : ''; }
    else c['f_' + f] = 'v' + i;
    return;
  });
  return c;
}

test('opening a fully populated client and saving it changes nothing', () => {
  const w = loadApp();
  resetStorage(w);
  formDom(w);

  const before = syntheticClient(w);
  w.clearForm();
  w.setFormData(before);
  const after = w.getFormData();

  const drifted = [];
  // Total Monthly is DERIVED — setFormData recomputes it from the premium and the ancillary rows,
  // so it is expected not to echo an arbitrary input back. Asserted separately below rather than
  // quietly skipped. (It did not move on any of the four real production records.)
  const COMPUTED = ['f_totalMonthly'];
  Object.keys(before).forEach((k) => {
    if (k === 'f_notes') return;           // deliberately re-packed with the status marker
    if (COMPUTED.indexOf(k) !== -1) return;
    const a = before[k], b = after[k];
    const norm = (v) => (v === true ? 'true' : v === false ? 'false' : String(v == null ? '' : v).trim());
    if (norm(a) !== norm(b)) drifted.push(k + ': ' + JSON.stringify(a) + ' -> ' + JSON.stringify(b));
  });
  assert.deepStrictEqual(drifted, [], 'fields changed on a no-edit save:\n  ' + drifted.join('\n  '));
});

test('Total Monthly is recomputed from the plan rows, not echoed back', () => {
  const w = loadApp();
  resetStorage(w);
  formDom(w);
  w.clearForm();
  // A stored total that does not match the parts must not survive a round trip — the point of a
  // derived field is that the parts win.
  w.setFormData({ f_premium: '100.00', f_totalMonthly: '999.99' });
  assert.notStrictEqual(w.getFormData().f_totalMonthly, '999.99');
});

test('a no-edit save omits every column the form cannot produce', () => {
  const w = loadApp();
  resetStorage(w);
  formDom(w);

  w.clearForm();
  w.setFormData(syntheticClient(w));
  // JSON round-trip, because that is what decides what the server actually receives:
  // undefined keys disappear, and the backend leaves any column it was not sent alone.
  const sent = JSON.parse(JSON.stringify(w.clientToDbRow(w.getFormData())));

  // No input exists for either of these, so neither may appear in the payload.
  assert.ok(!('waive_dental' in sent), 'waive_dental would overwrite the stored value');
  assert.ok(!('homecare_client_id' in sent), 'homecare_client_id would unlink the Home Care record');
});

test('the status marker is the only thing a no-edit save adds to the notes', () => {
  const w = loadApp();
  resetStorage(w);
  formDom(w);

  w.clearForm();
  w.setFormData({ f_notes: '[STATUS:Pending]\nPrefers morning calls.' });
  const out = w.getFormData();

  assert.strictEqual(out.f_notes, '[STATUS:Pending]\nPrefers morning calls.',
    'a record round-tripped through the form came back with different notes');
});
