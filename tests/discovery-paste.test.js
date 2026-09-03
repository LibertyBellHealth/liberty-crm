'use strict';
// The discovery paste takes free text typed during a call and turns it into a client record,
// its medication list and its doctor list. Anything it drops is data the agent believes they
// captured.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, jsonEqual } = require('./harness');

test('a medication written with a dash does not swallow the rest of the list', () => {
  const w = loadApp();
  const out = w.parseDiscoveryText([
    'Name- Ada Lovelace',
    'Rx-',
    'Metformin 500mg - twice daily',
    'Lisinopril 10mg',
    'Atorvastatin 20mg - at night',
  ].join('\n'));

  // "Metformin 500mg - twice daily" matches the Label-value pattern, so it used to be read as a
  // LABEL, discarded, and — because an unknown label clears the collector — every line after it
  // was discarded too. Three medications became zero, silently.
  assert.strictEqual(out.meds.length, 3, 'medications were dropped: ' + JSON.stringify(out.meds));
  assert.strictEqual(out.meds[0].name, 'Metformin');
  assert.strictEqual(out.meds[0].mg, '500mg');
  assert.strictEqual(out.meds[0].frequency, 'twice daily');
  assert.strictEqual(out.meds[1].name, 'Lisinopril');
  assert.strictEqual(out.meds[2].frequency, 'at night');
  assert.strictEqual(out.data.f_firstName, 'Ada');
});

test('a doctor written "Name - Specialty" keeps both, and the list continues', () => {
  const w = loadApp();
  const out = w.parseDiscoveryText([
    'Primary Dr-',
    'Dr. Smith - Cardiology',
    'Dr. Jones',
    'Specialist-',
    'Dr. Patel - Endocrinology',
  ].join('\n'));

  assert.strictEqual(out.doctors.length, 3, 'doctors were dropped: ' + JSON.stringify(out.doctors));
  assert.ok(jsonEqual(out.doctors[0], { name: 'Dr. Smith', specialty: 'Cardiology' }));
  // No specialty on the line — fall back to the section it was listed under.
  assert.ok(jsonEqual(out.doctors[1], { name: 'Dr. Jones', specialty: 'Primary' }));
  assert.ok(jsonEqual(out.doctors[2], { name: 'Dr. Patel', specialty: 'Endocrinology' }));
});

test('a real label still ends the list it follows', () => {
  const w = loadApp();
  const out = w.parseDiscoveryText([
    'Rx-',
    'Metformin 500mg',
    'DOB- 05/14/1961',
    'Zip- 48226',
  ].join('\n'));

  assert.strictEqual(out.meds.length, 1, 'a real label was swallowed into the medication list');
  assert.strictEqual(out.data.f_dob, '1961-05-14');
  assert.strictEqual(out.data.f_resZip, '48226');
});

test('a section header still ends the list and switches person', () => {
  const w = loadApp();
  const out = w.parseDiscoveryText([
    'Rx-',
    'Metformin 500mg',
    'Spouse-',
    'DOB- 02/02/1962',
  ].join('\n'));

  assert.strictEqual(out.meds.length, 1);
  assert.strictEqual(out.members.length, 1);
  assert.strictEqual(out.members[0].relation, 'Spouse');
  assert.strictEqual(out.members[0].dob, '1962-02-02');
});
