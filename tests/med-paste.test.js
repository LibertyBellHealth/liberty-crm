'use strict';
// Medication and doctor lists get pasted from wherever the agent has them — a discharge summary,
// a text message, a note the client typed. The parsers have to survive ordinary list formatting.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, jsonEqual } = require('./harness');

test('a numbered list still yields a dose', () => {
  const w = loadApp();
  // The leading "1" was being matched as the dose. With nothing before it the parser gave up and
  // returned the whole line as the NAME — so the medication was recorded with no dose at all.
  assert.ok(jsonEqual(w.parseMedLine('1. Metformin 500mg'),
    { name: 'Metformin', mg: '500mg', frequency: '' }));
  assert.ok(jsonEqual(w.parseMedLine('2) Lisinopril 10mg twice daily'),
    { name: 'Lisinopril', mg: '10mg', frequency: 'twice daily' }));
});

test('bullet characters do not end up in the medication name', () => {
  const w = loadApp();
  assert.strictEqual(w.parseMedLine('- Atorvastatin 20mg').name, 'Atorvastatin');
  assert.strictEqual(w.parseMedLine('• Aspirin 81mg').name, 'Aspirin');
  assert.strictEqual(w.parseMedLine('* Losartan 50mg daily').name, 'Losartan');
});

test('a medication that legitimately starts with a number is left alone', () => {
  const w = loadApp();
  assert.strictEqual(w.parseMedLine('5-HTP 100mg').name, '5-HTP');
});

test('list markers do not end up in a doctor name', () => {
  const w = loadApp();
  assert.ok(jsonEqual(w.parseDoctorLine('1. Dr. Smith, Cardiology'),
    { name: 'Dr. Smith', specialty: 'Cardiology' }));
  assert.ok(jsonEqual(w.parseDoctorLine('- Dr. Jones'), { name: 'Dr. Jones', specialty: '' }));
});

// ── Two medications on one line vs one medication and its frequency. Same punctuation, opposite
// meanings — the difference is whether what follows the comma carries its own dose.
test('two medications on one line are kept as two', () => {
  const w = loadApp();
  assert.ok(jsonEqual(w._splitMedLine('Metformin 500mg, Lisinopril 10mg'),
    ['Metformin 500mg', 'Lisinopril 10mg']),
    'the second medication was absorbed into the first as its frequency');
});

test('a medication and its frequency stay one medication', () => {
  const w = loadApp();
  assert.ok(jsonEqual(w._splitMedLine('Metformin 500mg, twice daily'), ['Metformin 500mg, twice daily']),
    'a frequency was split off into a medication of its own');
  assert.ok(jsonEqual(w._splitMedLine('Metformin 500mg, take 2 tablets'), ['Metformin 500mg, take 2 tablets']));
});

test('the discovery paste applies the same rule, both ways', () => {
  const w = loadApp();
  // Used to split on every comma, so "twice daily" became a medication in its own right.
  const one = w.parseDiscoveryText('Rx- Metformin 500mg, twice daily');
  assert.strictEqual(one.meds.length, 1, 'a frequency was imported as a medication: ' + JSON.stringify(one.meds));
  assert.strictEqual(one.meds[0].frequency, 'twice daily');

  const two = w.parseDiscoveryText('Rx- Metformin 500mg, Lisinopril 10mg');
  assert.strictEqual(two.meds.length, 2);
  assert.strictEqual(two.meds[1].name, 'Lisinopril');
});

test('a dose with a unit wins over a number in the name', () => {
  const w = loadApp();
  assert.ok(jsonEqual(w.parseMedLine('Vitamin D3 5000 IU'),
    { name: 'Vitamin D3', mg: '5000iu', frequency: '' }));
  assert.ok(jsonEqual(w.parseMedLine('Insulin 70/30 10 units'),
    { name: 'Insulin 70/30', mg: '10units', frequency: '' }));
});

test('a bare number is still read as the dose when no unit appears at all', () => {
  const w = loadApp();
  // Deliberate, and asked for: "Metformin 30 twice" means 30mg twice.
  assert.ok(jsonEqual(w.parseMedLine('Metformin 30 twice'),
    { name: 'Metformin', mg: '30mg', frequency: 'twice' }));
});
