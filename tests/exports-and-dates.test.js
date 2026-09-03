'use strict';
// Michigan is UTC-4 (EDT) / UTC-5 (EST), so from 8pm local until midnight, toISOString()
// already reports TOMORROW. Anything comparing a local calendar date against a UTC "today"
// is wrong for those hours every single day.
process.env.TZ = 'America/Detroit';

const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, stub } = require('./harness');

// 2026-09-03 21:30 EDT === 2026-09-04 01:30 UTC. Local date and UTC date disagree.
const EVENING = new Date('2026-09-04T01:30:00.000Z');

function freezeClock(w) {
  const Real = w.Date;
  class Frozen extends Real {
    constructor(...args) { if (!args.length) return new Real(EVENING.getTime()); return new Real(...args); }
    static now() { return EVENING.getTime(); }
  }
  stub(w, { Date: Frozen });
}

function todoDom(w) {
  if (!w.document.getElementById('todoListContainer')) {
    const d = w.document.createElement('div');
    d.innerHTML = '<div id="todoListContainer"></div><div id="todoEmpty"></div>';
    w.document.body.appendChild(d);
  }
}

test('a task due today is not shown as overdue during the evening', () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  freezeClock(w);
  // What the user sees on a local calendar right now.
  assert.strictEqual(w.fmtToday(), '2026-09-03', 'fmtToday should be the LOCAL date');

  w.eval('_todos=[{id:1,dbId:1,task:"Call Ada",done:false,priority:"normal",due:"2026-09-03",clientId:""}];');
  w.renderTodos();
  const html = w.document.getElementById('todoListContainer').textContent;

  assert.ok(!/Overdue/.test(html), 'a task due TODAY was flagged overdue');
  assert.ok(/Due Today/.test(html), 'a task due today was not flagged as due today');
});

test('a task due tomorrow is not shown as due today during the evening', () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  freezeClock(w);
  w.eval('_todos=[{id:2,dbId:2,task:"Call Bo",done:false,priority:"normal",due:"2026-09-04",clientId:""}];');
  w.renderTodos();
  const html = w.document.getElementById('todoListContainer').textContent;
  assert.ok(!/Due Today/.test(html), "tomorrow's task was flagged as due today");
});

test('the per-client task list uses the same local date', () => {
  const w = loadApp();
  resetStorage(w);
  freezeClock(w);
  const host = w.document.createElement('div');
  host.innerHTML = '<div id="viewForm"><div class="form-card"><div class="form-actions"></div></div></div>';
  w.document.body.appendChild(host);
  w.eval('_todos=[{id:3,dbId:3,task:"Call Cal",done:false,priority:"normal",due:"2026-09-03",clientId:"42"}];');
  w.renderClientTodos('42');
  const sec = w.document.getElementById('clientTodoSection');
  assert.ok(sec, 'expected the client task section');
  // Overdue is rendered as red text on the due date; today's task must not be red.
  assert.ok(!/#dc3545/.test(sec.innerHTML.split('Call Cal')[1] || ''),
    "a task due today was styled overdue in the client's task list");
});

// ── §164.528: a bulk extract of client records leaving the application is a disclosure.
test('every bulk export leaves a disclosure record', () => {
  const w = loadApp();
  resetStorage(w);
  const logged = [];
  const files = [];
  stub(w, {
    logActivity: (type, text) => logged.push({ type, text }),
    XLSX: { utils: { aoa_to_sheet: () => ({}), book_new: () => ({}), book_append_sheet: () => {} },
            writeFile: (wb, name) => files.push(name) },
  });
  w.eval('clients=[{_id:1,f_firstName:"Ada",f_lastName:"Lovelace",f_dob:"1990-01-01",f_phone:"5550100"}];');

  w.exportExcel();

  assert.strictEqual(logged.length, 1,
    'exportExcel sent the whole roster out with no disclosure record');
  assert.strictEqual(logged[0].type, 'export');
  assert.strictEqual(files.length, 1);
});

test('a backup filename is dated by the local day, not UTC', () => {
  const w = loadApp();
  resetStorage(w);
  freezeClock(w);
  const files = [];
  stub(w, {
    logActivity: () => {},
    XLSX: { utils: { aoa_to_sheet: () => ({}), book_new: () => ({}), book_append_sheet: () => {} },
            writeFile: (wb, name) => files.push(name) },
  });
  w.eval('clients=[{_id:1,f_firstName:"Ada",f_lastName:"Lovelace"}];');

  w.exportFullBackup();

  assert.strictEqual(files.length, 1);
  assert.match(files[0], /2026-09-03/,
    'an export taken on the evening of the 3rd was filed under the 4th');
});
