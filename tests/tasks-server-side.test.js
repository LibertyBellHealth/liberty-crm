'use strict';
// The task list moved from localStorage to the server-side Tasks table (source='health').
// These pin the two reasons it moved — PHI at rest, and the sign-out wipe destroying the list —
// and the failure paths, since nothing is cached locally any more.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, stub, confirmOk } = require('./harness');

const tick = () => new Promise(r => setTimeout(r, 0));
const toasts = (w) => Array.from(w.document.querySelectorAll('#toastContainer .toast')).map(t => t.textContent);

// Minimal DOM for the task pane. renderTodos bails without its container.
function todoDom(w) {
  let host = w.document.getElementById('todoListContainer');
  if (!host) {
    const d = w.document.createElement('div');
    d.innerHTML = '<div id="todoListContainer"></div><div id="todoEmpty"></div>' +
      '<div id="todoAddSection"></div><input id="todoTaskInput"><input id="todoDueInput">' +
      '<select id="todoPriorityInput"><option value="normal" selected>normal</option></select>' +
      '<input id="todoClientInput"><input id="todoClientId">';
    w.document.body.appendChild(d);
  }
}

function okJson(body) {
  return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve(body) });
}

test('the task list loads from the server and never touches localStorage', async () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  w.eval('_apiToken="fake";');
  const calls = [];
  stub(w, {
    fetch: (url, opts) => { calls.push({ url, opts }); return okJson([
      { id: 7, task_text: 'Call about renewal', done: 0, due_date: '2026-09-10T00:00:00.000Z',
        client_name: '42', priority: 'high' },
    ]); },
  });

  await w.loadTasksAPI();

  assert.match(calls[0].url, /\/tasks\?source=health$/, 'the load was not scoped to health');
  assert.strictEqual(w._todos.length, 1);
  assert.strictEqual(w._todos[0].task, 'Call about renewal');
  assert.strictEqual(w._todos[0].due, '2026-09-10', 'the date was not trimmed to a plain day');
  assert.strictEqual(w._todos[0].clientId, '42', 'client_name carries the id for health rows');
  assert.strictEqual(w._todos[0].dbId, 7);
  assert.strictEqual(w.localStorage.getItem('crm_todos'), null, 'task text was written to localStorage');
});

test('a failed task load says so instead of showing an empty list', async () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  w.eval('_apiToken="fake";');
  stub(w, { fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'db down' }) }) });

  await w.loadTasksAPI();

  assert.match(toasts(w).join(' '), /could not load tasks/i);
});

test('adding a task writes it to the server, and a refusal is not hidden', async () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  w.eval('_apiToken="fake"; _todos=[];');
  w.document.getElementById('todoTaskInput').value = 'Call about renewal';
  const posted = [];
  stub(w, { fetch: (url, opts) => { posted.push(JSON.parse(opts.body)); return okJson({ id: 99 }); } });

  w.saveTodo();
  await tick(); await tick();

  assert.strictEqual(posted.length, 1);
  assert.strictEqual(posted[0].source, 'health', 'the task was not tagged as a health row');
  assert.strictEqual(posted[0].text, 'Call about renewal');
  assert.strictEqual(w._todos[0].dbId, 99, 'the server id was not adopted');
  assert.strictEqual(w.localStorage.getItem('crm_todos'), null);
});

test('a task the server refused is marked NOT SAVED on the row', async () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  w.eval('_apiToken="fake"; _todos=[];');
  w.document.getElementById('todoTaskInput').value = 'Call about renewal';
  stub(w, { fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'nope' }) }) });

  w.saveTodo();
  await tick(); await tick();

  assert.strictEqual(w._todos[0]._unsaved, true);
  assert.match(w.document.getElementById('todoListContainer').textContent, /NOT SAVED/,
    'the row gave no sign the task was never stored');
  assert.match(toasts(w).join(' '), /not saved/i);
});

test('a tick that fails to save is put back', async () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  w.eval('_apiToken="fake"; _todos=[{id:1,dbId:1,task:"t",done:false,priority:"normal",due:"",clientId:""}];');
  stub(w, { fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'nope' }) }) });

  w.toggleTodo(1);
  await tick(); await tick();

  assert.strictEqual(w._todos[0].done, false, 'the task still reads as done after the save failed');
  assert.match(toasts(w).join(' '), /could not update/i);
});

test('a task is only removed locally once the server has accepted the delete', async () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  w.eval('_apiToken="fake"; _todos=[{id:1,dbId:5,task:"t",done:false,priority:"normal",due:"",clientId:""}];');
  const calls = [];
  stub(w, {
    fetch: (url, opts) => { calls.push(url + ' ' + (opts && opts.method)); 
      return Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'nope' }) }); },
  });

  w.deleteTodo(1);
  confirmOk(w);
  await tick(); await tick();

  assert.match(calls[0], /\/tasks\/5 DELETE/, 'the delete did not go to the server');
  assert.strictEqual(w._todos.length, 1, 'a refused delete still removed the task from the list');
  assert.match(toasts(w).join(' '), /still there/i);
});

// ── The migration is the part that can lose data if it is wrong, so it gets the most attention.
test('tasks already on the device are moved to the server before the key is dropped', async () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  w.eval('_apiToken="fake";');
  w.localStorage.setItem('crm_todos', JSON.stringify([
    { id: 1, task: 'Call about renewal', done: false, due: '2026-09-10', priority: 'high', clientId: '42' },
    { id: 2, task: 'Send ID card', done: true, due: '', priority: 'normal', clientId: '' },
  ]));
  const posted = [];
  stub(w, { fetch: (url, opts) => { posted.push(JSON.parse(opts.body)); return okJson({ id: posted.length }); } });

  await w.migrateLegacyTodos();

  assert.strictEqual(posted.length, 2, 'not every local task was sent');
  assert.deepStrictEqual(posted.map(p => p.text).sort(), ['Call about renewal', 'Send ID card']);
  assert.ok(posted.every(p => p.source === 'health'));
  assert.strictEqual(w.localStorage.getItem('crm_todos'), null, 'the local copy was not cleaned up');
});

test('a migration that fails keeps the local copy rather than losing the tasks', async () => {
  const w = loadApp();
  resetStorage(w);
  todoDom(w);
  w.eval('_apiToken="fake";');
  w.localStorage.setItem('crm_todos', JSON.stringify([{ id: 1, task: 'Call about renewal', priority: 'normal' }]));
  stub(w, { fetch: () => Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'nope' }) }) });

  await w.migrateLegacyTodos();

  assert.ok(w.localStorage.getItem('crm_todos'), 'the only copy of the tasks was deleted after a failed migration');
  assert.match(toasts(w).join(' '), /still only on this device/i);
});

test('the sign-out wipe spares the un-migrated task list but takes everything else', () => {
  const w = loadApp();
  resetStorage(w);
  w.localStorage.setItem('crm_todos', JSON.stringify([{ id: 1, task: 'Call about renewal' }]));
  w.localStorage.setItem('crm_recent', JSON.stringify([{ id: 5 }]));
  w.localStorage.setItem('lch_client_page_size', '25');

  w.clearCRMStorage();

  assert.ok(w.localStorage.getItem('crm_todos'), 'the only copy of the tasks was wiped on sign-out');
  assert.strictEqual(w.localStorage.getItem('crm_recent'), null, 'recent records survived the wipe');
  assert.strictEqual(w.localStorage.getItem('lch_client_page_size'), '25', 'a plain preference was wiped');
});
