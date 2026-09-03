'use strict';
// Settings live on the server (AppSettings, scope 'health') with localStorage as the local copy.
// Since the sign-out/tab-close wipe clears every crm_* key, a sign-in starts from nothing and
// depends on loadSettingsAPI to refill it. What happens when that fetch FAILS is the whole story:
// the in-memory lists fall back to hardcoded defaults, and carriers to an empty array.
const { test } = require('node:test');
const assert = require('node:assert');
const { loadApp, resetStorage, stub } = require('./harness');

const tick = () => new Promise(r => setTimeout(r, 0));
const toasts = (w) => Array.from(w.document.querySelectorAll('#toastContainer .toast')).map(t => t.textContent);

test('a failed settings load must not let the next edit overwrite the server', async () => {
  const w = loadApp();
  resetStorage(w);
  w.eval('_apiToken="fake"; _settingsPushInFlight=0;');
  const posts = [];
  stub(w, { fetch: (url, opts) => {
    if (/\/settings/.test(url) && opts && opts.method === 'POST') { posts.push(JSON.parse(opts.body)); return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) }); }
    return Promise.resolve({ ok: false, status: 500, json: () => Promise.resolve({ error: 'down' }) }); // the GET fails
  }});

  await w.loadSettingsAPI();
  // The agent list in memory is now just the hardcoded defaults, because nothing was restored.
  // Adding one and pushing would replace the real server list with defaults+1.
  w.localStorage.setItem('crm_agents', JSON.stringify(['Thomas Jaboro', 'Paul Jaboro Jr.', 'New Guy']));
  w._pushSettings();
  await tick(); await tick();

  assert.strictEqual(posts.length, 0,
    'settings were pushed to the server without ever having loaded the server copy');
  assert.match(toasts(w).join(' '), /settings/i, 'the user was not told settings are not syncing');
});

test('a successful load — even an empty one — allows pushes again', async () => {
  const w = loadApp();
  resetStorage(w);
  w.eval('_apiToken="fake"; _settingsPushInFlight=0;');
  const posts = [];
  stub(w, { fetch: (url, opts) => {
    if (/\/settings/.test(url) && opts && opts.method === 'POST') { posts.push(JSON.parse(opts.body)); return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) }); }
    // A brand-new install: the server genuinely has nothing yet. That IS a successful load, and
    // local values SHOULD be pushed up from it.
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) });
  }});

  await w.loadSettingsAPI();
  w.localStorage.setItem('crm_agents', JSON.stringify(['Thomas Jaboro']));
  w._pushSettings();
  await tick(); await tick();

  assert.strictEqual(posts.length, 1, 'a first-run push was blocked when it should have gone through');
  assert.strictEqual(posts[0].scope, 'health');
});

test('a load that returns real settings restores them and allows pushes', async () => {
  const w = loadApp();
  resetStorage(w);
  w.eval('_apiToken="fake"; _settingsPushInFlight=0;');
  const posts = [];
  stub(w, { fetch: (url, opts) => {
    if (/\/settings/.test(url) && opts && opts.method === 'POST') { posts.push(JSON.parse(opts.body)); return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({}) }); }
    return Promise.resolve({ ok: true, status: 200, json: () => Promise.resolve({
      agents: ['Ada', 'Bo', 'Cal'], carriers: [{ name: 'Ambetter' }],
    }) });
  }});

  await w.loadSettingsAPI();

  assert.deepStrictEqual(JSON.parse(w.localStorage.getItem('crm_agents')), ['Ada', 'Bo', 'Cal']);
  w._pushSettings();
  await tick(); await tick();
  assert.strictEqual(posts.length, 1);
});
