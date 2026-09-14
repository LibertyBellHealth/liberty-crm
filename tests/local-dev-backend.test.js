'use strict';
// The Health CRM had its backend and sign-in redirect hard-wired to production, so the only place
// it could run was the live site — there was nowhere safe to test it, ahead of a 1,000-record
// import. Served from localhost it now talks to the DEV backend (fake data) and redirects sign-in
// back to localhost, mirroring the switch Home Care has had for months.
//
// The property that matters most is the negative one: every deployed host must still talk to
// production, and localhost must never talk to production. A switch that got either wrong would be
// the worst bug this file could have, so each direction is asserted against the real app.js.
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const { JSDOM } = require('jsdom');

const PROD_API = 'https://liberty-crm-api-cyb3dkhnd2e7a3cy.centralus-01.azurewebsites.net/api';
const DEV_API = 'https://liberty-crm-api-dev.azurewebsites.net/api';
const PROD_SITE = 'https://polite-pebble-039f4a010.7.azurestaticapps.net';

// Load the real app.js at a given address and read what it chose. The choice is made in the first
// lines of the file, before anything that needs a full page, so a later startup error in this bare
// DOM cannot affect it.
function configAt(url) {
  const src = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
  const dom = new JSDOM('<!doctype html><html><body></body></html>', { url, runScripts: 'outside-only' });
  const w = dom.window;
  w.fetch = () => new Promise(() => {});
  try { w.eval(src); } catch (e) { /* startup may throw without the real page; config is already set */ }
  const out = { api: w.API_BASE, redirect: w.REDIRECT_URI };
  w.close();
  return out;
}

test('served from localhost, it talks to the dev backend and signs in back to localhost', () => {
  const c = configAt('http://localhost:4280/');
  assert.strictEqual(c.api, DEV_API, 'localhost must use the dev backend, got ' + c.api);
  assert.strictEqual(c.redirect, 'http://localhost:4280', 'sign-in must return to localhost, got ' + c.redirect);
});

test('the production site still talks to production, unchanged', () => {
  const c = configAt(PROD_SITE + '/');
  assert.strictEqual(c.api, PROD_API, 'production must never use the dev backend');
  assert.strictEqual(c.redirect, PROD_SITE);
});

test('any other deployed host (a PR preview) also stays on production', () => {
  const c = configAt('https://polite-pebble-039f4a010-25.centralus.7.azurestaticapps.net/');
  assert.strictEqual(c.api, PROD_API, 'only localhost may switch to dev — never a deployed host');
});
