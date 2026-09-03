# Smoke tests — Health CRM

Fast, dependency-light checks on `../app.js`. They load the **real** `app.js` in jsdom and
call its actual functions — nothing is copied or re-implemented here, so a test can only
pass if the shipped code behaves.

## Run

```bash
cd tests
npm install   # first time only (installs jsdom, dev-only, gitignored)
npm test
```

## Why it lives here (not the repo root)

Azure Static Web Apps deploys this site with `app_location: "/"` and no build command, so a
`package.json` at the **repo root** would make Azure treat the static site as a Node app and
try to build it. Keeping the setup self-contained in this folder leaves the root untouched.
`node_modules/` is gitignored and the folder is 404'd on the live site via
`staticwebapp.config.json`.

## Adding tests

New `*.test.js` files here are picked up automatically by `node --test`. Get the app's
functions with `const w = require('./harness').loadApp();`.

Two things to know:

- **Values returned from app.js live in the jsdom realm.** Compare primitives directly, and
  use `JSON.stringify` (or the `jsonEqual` helper) for objects and arrays rather than
  `deepStrictEqual`, which fails on prototype identity across realms.
- **Replace app.js globals with `stub(w, {...})`, never by assigning directly.** `loadApp()`
  caches one window for the whole run, so a directly-assigned stub stays in place for every
  later test — that is how the first version of this suite had one test calling another
  test's fake `deleteClientDoc`. `resetStorage(w)` restores everything registered via `stub`.

## Mutation-test what you add

A test that has never been observed to fail is not coverage. After writing one, break the
fix it covers on purpose, confirm the test goes red, then restore. Writing this suite turned
up one test that passed with its fix fully reverted.
