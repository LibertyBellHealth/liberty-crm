'use strict';
// A ratchet, not a bug report. Everything it flags today is fixed; the point is that the NEXT
// piece of code written in this file cannot quietly reintroduce the family.
//
// It checks two rules, because one is not enough:
//
//   1. No record-scoped global is read inside an async continuation or after an await. That is the
//      classic shape: the global says which record is open, the gap lets it change, the write lands
//      on whoever is on screen now.
//
//   2. No f_* form field is reached inside a continuation by a function that never asks whether it
//      is still on the record it started for. Rule 1 alone is blind to this: the ZIP -> county
//      lookup took no record id AT ALL and read no global, so there was nothing for a global-based
//      rule to find — and it was the worst instance in the file, because the f_* fields are static,
//      reused between patients, and read straight back by getFormData on save.
//
// Both lists are empty and must stay that way. An entry is a decision to be defended in review, not
// a way to make the test pass; a long list is how the sibling app's version of this became
// decoration. analyze() is exported to itself below and run against fixtures, because a static
// analyser with no self-test degrades silently into "always passes" and nobody finds out.
const { test } = require('node:test');
const assert = require('node:assert');
const fs = require('fs');
const path = require('path');
const acorn = require('acorn');
const walk = require('acorn-walk');

const RECORD_GLOBALS = new Set(['editingId', '_rowVersion', '_fullRecordFailed', '_fullRecordLoading', '_formDirty']);
const GAP_CALLEES = new Set(['setTimeout', 'setInterval', 'requestAnimationFrame', 'showConfirm', 'showPrompt', 'queueMicrotask']);
const GAP_METHODS = new Set(['then', 'catch', 'finally']);
const HANDLER_PROPS = new Set(['onload', 'onerror', 'onloadend', 'onreadystatechange']);
const GUARDS = new Set(['stillOnRecord', 'whenStillOnRecord']);

// Functions allowed to break each rule. Both empty on purpose — see the header.
const ALLOWED_GLOBAL_READ = {};
const ALLOWED_FORM_WRITE = {};

const isFn = (n) => n && (n.type === 'FunctionDeclaration' || n.type === 'FunctionExpression' || n.type === 'ArrowFunctionExpression');

function analyze(src) {
  const ast = acorn.parse(src, { ecmaVersion: 2022, sourceType: 'script', locations: true });

  // Which function is a node in? Report the outermost named declaration, the unit a reader edits.
  const owner = (anc) => {
    for (let i = 0; i < anc.length - 1; i++) {
      if (anc[i].type === 'FunctionDeclaration' && anc[i].id) return anc[i].id.name;
    }
    return '<toplevel>';
  };
  // A function passed to setTimeout/.then/showConfirm, or assigned to onload, is a continuation:
  // everything inside it runs after the gap.
  const isContinuation = (fn, parent) => {
    if (!parent) return false;
    if (parent.type === 'CallExpression' && parent.arguments.indexOf(fn) >= 0) {
      const c = parent.callee;
      if (c.type === 'Identifier' && GAP_CALLEES.has(c.name)) return true;
      if (c.type === 'MemberExpression' && c.property && GAP_METHODS.has(c.property.name)) return true;
    }
    if (parent.type === 'AssignmentExpression' && parent.right === fn &&
        parent.left.type === 'MemberExpression' && HANDLER_PROPS.has(parent.left.property && parent.left.property.name)) return true;
    return false;
  };

  const awaitsByFn = new Map();
  walk.ancestor(ast, { AwaitExpression(node, _s, anc) {
    for (let i = anc.length - 2; i >= 0; i--) {
      if (isFn(anc[i])) { if (!awaitsByFn.has(anc[i])) awaitsByFn.set(anc[i], []); awaitsByFn.get(anc[i]).push(node.start); break; }
    }
  } });
  const afterGap = (anc, start) => {
    let inner = null;
    for (let i = anc.length - 2; i >= 0; i--) {
      const a = anc[i];
      if (!isFn(a)) continue;
      if (!inner) inner = a;
      if (isContinuation(a, anc[i - 1])) return true;
    }
    return !!(inner && inner.async && awaitsByFn.has(inner) && awaitsByFn.get(inner).some((s) => s < start));
  };

  // Which functions ask the question at all.
  const guarded = new Set();
  walk.ancestor(ast, { CallExpression(node, _s, anc) {
    if (node.callee.type === 'Identifier' && GUARDS.has(node.callee.name)) guarded.add(owner(anc));
  } });

  const globalReads = [], formWrites = [];
  walk.ancestor(ast, {
    Identifier(node, _s, anc) {
      if (!RECORD_GLOBALS.has(node.name)) return;
      const p = anc[anc.length - 2];
      if (!p) return;
      if (p.type === 'MemberExpression' && p.property === node && !p.computed) return;   // obj.editingId
      if (p.type === 'Property' && p.key === node && !p.computed) return;                // {editingId: x}
      if (p.type === 'AssignmentExpression' && p.left === node) return;                  // a write, not a read
      const fn = owner(anc);
      if (GUARDS.has(fn)) return;                                                        // the guard itself
      if (afterGap(anc, node.start)) globalReads.push({ fn, name: node.name, line: node.loc.start.line });
    },
    CallExpression(node, _s, anc) {
      const c = node.callee;
      if (!(c.type === 'MemberExpression' && c.property && c.property.name === 'getElementById')) return;
      const arg = node.arguments[0];
      if (!arg) return;
      // 'f_notes', or 'f_'+prefix+'City'
      const mentionsFormField = (n) =>
        (n.type === 'Literal' && typeof n.value === 'string' && n.value.indexOf('f_') === 0) ||
        (n.type === 'BinaryExpression' && (mentionsFormField(n.left) || mentionsFormField(n.right)));
      if (!mentionsFormField(arg)) return;
      if (!afterGap(anc, node.start)) return;
      const fn = owner(anc);
      if (!guarded.has(fn)) formWrites.push({ fn, line: node.loc.start.line });
    },
  });
  return { globalReads, formWrites };
}

const appSrc = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const fmt = (rows) => rows.map((r) => '  ' + r.fn + ' (app.js:' + r.line + (r.name ? ', ' + r.name : '') + ')').join('\n');

test('no record global is read after an async gap', () => {
  const bad = analyze(appSrc).globalReads.filter((r) => !Object.prototype.hasOwnProperty.call(ALLOWED_GLOBAL_READ, r.fn));
  assert.strictEqual(bad.length, 0,
    'These read a record global AFTER a gap, so it may name a different record by then.\n' +
    'Capture the id before the gap and pass it in; guard the write with stillOnRecord.\n' + fmt(bad));
});

test('no f_* field is reached after an async gap without asking whether the record changed', () => {
  const bad = analyze(appSrc).formWrites.filter((r) => !Object.prototype.hasOwnProperty.call(ALLOWED_FORM_WRITE, r.fn));
  assert.strictEqual(bad.length, 0,
    'These reach a static, save-backed form field after a gap and never call stillOnRecord.\n' +
    'The f_* fields are reused between patients and read back by getFormData on save.\n' + fmt(bad));
});

// ---- the analyser's own tests ----
// Each rule gets a case it must flag and a case it must NOT, so the detector cannot rot into
// "always passes" without one of these going red.

test('rule 1 flags a global read inside a continuation', () => {
  const r = analyze('function f(){ fetch(u).then(function(){ save(editingId); }); }');
  assert.strictEqual(r.globalReads.length, 1, JSON.stringify(r.globalReads));
  assert.strictEqual(r.globalReads[0].fn, 'f');
});

test('rule 1 flags a global read after an await', () => {
  const r = analyze('async function f(){ await go(); save(editingId); }');
  assert.strictEqual(r.globalReads.length, 1, JSON.stringify(r.globalReads));
});

test('rule 1 ignores a synchronous read, and a captured id passed into the gap', () => {
  assert.strictEqual(analyze('function f(){ var id=editingId; save(id); }').globalReads.length, 0,
    'a read with no gap in front of it is fine');
  assert.strictEqual(analyze('function f(){ var id=editingId; fetch(u).then(function(){ save(id); }); }').globalReads.length, 0,
    'capturing before the gap is exactly the fix — it must not be flagged');
});

test('rule 1 ignores assignments and property names that merely share the word', () => {
  assert.strictEqual(analyze('function f(){ fetch(u).then(function(){ editingId=null; }); }').globalReads.length, 0,
    'clearing the global is a write, not a stale read');
  assert.strictEqual(analyze('function f(){ fetch(u).then(function(){ log(o.editingId); }); }').globalReads.length, 0,
    'a property of some other object is not the global');
});

test('rule 2 flags an unguarded form write behind a gap, in both spellings', () => {
  const lit = analyze("function f(){ fetch(u).then(function(){ document.getElementById('f_resCity').value=x; }); }");
  assert.strictEqual(lit.formWrites.length, 1, JSON.stringify(lit.formWrites));
  const built = analyze("function f(p){ fetch(u).then(function(){ document.getElementById('f_'+p+'City').value=x; }); }");
  assert.strictEqual(built.formWrites.length, 1, 'a built-up id must be caught too: ' + JSON.stringify(built.formWrites));
});

test('rule 2 ignores a guarded write, a synchronous one, and a non-form element', () => {
  assert.strictEqual(
    analyze("function f(id){ fetch(u).then(function(){ if(!stillOnRecord(id))return; document.getElementById('f_resCity').value=x; }); }").formWrites.length, 0,
    'asking the question is the whole point — a guarded write must not be flagged');
  assert.strictEqual(analyze("function f(){ document.getElementById('f_resCity').value=x; }").formWrites.length, 0,
    'no gap, no problem');
  assert.strictEqual(analyze("function f(){ fetch(u).then(function(){ document.getElementById('toastContainer').innerHTML=x; }); }").formWrites.length, 0,
    'only the per-record f_* fields are save-backed');
});

test('the analyser is being run against the real app.js, not an empty string', () => {
  assert.ok(appSrc.length > 50000, 'app.js looks too small to be the real file');
  assert.ok(/function stillOnRecord/.test(appSrc), 'the guard this test assumes is missing from app.js');
});
