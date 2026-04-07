#!/usr/bin/env node
// scripts/run-tests.js — jsdom-based test harness for semai patches
//
// Usage:
//   node scripts/run-tests.js                  # run all cases
//   node scripts/run-tests.js --patch=<id>     # only cases targeting this patch
//   node scripts/run-tests.js --case=<id>      # run a single case by ID
//   node scripts/run-tests.js --format=json    # machine-readable output (for LLM)

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import { JSDOM } from 'jsdom';

const __dirname = path.dirname(fileURLToPath(import.meta.url));

// ── CLI args ──────────────────────────────────────────────────────────────────

const args = process.argv.slice(2);
const arg = (name) => {
  const flag = args.find((a) => a.startsWith(`--${name}=`));
  return flag ? flag.split('=').slice(1).join('=') : null;
};

const filterPatch = arg('patch');
const filterCase = arg('case');
const formatJson = arg('format') === 'json';

// ── Load manifests ────────────────────────────────────────────────────────────

const ROOT = path.resolve(__dirname, '..');
const casesPath = path.join(ROOT, 'tests', 'test-cases.json');
const patchesPath = path.join(ROOT, 'docs', 'patches', 'patches.json');

const { fixtures, cases } = JSON.parse(fs.readFileSync(casesPath, 'utf8'));
const { patches } = JSON.parse(fs.readFileSync(patchesPath, 'utf8'));

const patchMap = Object.fromEntries(patches.map((p) => [p.id, p]));

// ── Assertion runners ─────────────────────────────────────────────────────────

function runAssertion(dom, assertion) {
  const doc = dom.window.document;
  const { type, selector } = assertion;

  switch (type) {
    case 'selectorExists': {
      const el = doc.querySelector(selector);
      return el
        ? { pass: true }
        : { pass: false, reason: `selector "${selector}" matched nothing` };
    }

    case 'selectorNotExists': {
      const el = doc.querySelector(selector);
      return el
        ? { pass: false, reason: `selector "${selector}" unexpectedly matched` }
        : { pass: true };
    }

    case 'selectorCount': {
      const els = doc.querySelectorAll(selector);
      const got = els.length;
      const want = assertion.count;
      return got === want
        ? { pass: true }
        : { pass: false, reason: `selector "${selector}" matched ${got}, expected ${want}` };
    }

    case 'attributeEquals': {
      const el = doc.querySelector(selector);
      if (!el) return { pass: false, reason: `selector "${selector}" matched nothing` };
      const val = el.getAttribute(assertion.attribute);
      return val === assertion.value
        ? { pass: true }
        : { pass: false, reason: `attribute "${assertion.attribute}" is "${val}", expected "${assertion.value}"` };
    }

    case 'textContains': {
      const el = doc.querySelector(selector);
      if (!el) return { pass: false, reason: `selector "${selector}" matched nothing` };
      const text = el.textContent || '';
      return text.includes(assertion.text)
        ? { pass: true }
        : { pass: false, reason: `text "${text.slice(0, 80)}" does not contain "${assertion.text}"` };
    }

    default:
      return { pass: false, reason: `unknown assertion type "${type}"` };
  }
}

// ── Build DOM for a test case ─────────────────────────────────────────────────

function buildDom(tc) {
  const fixtureDef = fixtures[tc.fixture];
  if (!fixtureDef) throw new Error(`Unknown fixture: ${tc.fixture}`);

  const fixturePath = path.join(ROOT, fixtureDef.path);
  const html = fs.readFileSync(fixturePath, 'utf8');
  const dom = new JSDOM(html, { runScripts: 'outside-only' });

  // If this case targets a patch, simulate injection
  if (tc.patch) {
    const patch = patchMap[tc.patch];
    if (!patch) throw new Error(`Unknown patch: ${tc.patch}`);

    const doc = dom.window.document;
    if (patch.type === 'js') {
      const script = doc.createElement('script');
      script.textContent = patch.code;
      script.setAttribute('data-semai-patch', patch.id);
      (doc.head || doc.documentElement).appendChild(script);
    } else if (patch.type === 'css') {
      const style = doc.createElement('style');
      style.textContent = patch.code;
      style.setAttribute('data-semai-patch', patch.id);
      (doc.head || doc.documentElement).appendChild(style);
    }
  }

  return dom;
}

// ── Run cases ─────────────────────────────────────────────────────────────────

let selected = cases;
if (filterCase) selected = selected.filter((c) => c.id === filterCase);
if (filterPatch) selected = selected.filter((c) => c.patch === filterPatch);

if (selected.length === 0) {
  console.error('No matching test cases found.');
  process.exit(1);
}

const results = [];
let passed = 0;
let failed = 0;

for (const tc of selected) {
  const caseResult = { id: tc.id, description: tc.description, assertions: [] };
  let dom;

  try {
    dom = buildDom(tc);
  } catch (err) {
    caseResult.error = err.message;
    caseResult.pass = false;
    results.push(caseResult);
    failed++;
    continue;
  }

  let casePassed = true;
  for (const assertion of tc.assertions) {
    const { pass, reason } = runAssertion(dom, assertion);
    const assertResult = { type: assertion.type, pass, message: assertion.message };
    if (!pass) {
      assertResult.reason = reason;
      casePassed = false;
    }
    caseResult.assertions.push(assertResult);
  }

  caseResult.pass = casePassed;
  results.push(caseResult);
  if (casePassed) passed++; else failed++;
}

// ── Output ────────────────────────────────────────────────────────────────────

if (formatJson) {
  console.log(JSON.stringify({ passed, failed, total: results.length, cases: results }, null, 2));
} else {
  for (const r of results) {
    const icon = r.pass ? '\u2713' : '\u2717';
    console.log(`\n${icon} ${r.id}: ${r.description}`);
    if (r.error) {
      console.log(`  ERROR: ${r.error}`);
    } else {
      for (const a of r.assertions) {
        const aIcon = a.pass ? '  \u2713' : '  \u2717';
        console.log(`${aIcon} [${a.type}] ${a.message}`);
        if (!a.pass) console.log(`       -> ${a.reason}`);
      }
    }
  }
  console.log(`\n${passed}/${results.length} passed.`);
}

process.exit(failed > 0 ? 1 : 0);
