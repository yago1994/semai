#!/usr/bin/env node
// run-tests.js — Test harness for self-healing extension patches.
//
// Loads patches.json + test-cases.json, runs each test case in jsdom,
// and emits structured JSON results to stdout. Designed to be consumed
// by both humans and the LLM-driven auto-fix engine.
//
// Usage:
//   node scripts/run-tests.js                       # all tests
//   node scripts/run-tests.js --patch=patch-id      # only tests for one patch
//   node scripts/run-tests.js --case=tc-001         # only one case
//   node scripts/run-tests.js --format=pretty       # human-readable output

const fs = require("fs");
const path = require("path");
const { JSDOM } = require("jsdom");

const ROOT = path.resolve(__dirname, "..");
const PATCHES_PATH = path.join(ROOT, "patches", "patches.json");
const CASES_PATH = path.join(ROOT, "tests", "test-cases.json");
const FIXTURES_DIR = path.join(ROOT, "tests", "fixtures");

// ---- CLI args --------------------------------------------------------------

const args = Object.fromEntries(
  process.argv.slice(2).map((a) => {
    const [k, v] = a.replace(/^--/, "").split("=");
    return [k, v ?? true];
  })
);

// ---- Load inputs -----------------------------------------------------------

const manifest = JSON.parse(fs.readFileSync(PATCHES_PATH, "utf8"));
const testFile = JSON.parse(fs.readFileSync(CASES_PATH, "utf8"));
const patchesById = Object.fromEntries(manifest.patches.map((p) => [p.id, p]));

let cases = testFile.cases;
if (args.patch) cases = cases.filter((c) => c.patchId === args.patch);
if (args.case) cases = cases.filter((c) => c.id === args.case);

// ---- Assertion runners -----------------------------------------------------

const ASSERTIONS = {
  selectorExists(doc, a) {
    const el = doc.querySelector(a.selector);
    return el
      ? { pass: true }
      : { pass: false, reason: `selector "${a.selector}" matched nothing` };
  },
  selectorNotExists(doc, a) {
    const el = doc.querySelector(a.selector);
    return el
      ? { pass: false, reason: `selector "${a.selector}" should not have matched` }
      : { pass: true };
  },
  selectorCount(doc, a) {
    const n = doc.querySelectorAll(a.selector).length;
    return n === a.count
      ? { pass: true }
      : { pass: false, reason: `expected ${a.count} matches for "${a.selector}", got ${n}` };
  },
  attributeEquals(doc, a) {
    const el = doc.querySelector(a.selector);
    if (!el) return { pass: false, reason: `selector "${a.selector}" not found` };
    const actual = el.getAttribute(a.attribute);
    return actual === a.value
      ? { pass: true }
      : { pass: false, reason: `attr "${a.attribute}" was "${actual}", expected "${a.value}"` };
  },
  textContains(doc, a) {
    const el = doc.querySelector(a.selector);
    if (!el) return { pass: false, reason: `selector "${a.selector}" not found` };
    return el.textContent.includes(a.text)
      ? { pass: true }
      : { pass: false, reason: `text did not contain "${a.text}"` };
  },
};

// ---- Run a single test case ------------------------------------------------

function runCase(testCase) {
  const patch = patchesById[testCase.patchId];
  if (!patch) {
    return {
      id: testCase.id,
      status: "error",
      error: `no patch found with id "${testCase.patchId}"`,
    };
  }

  const fixturePath = path.join(FIXTURES_DIR, testCase.fixture);
  if (!fs.existsSync(fixturePath)) {
    return { id: testCase.id, status: "error", error: `fixture not found: ${testCase.fixture}` };
  }
  const html = fs.readFileSync(fixturePath, "utf8");

  const dom = new JSDOM(html, { runScripts: "outside-only" });
  const { window } = dom;

  const errors = [];
  window.addEventListener("error", (e) => errors.push(e.message));

  // Apply the patch (CSS or JS) the requested number of times
  const times = testCase.runPatchTimes || 1;
  try {
    for (let i = 0; i < times; i++) {
      if (patch.type === "js") {
        window.eval(patch.code);
      } else if (patch.type === "css") {
        const style = window.document.createElement("style");
        style.textContent = patch.code;
        window.document.head.appendChild(style);
      }
    }
  } catch (err) {
    return {
      id: testCase.id,
      status: "fail",
      patchId: testCase.patchId,
      error: `patch threw: ${err.message}`,
      assertions: [],
    };
  }

  // Run assertions
  const results = testCase.assertions.map((a) => {
    const fn = ASSERTIONS[a.type];
    if (!fn) return { type: a.type, pass: false, reason: `unknown assertion type` };
    const r = fn(window.document, a);
    return { type: a.type, ...a, ...r };
  });

  const allPass = results.every((r) => r.pass) && errors.length === 0;
  return {
    id: testCase.id,
    status: allPass ? "pass" : "fail",
    patchId: testCase.patchId,
    description: testCase.description,
    assertions: results,
    runtimeErrors: errors,
  };
}

// ---- Run all + emit --------------------------------------------------------

const results = cases.map(runCase);
const summary = {
  total: results.length,
  passed: results.filter((r) => r.status === "pass").length,
  failed: results.filter((r) => r.status === "fail").length,
  errored: results.filter((r) => r.status === "error").length,
};

if (args.format === "pretty") {
  for (const r of results) {
    const icon = r.status === "pass" ? "✓" : r.status === "fail" ? "✗" : "!";
    console.log(`${icon} ${r.id} — ${r.description || ""}`);
    if (r.status !== "pass") {
      if (r.error) console.log(`    error: ${r.error}`);
      for (const a of r.assertions || []) {
        if (!a.pass) console.log(`    fail: ${a.type} — ${a.reason}`);
      }
      for (const e of r.runtimeErrors || []) console.log(`    runtime: ${e}`);
    }
  }
  console.log(`\n${summary.passed}/${summary.total} passed`);
} else {
  console.log(JSON.stringify({ summary, results }, null, 2));
}

process.exit(summary.failed + summary.errored > 0 ? 1 : 0);
