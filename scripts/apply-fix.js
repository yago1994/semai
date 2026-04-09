#!/usr/bin/env node
// scripts/apply-fix.js
// Reads /tmp/claude_response.json and applies Claude's proposed changes:
//   - writes updated source files
//   - appends new test case to tests/test-cases.json
//   - appends new patch entry to docs/patches/patches.json

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

const responseFile = process.argv[2] ?? '/tmp/claude_response.json';

// Fix unrecognized JSON escape sequences (e.g. \. in regex patterns).
// Valid JSON escapes: \" \\ \/ \b \f \n \r \t \uXXXX — anything else is illegal.
function fixJsonEscapes(str) {
  return str.replace(/\\(?!["\\/bfnrtu\d])/g, '\\\\');
}

function parseJson(str, label) {
  try {
    return JSON.parse(str);
  } catch {
    try {
      return JSON.parse(fixJsonEscapes(str));
    } catch (e2) {
      console.error(`ERROR: Could not parse ${label}: ${e2.message}`);
      process.exit(1);
    }
  }
}

let resp;
resp = parseJson(fs.readFileSync(responseFile, 'utf8'), responseFile);

// Apply file changes
for (const f of (resp.files ?? [])) {
  const abs = path.join(ROOT, f.path);
  fs.mkdirSync(path.dirname(abs), { recursive: true });
  fs.writeFileSync(abs, f.content, 'utf8');
  console.log('Wrote', f.path);
}

// Append new test case
if (resp.newTestCase) {
  const tcPath = path.join(ROOT, 'tests/test-cases.json');
  const tc = parseJson(fs.readFileSync(tcPath, 'utf8'), tcPath);
  if (!tc.cases.some((c) => c.id === resp.newTestCase.id)) {
    tc.cases.push(resp.newTestCase);
    fs.writeFileSync(tcPath, JSON.stringify(tc, null, 2) + '\n');
    console.log('Added test case', resp.newTestCase.id);
  }
}

// Append new patch entry
if (resp.patchEntry) {
  const ppPath = path.join(ROOT, 'docs/patches/patches.json');
  const pm = parseJson(fs.readFileSync(ppPath, 'utf8'), ppPath);
  if (!pm.patches.some((p) => p.id === resp.patchEntry.id)) {
    pm.patches.push(resp.patchEntry);
    fs.writeFileSync(ppPath, JSON.stringify(pm, null, 2) + '\n');
    console.log('Added patch entry', resp.patchEntry.id);
  }
}
