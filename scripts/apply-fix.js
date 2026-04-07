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

let resp;
try {
  resp = JSON.parse(fs.readFileSync(responseFile, 'utf8'));
} catch (e) {
  console.error(`ERROR: Could not parse ${responseFile}: ${e.message}`);
  process.exit(1);
}

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
  const tc = JSON.parse(fs.readFileSync(tcPath, 'utf8'));
  if (!tc.cases.some((c) => c.id === resp.newTestCase.id)) {
    tc.cases.push(resp.newTestCase);
    fs.writeFileSync(tcPath, JSON.stringify(tc, null, 2) + '\n');
    console.log('Added test case', resp.newTestCase.id);
  }
}

// Append new patch entry
if (resp.patchEntry) {
  const ppPath = path.join(ROOT, 'docs/patches/patches.json');
  const pm = JSON.parse(fs.readFileSync(ppPath, 'utf8'));
  if (!pm.patches.some((p) => p.id === resp.patchEntry.id)) {
    pm.patches.push(resp.patchEntry);
    fs.writeFileSync(ppPath, JSON.stringify(pm, null, 2) + '\n');
    console.log('Added patch entry', resp.patchEntry.id);
  }
}
