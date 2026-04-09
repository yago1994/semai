#!/usr/bin/env node
// scripts/extract-approved-patch.js
// Extracts a pre-approved patch from a GitHub issue body (embedded by the extension's
// "Approve & Report" flow) and adds it to docs/patches/patches.json.
//
// Required env vars:
//   ISSUE_BODY    — full markdown body of the GitHub issue
//   ISSUE_NUMBER  — GitHub issue number (used to generate patch ID)

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

const issueBody = process.env.ISSUE_BODY ?? '';
const issueNumber = process.env.ISSUE_NUMBER ?? '0';

// ── Extract the approved patch JSON from HTML comment markers ────────────────

const START_MARKER = '<!-- SEMAI_APPROVED_PATCH';
const END_MARKER = 'SEMAI_APPROVED_PATCH -->';

const startIdx = issueBody.indexOf(START_MARKER);
const endIdx = issueBody.indexOf(END_MARKER);

if (startIdx === -1 || endIdx === -1 || endIdx <= startIdx) {
  console.log('No approved patch found in issue body — falling back to Claude-based auto-fix.');
  process.exit(2); // exit code 2 = no approved patch, workflow should fall back
}

const jsonStr = issueBody.slice(startIdx + START_MARKER.length, endIdx).trim();

let patch;
try {
  patch = JSON.parse(jsonStr);
} catch (e) {
  console.error(`ERROR: Could not parse approved patch JSON: ${e.message}`);
  console.error('Raw content:', jsonStr.slice(0, 500));
  process.exit(1);
}

if (!patch.patchType || !patch.patchCode) {
  console.error('ERROR: Approved patch is missing required fields (patchType, patchCode).');
  process.exit(1);
}

// ── Generate patch entry and add to patches.json ─────────────────────────────

const patchId = `fix-issue-${issueNumber}-v1`;
const patchEntry = {
  id: patchId,
  description: patch.explanation || `Auto-fix for issue #${issueNumber}`,
  urlPattern: patch.urlPattern || '^https://outlook\\.(office(365)?\\.com|cloud\\.microsoft)/',
  minExtensionVersion: '0.1',
  severity: 'low',
  type: patch.patchType,
  target: 'content',
  code: patch.patchCode,
  checksum: '',
  rollback: null,
  createdAt: new Date().toISOString(),
};

const patchesPath = path.join(ROOT, 'docs/patches/patches.json');
const manifest = JSON.parse(fs.readFileSync(patchesPath, 'utf8'));

if (manifest.patches.some((p) => p.id === patchId)) {
  console.log(`Patch ${patchId} already exists — skipping.`);
} else {
  manifest.patches.push(patchEntry);
  fs.writeFileSync(patchesPath, JSON.stringify(manifest, null, 2) + '\n');
  console.log(`Added patch ${patchId} to patches.json`);
}

// ── Write summary for workflow outputs ───────────────────────────────────────

console.log(`Patch type: ${patch.patchType}`);
console.log(`Explanation: ${patch.explanation || '(none)'}`);
