#!/usr/bin/env node
// scripts/call-claude.js
// Assembles a prompt from issue context + source files, calls Claude, writes response.
//
// Usage: node scripts/call-claude.js <response-file> [prev-explanation-file]
//
// Required env vars:
//   ANTHROPIC_API_KEY, ISSUE_TITLE, ISSUE_BODY
//
// Optional env vars:
//   FAILING_TESTS_FILE   path to JSON file with test results (default: /tmp/baseline.json)
//   ATTEMPT              attempt number for logging (default: 1)

import fs from 'fs';
import https from 'https';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT = path.resolve(__dirname, '..');

const [,, responseFile, prevExplanationFile] = process.argv;

if (!responseFile) {
  console.error('Usage: node scripts/call-claude.js <response-file> [prev-explanation-file]');
  process.exit(1);
}

const apiKey = process.env.ANTHROPIC_API_KEY;
if (!apiKey) { console.error('ERROR: ANTHROPIC_API_KEY is not set.'); process.exit(1); }

const issueTitle = process.env.ISSUE_TITLE ?? '(no title)';
const issueBody  = process.env.ISSUE_BODY  ?? '(no body)';
const attempt    = process.env.ATTEMPT     ?? '1';

// ── Read optional context files ───────────────────────────────────────────────

const failingTestsFile = process.env.FAILING_TESTS_FILE ?? '/tmp/baseline.json';
let failingTests = '{}';
try { failingTests = fs.readFileSync(failingTestsFile, 'utf8'); } catch { /* none yet */ }

let prevExplanation = '';
if (prevExplanationFile) {
  try { prevExplanation = fs.readFileSync(prevExplanationFile, 'utf8').trim(); } catch { /* first attempt */ }
}

// ── Read source context ───────────────────────────────────────────────────────

function readSource(relPath, maxLines) {
  const abs = path.join(ROOT, relPath);
  try {
    const lines = fs.readFileSync(abs, 'utf8').split('\n');
    const slice = maxLines ? lines.slice(0, maxLines) : lines;
    return slice.join('\n');
  } catch {
    return `(file not found: ${relPath})`;
  }
}

const sourceFiles = [
  { path: 'semai Extension/Resources/background.js',    maxLines: null },
  { path: 'semai Extension/Resources/content.js',       maxLines: null },
  { path: 'docs/patches/patches.json',                  maxLines: null },
  { path: 'tests/test-cases.json',                      maxLines: null },
  { path: 'semai Extension/Resources/contentScript.js', maxLines: 200  }, // token budget
];

let sourceContext = '';
for (const { path: p, maxLines } of sourceFiles) {
  const label = maxLines ? `${p} (first ${maxLines} lines)` : p;
  sourceContext += `\n### ${label}\n\`\`\`\n${readSource(p, maxLines)}\n\`\`\`\n`;
}

// ── Build prompt ──────────────────────────────────────────────────────────────

const prevSection = prevExplanation
  ? `\n## Previous attempt explanation (attempt ${Number(attempt) - 1})\n${prevExplanation}\n`
  : '';

const prompt = `You are a self-healing Safari Web Extension auto-fix engine. Attempt ${attempt} of 3.

## Bug report
Title: ${issueTitle}
Body:
${issueBody}

## Failing tests (JSON)
${failingTests}
${prevSection}
## Source context
${sourceContext}

## Task
1. Analyse the bug report and failing tests.
2. Propose a minimal fix. The fix may be a source file change and/or a new patches.json entry.
3. Propose a new test case for tests/test-cases.json that captures this specific bug.

## Constraints for JS patches
- The test harness runs patch code via eval() against a pre-built jsdom DOM — there is no browser event loop.
- MutationObserver callbacks will NOT fire for elements that already exist in the DOM when the patch runs.
- JS patches MUST process all matching elements synchronously (e.g. querySelectorAll loop) immediately when executed, in addition to any MutationObserver for live pages.
- Do not rely solely on MutationObserver — always also do an initial synchronous pass over existing elements.

Respond with ONLY valid JSON — no markdown fences, no text outside the JSON object:
{
  "explanation": "<one paragraph describing the fix>",
  "files": [
    { "path": "<repo-relative path>", "content": "<complete updated file content>" }
  ],
  "newTestCase": { "id": "...", "description": "...", "fixture": "...", "patch": null, "assertions": [] },
  "patchEntry": null
}

Use "files": [] if no file changes are needed. Set "newTestCase" or "patchEntry" to null if not applicable.`;

// ── Call Anthropic API ────────────────────────────────────────────────────────

const requestBody = JSON.stringify({
  model: 'claude-sonnet-4-6',
  max_tokens: 4096,
  messages: [{ role: 'user', content: prompt }]
});

const options = {
  hostname: 'api.anthropic.com',
  path: '/v1/messages',
  method: 'POST',
  headers: {
    'x-api-key': apiKey,
    'anthropic-version': '2023-06-01',
    'content-type': 'application/json',
    'content-length': Buffer.byteLength(requestBody)
  }
};

console.log(`[call-claude] Attempt ${attempt} — sending prompt (${prompt.length} chars)`);

const req = https.request(options, (res) => {
  let raw = '';
  res.on('data', (chunk) => { raw += chunk; });
  res.on('end', () => {
    if (res.statusCode !== 200) {
      console.error(`ERROR: Anthropic API returned HTTP ${res.statusCode}`);
      console.error(raw.slice(0, 500));
      process.exit(1);
    }

    let parsed;
    try { parsed = JSON.parse(raw); } catch {
      console.error('ERROR: Could not parse Anthropic response as JSON.');
      console.error(raw.slice(0, 500));
      process.exit(1);
    }

    if (parsed.type === 'error') {
      console.error(`ERROR from Anthropic API: ${parsed.error?.message ?? JSON.stringify(parsed.error)}`);
      process.exit(1);
    }

    const text = parsed.content?.[0]?.text;
    if (typeof text !== 'string') {
      console.error('ERROR: Unexpected response shape — no content[0].text');
      console.error(JSON.stringify(parsed).slice(0, 500));
      process.exit(1);
    }

    fs.writeFileSync(responseFile, text, 'utf8');
    console.log(`[call-claude] Response written to ${responseFile} (${text.length} chars)`);
  });
});

req.on('error', (err) => {
  console.error(`ERROR: HTTPS request failed: ${err.message}`);
  process.exit(1);
});

req.setTimeout(120000, () => {
  console.error('ERROR: Anthropic API request timed out after 120s.');
  req.destroy();
  process.exit(1);
});

req.write(requestBody);
req.end();
