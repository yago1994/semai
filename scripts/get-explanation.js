#!/usr/bin/env node
// scripts/get-explanation.js
// Reads /tmp/claude_response.json and prints the explanation field to stdout.
// Used by the workflow to extract Claude's explanation for issue comments
// and for passing context to the next retry attempt.

import fs from 'fs';

const responseFile = process.argv[2] ?? '/tmp/claude_response.json';
const outputFile   = process.argv[3]; // optional: write to file instead of stdout

let explanation = '(could not parse last response)';
try {
  const r = JSON.parse(fs.readFileSync(responseFile, 'utf8'));
  explanation = r.explanation ?? '(no explanation provided)';
} catch { /* use default */ }

if (outputFile) {
  fs.writeFileSync(outputFile, explanation, 'utf8');
} else {
  process.stdout.write(explanation);
}
