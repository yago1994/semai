#!/usr/bin/env node
// scripts/call-claude.js
// Usage: node scripts/call-claude.js <prompt-file> <response-file>
//
// Reads a plain-text prompt from <prompt-file>, calls the Anthropic Messages API,
// and writes the full response JSON to <response-file>.
// Exits non-zero on any HTTP or parse error so GitHub Actions can detect failure.

import fs from 'fs';
import https from 'https';

const [,, promptFile, responseFile] = process.argv;

if (!promptFile || !responseFile) {
  console.error('Usage: node scripts/call-claude.js <prompt-file> <response-file>');
  process.exit(1);
}

const apiKey = process.env.ANTHROPIC_API_KEY;
if (!apiKey) {
  console.error('ERROR: ANTHROPIC_API_KEY env var is not set.');
  process.exit(1);
}

const prompt = fs.readFileSync(promptFile, 'utf8');

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
    try {
      parsed = JSON.parse(raw);
    } catch (e) {
      console.error('ERROR: Could not parse Anthropic API response as JSON.');
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

    // Write the extracted text as the response file so the workflow can parse it directly.
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
