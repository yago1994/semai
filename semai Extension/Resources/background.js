// background.js — patch loader + Xcode log relay service worker

// ── Patch loader ──────────────────────────────────────────────────────────────
// Fetches patches.json hourly, validates schema, caches in storage.
// Answers content-script queries with URL-filtered patches.

const PATCH_MANIFEST_URL =
  'https://yago1994.github.io/semai/patches/patches.json';

const STORAGE_KEY = 'semai_patches_cache';
const ALARM_NAME = 'semai_patch_fetch';
const FETCH_INTERVAL_MINUTES = 60;
const MAX_PATCH_BYTES = 50 * 1024; // 50 KB per patch
const EXTENSION_VERSION = chrome.runtime.getManifest().version;

chrome.runtime.onInstalled.addListener(() => {
  fetchAndCachePatches();
  chrome.alarms.create(ALARM_NAME, { periodInMinutes: FETCH_INTERVAL_MINUTES });
});

chrome.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === ALARM_NAME) fetchAndCachePatches();
});

async function fetchAndCachePatches() {
  try {
    const res = await fetch(PATCH_MANIFEST_URL, { cache: 'no-cache' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const manifest = await res.json();
    const valid = validateManifest(manifest);
    if (valid.length > 0) {
      await chrome.storage.local.set({ [STORAGE_KEY]: valid });
      console.log(`[semai] Cached ${valid.length} patch(es).`);
    }
  } catch (err) {
    console.warn('[semai] Patch fetch failed:', err.message);
  }
}

function validateManifest(manifest) {
  if (!Array.isArray(manifest?.patches)) return [];

  return manifest.patches.filter((p) => {
    if (
      typeof p.id !== 'string' ||
      typeof p.urlPattern !== 'string' ||
      typeof p.code !== 'string' ||
      !['js', 'css'].includes(p.type) ||
      !['content', 'background'].includes(p.target)
    ) {
      console.warn('[semai] Dropping malformed patch:', p.id ?? '(no id)');
      return false;
    }
    if (new Blob([p.code]).size > MAX_PATCH_BYTES) {
      console.warn(`[semai] Dropping oversized patch: ${p.id}`);
      return false;
    }
    if (p.minExtensionVersion && !semverSatisfies(p.minExtensionVersion)) {
      console.log(`[semai] Skipping patch ${p.id} (requires >= ${p.minExtensionVersion})`);
      return false;
    }
    return true;
  });
}

function semverSatisfies(required) {
  const parse = (v) => (v || '0.0.0').split('.').map(Number);
  const [rMaj, rMin, rPat] = parse(required);
  const [cMaj, cMin, cPat] = parse(EXTENSION_VERSION);
  if (cMaj !== rMaj) return cMaj > rMaj;
  if (cMin !== rMin) return cMin > rMin;
  return cPat >= rPat;
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type !== 'GET_PATCHES') return false;

  chrome.storage.local.get(STORAGE_KEY, (result) => {
    const all = result[STORAGE_KEY] ?? [];
    const pageUrl = sender.url ?? '';
    const applicable = all.filter((p) => {
      try {
        return new RegExp(p.urlPattern).test(pageUrl);
      } catch {
        return false;
      }
    });
    sendResponse({ patches: applicable });
  });

  return true; // keep message channel open for async response
});

// ── Live fix preview via Claude API ──────────────────────────────────────────
// Receives PREVIEW_FIX from contentScript.js, calls Claude with tool_use,
// returns the structured patch for live DOM injection.

const APPLY_FIX_TOOL = {
  name: 'apply_fix',
  description:
    'Apply a CSS or JS patch to fix the reported Outlook rendering issue. ' +
    'The patch code will be injected into the page via a <script> or <style> tag.',
  input_schema: {
    type: 'object',
    properties: {
      explanation: {
        type: 'string',
        description: 'One paragraph explaining what the issue is and how the patch fixes it.',
      },
      patchType: {
        type: 'string',
        enum: ['js', 'css'],
        description: 'Whether the fix is a JavaScript patch or a CSS patch.',
      },
      patchCode: {
        type: 'string',
        description:
          'The complete, self-contained patch code. ' +
          'JS patches are executed via a <script> tag in the page main world. ' +
          'CSS patches are injected as a <style> tag. ' +
          'JS patches MUST process existing DOM elements synchronously (querySelectorAll loop) ' +
          'and may additionally install a MutationObserver for future elements.',
      },
      urlPattern: {
        type: 'string',
        description:
          'A regex pattern matching Outlook URLs where this patch should apply. ' +
          'Example: ^https://outlook\\.(office(365)?\\.com|cloud\\.microsoft)/',
      },
    },
    required: ['explanation', 'patchType', 'patchCode', 'urlPattern'],
  },
};

const PREVIEW_FIX_SYSTEM_PROMPT = `You are a self-healing engine for a Safari Web Extension called "semai" (also known as "remou").
The extension transforms Outlook Web email threads into a chat-like interface.

When users report rendering issues, you produce a minimal CSS or JS patch to fix them.

## How patches work
- JS patches: injected as a <script> tag in the page's main world. They have full DOM/window access.
- CSS patches: injected as a <style> tag.
- Patches must be self-contained (no imports, no external dependencies).
- JS patches MUST process all matching elements synchronously with querySelectorAll on first execution.
  MutationObserver may be added for future elements, but the initial pass is mandatory.

## Outlook URL patterns
Outlook Web runs on these domains:
- outlook.office.com
- outlook.office365.com
- outlook.cloud.microsoft

Use this urlPattern to match all of them: ^https://outlook\\.(office(365)?\\.com|cloud\\.microsoft)/

## Guidelines
- Prefer CSS patches when the fix is purely visual (hiding, repositioning, styling).
- Use JS patches only when DOM manipulation is required (rewriting text, moving elements, etc.).
- Keep patches minimal — fix only the reported issue.
- Do not modify elements outside the scope of the bug report.`;

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type !== 'PREVIEW_FIX') return false;

  const { reason, cleanHtml, rawHtml, senderInfo, subject, pageUrl, anthropicApiKey } =
    msg.payload;

  if (!anthropicApiKey) {
    sendResponse({ ok: false, error: 'Anthropic API key not configured in secrets.js' });
    return true;
  }

  const userMessage = [
    `## Bug report`,
    `The user reported an issue while viewing: ${pageUrl}`,
    `Subject: ${subject || '(no subject)'}`,
    `Sender: ${senderInfo?.name || 'Unknown'} <${senderInfo?.email || 'unknown'}>`,
    ``,
    `## User's description`,
    reason || '(no description)',
    ``,
    `## Clean HTML (processed by extension)`,
    '```html',
    (cleanHtml || '').slice(0, 8000),
    '```',
    ``,
    `## Original HTML (raw from Outlook DOM)`,
    '```html',
    (rawHtml || '').slice(0, 8000),
    '```',
  ].join('\n');

  const body = JSON.stringify({
    model: 'claude-sonnet-4-6',
    max_tokens: 4096,
    system: PREVIEW_FIX_SYSTEM_PROMPT,
    tools: [APPLY_FIX_TOOL],
    tool_choice: { type: 'tool', name: 'apply_fix' },
    messages: [{ role: 'user', content: userMessage }],
  });

  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 60000);

  fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': anthropicApiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json',
    },
    body,
    signal: controller.signal,
  })
    .then((res) => {
      clearTimeout(timeout);
      if (!res.ok) return res.json().then((e) => Promise.reject(new Error(e.error?.message || `HTTP ${res.status}`)));
      return res.json();
    })
    .then((data) => {
      const toolUse = data.content?.find(
        (block) => block.type === 'tool_use' && block.name === 'apply_fix'
      );
      if (!toolUse) {
        sendResponse({ ok: false, error: 'Claude did not return a fix suggestion.' });
        return;
      }
      sendResponse({
        ok: true,
        explanation: toolUse.input.explanation,
        patchType: toolUse.input.patchType,
        patchCode: toolUse.input.patchCode,
        urlPattern: toolUse.input.urlPattern,
      });
    })
    .catch((err) => {
      clearTimeout(timeout);
      sendResponse({ ok: false, error: err.message || 'Preview fix request failed.' });
    });

  return true; // keep message channel open for async response
});

// ── Xcode log relay ───────────────────────────────────────────────────────────
// Receives { type: "semaiLog", text: "..." } from contentScript.js and
// forwards to the native host so the message appears in the Xcode console.
browser.runtime.onMessage.addListener((message) => {
  if (message && message.type === "semaiLog") {
    browser.runtime.sendNativeMessage("yam.team.remou", { log: message.text });
  }
});
