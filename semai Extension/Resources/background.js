// background.js — patch loader + Xcode log relay + live fix preview

// ── Patch loader ──────────────────────────────────────────────────────────────
// Fetches patches.json hourly, validates schema, caches in storage.

const PATCH_MANIFEST_URL =
  'https://yago1994.github.io/semai/patches/patches.json';

const STORAGE_KEY = 'semai_patches_cache';
const ANALYTICS_INSTALL_ID_KEY = 'semai_install_id';
const ALARM_NAME = 'semai_patch_fetch';
const FETCH_INTERVAL_MINUTES = 60;
const MAX_PATCH_BYTES = 50 * 1024; // 50 KB per patch
const EXTENSION_VERSION = chrome.runtime.getManifest().version;
const ANALYTICS_URL = 'https://script.google.com/macros/s/AKfycbxTfl5yMdqcHAmXjRmvGIBQ_ILMmR6XX7vyUolxjvR2h17f0cCbsUAFnEdqcw9GMNfL/exec';
const PATCH_DEBUG = false;

function semaiPatchDebug(...args) {
  if (PATCH_DEBUG) {
    console.warn(...args);
  }
}

// Pre-load sig detector source at startup so the PREVIEW_FIX handler
// doesn't need to fetch it on every request (which crashes Safari's SW).
let cachedSigDetectorSource = null;
fetch(chrome.runtime.getURL('semaiSigDetector.js'))
  .then(r => r.text())
  .then(src => {
    cachedSigDetectorSource = src;
    console.log('[semai] semaiSigDetector.js loaded — ' + src.length + ' chars');
  })
  .catch(err => console.warn('[semai] Failed to load semaiSigDetector.js:', err.message));

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
    semaiPatchDebug('[semai] Patch fetch failed:', err.message);
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
      semaiPatchDebug('[semai] Dropping malformed patch:', p.id ?? '(no id)');
      return false;
    }
    if (new Blob([p.code]).size > MAX_PATCH_BYTES) {
      semaiPatchDebug(`[semai] Dropping oversized patch: ${p.id}`);
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

async function semaiLoadOrCreateInstallID() {
  const stored = await chrome.storage.local.get(ANALYTICS_INSTALL_ID_KEY);
  const existing = stored?.[ANALYTICS_INSTALL_ID_KEY];
  if (typeof existing === 'string' && existing.length > 0) {
    return existing;
  }

  const installID = crypto.randomUUID();
  await chrome.storage.local.set({ [ANALYTICS_INSTALL_ID_KEY]: installID });
  return installID;
}

async function semaiTrackEvent(eventName, details = {}) {
  try {
    const payload = {
      install_id: await semaiLoadOrCreateInstallID(),
      extension_version: EXTENSION_VERSION,
      platform: navigator.userAgent,
      event: eventName,
      ...details
    };

    await fetch(ANALYTICS_URL, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
  } catch (error) {
    semaiPatchDebug('[semai] Analytics event failed:', eventName, error?.message || error);
  }
}

// ── Live fix preview — Claude API tool definition ────────────────────────────

const APPLY_FIX_TOOL = {
  name: 'apply_fix',
  description:
    'Apply a CSS or JS patch to fix the reported Outlook rendering issue.',
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
          'A regex pattern matching Outlook URLs where this patch should apply.',
      },
    },
    required: ['explanation', 'patchType', 'patchCode', 'urlPattern'],
  },
};

const PREVIEW_FIX_SYSTEM_PROMPT = [
  'You are a self-healing engine for a Safari Web Extension called "semai" (also known as "remou").',
  'The extension transforms Outlook Web email threads into a chat-like interface.',
  '',
  'When users report rendering issues, you produce a minimal CSS or JS patch to fix them.',
  '',
  '## How patches work',
  '- JS patches: injected into the page main world via chrome.scripting.executeScript. They have full DOM/window access.',
  '- CSS patches: injected via chrome.scripting.insertCSS.',
  '- Patches must be self-contained (no imports, no external dependencies).',
  '- JS patches MUST process all matching elements synchronously with querySelectorAll on first execution.',
  '  MutationObserver may be added for future elements, but the initial pass is mandatory.',
  '',
  '## Extension DOM structure',
  'The extension renders a chat overlay with these key elements:',
  '- #semai-chat-overlay — the main overlay container',
  '- .semai-chat-row — one row per email message (has data-report-index attribute)',
  '- .semai-chat-row.semai-chat-me — rows sent by the current user',
  '- .semai-chat-row.semai-chat-them — rows sent by others',
  '- .semai-chat-avatar — the avatar circle showing sender initials (e.g. "GS" for "Gaelle Sabben")',
  '- .semai-chat-bubble — the message bubble containing the email body',
  '- .semai-chat-sender — sender name label shown above the bubble (on .semai-chat-them rows)',
  '',
  '## Outlook URL patterns',
  'Outlook Web runs on these domains: outlook.office.com, outlook.office365.com, outlook.cloud.microsoft',
  'Use this urlPattern to match all: ^https://outlook\\\\.(office(365)?\\\\.com|cloud\\\\.microsoft)/',
  '',
  '## Guidelines',
  '- Prefer CSS patches when the fix is purely visual.',
  '- Use JS patches only when DOM manipulation is required.',
  '- Keep patches minimal — fix only the reported issue.',
  '- When the rendered HTML is provided, base your selectors on the ACTUAL class names visible in it.',
].join('\n');

// ── Single chrome.runtime message handler ────────────────────────────────────
// Uses chrome.runtime (same API as the working GET_PATCHES handler).
// All message types handled in one listener to avoid channel conflicts.

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (!msg || !msg.type) return false;

  // ── GET_PATCHES ──
  if (msg.type === 'GET_PATCHES') {
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
    return true;
  }

  // ── Xcode log relay ──
  if (msg.type === 'semaiLog') {
    try {
      browser.runtime.sendNativeMessage('yam.team.remou', { log: msg.text });
    } catch { /* ignore */ }
    return false;
  }

  if (msg.type === 'OPEN_ONBOARDING_APP') {
    browser.runtime.sendNativeMessage(
      'yam.team.remou',
      { command: 'open-onboarding-app' },
      (response) => {
        if (chrome.runtime.lastError) {
          sendResponse({ ok: false, error: chrome.runtime.lastError.message });
          return;
        }

        sendResponse(response && typeof response === 'object' ? response : { ok: true });
      }
    );
    return true;
  }

  if (msg.type === 'TRACK_EVENT') {
    semaiTrackEvent(msg.eventName, msg.details)
      .then(() => sendResponse({ ok: true }))
      .catch((error) => sendResponse({ ok: false, error: error.message }));
    return true;
  }

  // ── Patch injection (bypasses page CSP via chrome.scripting) ──
  if (msg.type === 'INJECT_PATCH') {
    const { patchType, patchCode } = msg.payload || {};
    const tabId = sender.tab?.id;
    if (!tabId) {
      sendResponse({ ok: false, error: 'No tab ID available' });
      return false;
    }
    const target = { tabId };
    if (patchType === 'css') {
      chrome.scripting.insertCSS({ target, css: patchCode })
        .then(() => sendResponse({ ok: true }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;
    }
    if (patchType === 'js') {
      chrome.scripting.executeScript({
        target,
        world: 'MAIN',
        func: (code) => { (0, eval)(code); },
        args: [patchCode],
      })
        .then(() => sendResponse({ ok: true }))
        .catch((err) => sendResponse({ ok: false, error: err.message }));
      return true;
    }
    sendResponse({ ok: false, error: 'Unknown patchType: ' + patchType });
    return false;
  }

  // ── Remove CSS patch (undo preview) ──
  if (msg.type === 'REMOVE_CSS_PATCH') {
    const { css } = msg.payload || {};
    const tabId = sender.tab?.id;
    if (tabId && css) {
      chrome.scripting.removeCSS({ target: { tabId }, css }).catch(() => {});
    }
    return false;
  }

  // ── Live fix preview ──
  if (msg.type === 'PREVIEW_FIX') {
    const { reason, cleanHtml, rawHtml, renderedHtml, senderInfo, subject, pageUrl, anthropicApiKey, conversationHistory } =
      msg.payload || {};

    console.log('[semai-preview] Received PREVIEW_FIX message');

    if (!anthropicApiKey) {
      sendResponse({ ok: false, error: 'Anthropic API key not configured in secrets.js' });
      return false;
    }

    // Wrap in async IIFE — the onMessage callback itself must stay synchronous
    // (it returns `true` below to keep the channel open).
    (async () => {
      console.log('[semai-preview] IIFE started');
      const sigDetectorSource = cachedSigDetectorSource;
      console.log('[semai-preview] sigDetectorSource available:', !!sigDetectorSource);

      const userMessage = [
        '## Bug report',
        'The user reported an issue while viewing: ' + (pageUrl || ''),
        'Subject: ' + (subject || '(no subject)'),
        'Sender: ' + (senderInfo?.name || 'Unknown') + ' <' + (senderInfo?.email || 'unknown') + '>',
        '',
        '## User description',
        reason || '(no description)',
        '',
        // TODO: re-enable once end-to-end flow is verified
        // ...(sigDetectorSource ? ['## Extension source: semaiSigDetector.js', '```javascript', sigDetectorSource, '```', ''] : []),
      ].join('\n');

      // Build the messages array: initial user prompt + any retry history
      const messages = [
        { role: 'user', content: userMessage },
        ...(Array.isArray(conversationHistory) ? conversationHistory : []),
      ];

      const body = JSON.stringify({
        model: 'claude-sonnet-4-6',
        max_tokens: 8000,
        system: PREVIEW_FIX_SYSTEM_PROMPT,
        tools: [APPLY_FIX_TOOL],
        tool_choice: { type: 'tool', name: 'apply_fix' },
        messages,
      });

      console.log('[semai-preview] Calling Anthropic API — turns:', messages.length, 'body:', body.length, 'chars');
      console.log('[semai-preview] First 200 chars of body:', body.slice(0, 200));

      const res = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'x-api-key': anthropicApiKey,
          'anthropic-version': '2023-06-01',
          'anthropic-dangerous-direct-browser-access': 'true',
          'content-type': 'application/json',
        },
        body,
      });

      console.log('[semai-preview] Response status:', res.status);

      if (!res.ok) {
        const t = await res.text();
        console.error('[semai-preview] Error body:', t.slice(0, 500));
        const parsed = (() => { try { return JSON.parse(t); } catch { return {}; } })();
        sendResponse({ ok: false, error: parsed.error?.message || 'HTTP ' + res.status });
        return;
      }

      const data = await res.json();
      console.log('[semai-preview] Full API response:', JSON.stringify(data));

      const toolUse = data.content?.find(b => b.type === 'tool_use' && b.name === 'apply_fix');
      if (!toolUse) {
        console.warn('[semai-preview] No tool_use block in response');
        sendResponse({ ok: false, error: 'Claude did not return a fix suggestion.' });
        return;
      }

      console.log('[semai-preview] Got patch:', toolUse.input.patchType, '— toolUseId:', toolUse.id);
      sendResponse({
        ok: true,
        toolUseId: toolUse.id,
        explanation: toolUse.input.explanation,
        patchType: toolUse.input.patchType,
        patchCode: toolUse.input.patchCode,
        urlPattern: toolUse.input.urlPattern,
      });
    })().catch((err) => {
      console.error('[semai-preview] Error: ' + err.message);
      sendResponse({ ok: false, error: err.message || 'Preview fix request failed.' });
    });

    return true; // keep channel open for async response
  }

  return false;
});
