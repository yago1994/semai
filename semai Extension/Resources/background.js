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

// ── Xcode log relay ───────────────────────────────────────────────────────────
// Receives { type: "semaiLog", text: "..." } from contentScript.js and
// forwards to the native host so the message appears in the Xcode console.
browser.runtime.onMessage.addListener((message) => {
  if (message && message.type === "semaiLog") {
    browser.runtime.sendNativeMessage("yam.team.remou", { log: message.text });
  }
});
