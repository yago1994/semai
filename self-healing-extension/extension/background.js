// background.js — Patch loader for Safari Web Extension
// Periodically fetches patches.json from your GitHub Pages site,
// validates and caches them, then makes them available to content scripts.

const PATCH_MANIFEST_URL = "https://YOUR_USERNAME.github.io/YOUR_REPO/patches.json";
const FETCH_INTERVAL_MIN = 60; // check hourly
const STORAGE_KEY = "she.patchManifest";
const STORAGE_KEY_LAST_FETCH = "she.lastFetchAt";
const EXTENSION_VERSION = browser.runtime.getManifest().version;

// ---- Manifest validation ---------------------------------------------------

function isValidManifest(m) {
  if (!m || typeof m !== "object") return false;
  if (m.manifestVersion !== 1) return false;
  if (!Array.isArray(m.patches)) return false;
  return m.patches.every(isValidPatch);
}

function isValidPatch(p) {
  if (!p || typeof p !== "object") return false;
  if (typeof p.id !== "string") return false;
  if (!["js", "css"].includes(p.type)) return false;
  if (!["content", "background"].includes(p.target)) return false;
  if (typeof p.code !== "string") return false;
  // Hard cap on code size to limit blast radius of a compromised manifest
  if (p.code.length > 50_000) return false;
  return true;
}

// ---- Version matching ------------------------------------------------------

function semverGte(a, b) {
  const pa = a.split(".").map(Number);
  const pb = b.split(".").map(Number);
  for (let i = 0; i < 3; i++) {
    if ((pa[i] || 0) > (pb[i] || 0)) return true;
    if ((pa[i] || 0) < (pb[i] || 0)) return false;
  }
  return true;
}

function patchAppliesToThisVersion(patch) {
  if (!patch.matches?.extensionVersions) return true;
  return patch.matches.extensionVersions.includes(EXTENSION_VERSION);
}

// ---- Fetch + cache ---------------------------------------------------------

async function fetchManifest() {
  try {
    const res = await fetch(PATCH_MANIFEST_URL, {
      cache: "no-cache",
      headers: { Accept: "application/json" },
    });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const manifest = await res.json();

    if (!isValidManifest(manifest)) {
      console.warn("[SHE] Invalid patch manifest, ignoring");
      return null;
    }
    if (manifest.minExtensionVersion && !semverGte(EXTENSION_VERSION, manifest.minExtensionVersion)) {
      console.info("[SHE] Extension version too old for this manifest");
      return null;
    }

    await browser.storage.local.set({
      [STORAGE_KEY]: manifest,
      [STORAGE_KEY_LAST_FETCH]: Date.now(),
    });
    console.info(`[SHE] Loaded ${manifest.patches.length} patch(es)`);
    return manifest;
  } catch (err) {
    console.warn("[SHE] Failed to fetch patch manifest:", err.message);
    return null;
  }
}

async function getCachedManifest() {
  const stored = await browser.storage.local.get(STORAGE_KEY);
  return stored[STORAGE_KEY] || null;
}

// ---- Message handler for content scripts -----------------------------------

browser.runtime.onMessage.addListener(async (msg) => {
  if (msg?.type === "SHE_GET_PATCHES_FOR_URL") {
    const manifest = (await getCachedManifest()) || (await fetchManifest());
    if (!manifest) return { patches: [] };

    const applicable = manifest.patches.filter((p) => {
      if (!patchAppliesToThisVersion(p)) return false;
      if (p.target !== "content") return false;
      const pattern = p.matches?.urlPattern;
      if (!pattern) return true;
      try {
        return new RegExp(pattern).test(msg.url);
      } catch {
        return false;
      }
    });
    return { patches: applicable };
  }
});

// ---- Schedule periodic fetches ---------------------------------------------

browser.alarms.create("she-fetch-patches", { periodInMinutes: FETCH_INTERVAL_MIN });
browser.alarms.onAlarm.addListener((alarm) => {
  if (alarm.name === "she-fetch-patches") fetchManifest();
});

// Fetch on startup too
fetchManifest();
