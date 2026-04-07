// content.js — Runtime patch applier
// Asks the background script which patches apply to this URL,
// then injects them into the page. Tracks applied patches to avoid duplicates.

const APPLIED = new Set();

async function applyPatches() {
  let response;
  try {
    response = await browser.runtime.sendMessage({
      type: "SHE_GET_PATCHES_FOR_URL",
      url: window.location.href,
    });
  } catch (err) {
    console.warn("[SHE] Could not reach background:", err);
    return;
  }

  const patches = response?.patches || [];
  for (const patch of patches) {
    if (APPLIED.has(patch.id)) continue;
    try {
      if (patch.type === "css") {
        injectCss(patch);
      } else if (patch.type === "js") {
        injectJs(patch);
      }
      APPLIED.add(patch.id);
      console.info(`[SHE] Applied patch ${patch.id}`);
    } catch (err) {
      console.error(`[SHE] Patch ${patch.id} failed:`, err);
    }
  }
}

function injectCss(patch) {
  const style = document.createElement("style");
  style.dataset.shePatchId = patch.id;
  style.textContent = patch.code;
  (document.head || document.documentElement).appendChild(style);
}

function injectJs(patch) {
  // Inject into the page's main world via a <script> element so the patch
  // can interact with page-level globals. The content script's isolated world
  // can't see those directly.
  const script = document.createElement("script");
  script.dataset.shePatchId = patch.id;
  script.textContent = `try { ${patch.code} } catch (e) { console.error('[SHE] patch ${patch.id}', e); }`;
  (document.head || document.documentElement).appendChild(script);
  script.remove(); // tag stays in console; element no longer needed
}

// Apply early, then again after DOM is ready in case patches depend on it
applyPatches();
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", applyPatches, { once: true });
}
