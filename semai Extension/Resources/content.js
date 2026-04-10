// content.js — patch injector (runs at document_start)
// Requests applicable patches from background, injects JS via main-world
// <script> tag and CSS via <style> tag. Deduplicates by patch ID.

(function () {
  const INJECTED_KEY = '__semai_injected_patches__';
  const PATCH_DEBUG = false;

  function semaiPatchDebug(...args) {
    if (PATCH_DEBUG) {
      console.warn(...args);
    }
  }

  function getInjected() {
    return window[INJECTED_KEY] || (window[INJECTED_KEY] = new Set());
  }

  function injectPatch(patch) {
    if (getInjected().has(patch.id)) return;
    getInjected().add(patch.id);

    if (patch.type === 'js') {
      const script = document.createElement('script');
      script.textContent = patch.code;
      script.dataset.semaiPatch = patch.id;
      (document.head || document.documentElement).appendChild(script);
    } else if (patch.type === 'css') {
      const style = document.createElement('style');
      style.textContent = patch.code;
      style.dataset.semaiPatch = patch.id;
      (document.head || document.documentElement).appendChild(style);
    }
  }

  chrome.runtime.sendMessage({ type: 'GET_PATCHES' }, (response) => {
    if (chrome.runtime.lastError) {
      semaiPatchDebug('[semai] Patch request failed:', chrome.runtime.lastError.message);
      return;
    }
    const patches = response?.patches ?? [];
    patches.forEach(injectPatch);
    if (patches.length > 0) {
      console.log(`[semai] Injected ${patches.length} patch(es).`);
    }
  });
})();
