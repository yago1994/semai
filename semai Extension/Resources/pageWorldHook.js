// pageWorldHook.js — runs in the PAGE's main world (not the isolated
// content-script world) so that wrapping window.fetch / XMLHttpRequest
// actually sees Outlook's API traffic. Loaded via <script src=...> from
// content.js at document_start, before Outlook's SPA bundles execute.
//
// On each captured bearer token, dispatches a CustomEvent on `document`
// that the content script listens for. We never talk directly to the
// isolated world — CustomEvent is the one-way bridge that works across
// both worlds.

(function semaiPageWorldTokenHook() {
  if (window.__semaiTokenHookInstalled) return;
  window.__semaiTokenHookInstalled = true;

  var HOSTS = [
    "outlook.office.com",
    "outlook.office365.com",
    "outlook.live.com",
    "outlook.cloud.microsoft",
    "substrate.office.com",
    "graph.microsoft.com"
  ];

  function isOutlookUrl(urlStr) {
    if (!urlStr) return false;
    try {
      var u = new URL(urlStr, window.location.href);
      for (var i = 0; i < HOSTS.length; i++) {
        if (u.hostname === HOSTS[i] || u.hostname.endsWith("." + HOSTS[i])) return true;
      }
      return false;
    } catch (_) {
      return false;
    }
  }

  function publishToken(token) {
    if (typeof token !== "string") return;
    if (!/^Bearer\s+\S+/i.test(token)) return;
    try {
      document.dispatchEvent(new CustomEvent("semai-outlook-token", {
        detail: { token: token, at: Date.now() }
      }));
    } catch (_) {
      // Never throw from a hook.
    }
  }

  // ----- fetch hook -----
  try {
    var originalFetch = window.fetch;
    if (typeof originalFetch === "function") {
      window.fetch = function semaiHookedFetch(input, init) {
        try {
          var url = typeof input === "string" ? input : (input && input.url) || "";
          if (isOutlookUrl(url)) {
            var auth = "";
            if (init && init.headers) {
              if (init.headers instanceof Headers) {
                auth = init.headers.get("Authorization") || init.headers.get("authorization") || "";
              } else if (Array.isArray(init.headers)) {
                for (var i = 0; i < init.headers.length; i++) {
                  var h = init.headers[i];
                  if (Array.isArray(h) && /^authorization$/i.test(h[0])) { auth = h[1]; break; }
                }
              } else if (typeof init.headers === "object") {
                auth = init.headers.Authorization || init.headers.authorization || "";
              }
            } else if (input && typeof input === "object" && input.headers && typeof input.headers.get === "function") {
              try { auth = input.headers.get("Authorization") || input.headers.get("authorization") || ""; } catch (_) {}
            }
            if (auth) publishToken(auth);
          }
        } catch (_) {}
        return originalFetch.apply(this, arguments);
      };
    }
  } catch (_) {}

  // ----- XHR hook -----
  try {
    var XHR = window.XMLHttpRequest;
    if (XHR && XHR.prototype) {
      var origOpen = XHR.prototype.open;
      var origSetHeader = XHR.prototype.setRequestHeader;

      XHR.prototype.open = function semaiHookedXhrOpen(method, url) {
        try { this.__semaiUrl = url; } catch (_) {}
        return origOpen.apply(this, arguments);
      };

      XHR.prototype.setRequestHeader = function semaiHookedXhrHeader(name, value) {
        try {
          if (typeof name === "string" &&
              /^authorization$/i.test(name) &&
              isOutlookUrl(this.__semaiUrl) &&
              typeof value === "string" &&
              /^Bearer\s+/i.test(value)) {
            publishToken(value);
          }
        } catch (_) {}
        return origSetHeader.apply(this, arguments);
      };
    }
  } catch (_) {}
})();
