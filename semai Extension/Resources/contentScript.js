// =====================================================================
// SEMAI REST-API REPLY (ADDITIVE — DOES NOT REPLACE COMPOSE-UI FLOW)
// ---------------------------------------------------------------------
// The compose-UI reply path (semaiOpenReplyAllCompose / semaiInsertComposeText
// / semaiSendCompose) is preserved exactly as-is. It is used as the fallback
// whenever any of the following are true:
//   - SEMAI_USE_REST_API_REPLY is false (kill-switch)
//   - No OAuth bearer token has been captured yet
//   - The message ID cannot be determined from the DOM
//   - The REST POST itself fails (network, 4xx, 5xx)
//
// To revert entirely, flip SEMAI_USE_REST_API_REPLY to false. The existing
// compose-UI path will run unchanged.
// =====================================================================
const SEMAI_USE_REST_API_REPLY = true;

// Module-level cache. Outlook's OAuth bearer is short-lived (usually ~1h) but
// is renewed continuously by the SPA, so passive hook capture is reliable as
// long as the user keeps Outlook open. We refresh the cached token every time
// we see a fresh Authorization header on an Outlook API call.
let semaiCachedOutlookToken = "";
let semaiCachedOutlookTokenAt = 0;

// Endpoints we can use the captured token against. We accept tokens seen on
// any of these hosts because Outlook fetches them all with the same bearer.
const SEMAI_OUTLOOK_API_HOSTS = [
  "outlook.office.com",
  "outlook.office365.com",
  "substrate.office.com",
  "graph.microsoft.com"
];

function semaiIsOutlookApiUrl(urlStr) {
  if (!urlStr) return false;
  try {
    const u = new URL(urlStr, window.location.href);
    return SEMAI_OUTLOOK_API_HOSTS.some(host => u.hostname.endsWith(host));
  } catch (_) {
    return false;
  }
}

function semaiRecordCapturedToken(rawAuthHeader, sourceLabel) {
  if (typeof rawAuthHeader !== "string") return;
  const trimmed = rawAuthHeader.trim();
  if (!/^Bearer\s+\S+/i.test(trimmed)) return;
  if (trimmed === semaiCachedOutlookToken) return;
  semaiCachedOutlookToken = trimmed;
  semaiCachedOutlookTokenAt = Date.now();
  try {
    if (typeof semaiNativeLog === "function") {
      const tokenLen = trimmed.length;
      semaiNativeLog(`[semai-rest] captured Outlook bearer token (source=${sourceLabel}, len=${tokenLen})`);
    }
    semaiDebugLine(`✓ token captured (${sourceLabel}, len=${trimmed.length})`);
  } catch (_) {
    // semaiNativeLog / semaiDebugLine may not be defined on the very first hook call; ignore.
  }
}

// Listen for bearer tokens captured by pageWorldHook.js.
//
// Content scripts run in an isolated world, so wrapping window.fetch /
// XMLHttpRequest here would only wrap OUR fetch — Outlook's own traffic
// would bypass us entirely. Instead, pageWorldHook.js (injected at
// document_start via content.js) runs in the page's main world, wraps the
// real fetch/XHR, and pipes captured bearer tokens back to us via a
// CustomEvent on `document`. This listener just caches whatever it hears.
document.addEventListener("semai-outlook-token", (event) => {
  const token = event && event.detail && event.detail.token;
  if (typeof token === "string") {
    semaiRecordCapturedToken(token, "page-world-hook");
  }
});

// Walk DOM ancestors of the message body looking for a stable Outlook item
// ID. Outlook stamps message IDs into a few different attributes depending on
// the SPA build — we try several patterns and return the first match. On
// failure we dump a *full* attribute summary of the ancestor chain into the
// chat debug panel so we can see exactly what is available in the user's
// build and tune the extractor.
function semaiExtractMessageId(bodyEl) {
  if (!(bodyEl instanceof Element)) return "";

  const directAttrs = ["data-convid", "data-itemid", "data-message-id", "data-itemserverid"];
  const chainDump = [];
  let node = bodyEl;
  for (let d = 0; d < 16 && node; d++, node = node.parentElement) {
    const attrNames = node.getAttributeNames ? node.getAttributeNames() : [];

    // 1) Direct attribute hit
    for (const attr of directAttrs) {
      const v = node.getAttribute && node.getAttribute(attr);
      if (v && v.length > 8) {
        try {
          semaiNativeLog(`[semai-rest] messageId via [${attr}] at depth=${d}`);
          semaiDebugLine(`msgId via [${attr}] at d=${d}`);
        } catch (_) {}
        return v;
      }
    }

    // 2) ID matches the EWS base64 prefix
    const elId = node.id || "";
    const idMatch = elId.match(/(AAMkA[A-Za-z0-9_\-+/=]{20,})/);
    if (idMatch) {
      try {
        semaiNativeLog(`[semai-rest] messageId via id-pattern at depth=${d}`);
        semaiDebugLine(`msgId via id-pattern at d=${d}`);
      } catch (_) {}
      return idMatch[1];
    }

    // 3) Any data-* attribute value matches the EWS base64 prefix
    for (const name of attrNames) {
      if (!/^data-/.test(name)) continue;
      const v = node.getAttribute(name);
      const m = v && typeof v === "string" ? v.match(/(AAMkA[A-Za-z0-9_\-+/=]{20,})/) : null;
      if (m) {
        try {
          semaiNativeLog(`[semai-rest] messageId via ${name} at depth=${d}`);
          semaiDebugLine(`msgId via ${name} at d=${d}`);
        } catch (_) {}
        return m[1];
      }
    }

    // 4) Build a verbose dump for diagnostics. Includes ALL attribute names
    // and the first 60 chars of each value, so we can spot any candidate.
    const attrDump = attrNames.map((name) => {
      let v = node.getAttribute(name) || "";
      if (v.length > 60) v = v.slice(0, 60) + "…";
      return `${name}=${v}`;
    });
    chainDump.push(`d=${d} <${(node.tagName || "?").toLowerCase()}> ${attrDump.join(" ")}`);
  }

  try {
    semaiNativeLog(`[semai-rest] messageId NOT FOUND — ancestor chain dump follows`);
    semaiDebugLine("---- DOM CHAIN (for diagnostics) ----");
    chainDump.forEach((line) => semaiDebugLine(line));
    semaiDebugLine("---- end chain ----");
  } catch (_) {}
  return "";
}

// Wrapper that routes Outlook REST API calls through the background service
// worker. Safari's content-script context blocks Authorization-bearing fetches
// to outlook.office.com with "Load failed", but background workers have full
// host_permissions and succeed. Returns { ok, status, body, error }.
function semaiCallOutlookApi(url, method, body) {
  return new Promise((resolve) => {
    if (!semaiCachedOutlookToken) {
      resolve({ ok: false, status: 0, error: "No Outlook bearer token captured yet." });
      return;
    }
    try {
      chrome.runtime.sendMessage(
        {
          type: "OUTLOOK_API_CALL",
          payload: { url, method: method || "GET", token: semaiCachedOutlookToken, body: body || null }
        },
        (response) => {
          if (chrome.runtime.lastError) {
            resolve({ ok: false, status: 0, error: chrome.runtime.lastError.message });
            return;
          }
          resolve(response || { ok: false, status: 0, error: "No response from background worker" });
        }
      );
    } catch (err) {
      resolve({ ok: false, status: 0, error: err && err.message ? err.message : String(err) });
    }
  });
}

async function semaiPostOutlookReplyAll(messageId, comment) {
  if (!messageId) {
    throw new Error("No Outlook message ID resolved.");
  }

  const url = `https://outlook.office.com/api/v2.0/me/messages/${encodeURIComponent(messageId)}/replyAll`;
  const response = await semaiCallOutlookApi(url, "POST", { Comment: comment });

  if (response.ok && (response.status === 202 || response.status === 200 || response.status === 204)) {
    return true;
  }

  if (response.error) {
    throw new Error(`Outlook REST replyAll failed: ${response.error}`);
  }
  throw new Error(`Outlook REST replyAll failed: ${response.status} ${response.statusText || ""} ${(response.body || "").slice(0, 200)}`);
}

// Pull a small, distinctive normalized snippet from a message body to compare
// against the API's BodyPreview. We strip HTML, collapse whitespace, lowercase,
// and pick a 30-char window from the middle of the first usable line. The
// middle is more likely to be unique than a greeting/signature.
function semaiBodySnippet(bodyText) {
  if (!bodyText) return "";
  const stripped = String(bodyText)
    .replace(/<[^>]+>/g, " ")
    .replace(/&[a-z]+;|&#\d+;/gi, " ")
    .replace(/[\r\n]+/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
  if (stripped.length < 12) return "";
  // If long enough, drop the first 20 chars to skip "hi joe," etc.
  const start = stripped.length > 60 ? 20 : 0;
  return stripped.slice(start, start + 50);
}

// Resolves the messageId we should reply-all to, given the thread we're
// looking at. Strategy:
//   1. Smoke test the token.
//   2. List the most recent N messages from the user's mailbox by date (no
//      $search — search has a 5–60s indexing lag and may miss the user's
//      just-sent reply or recent activity).
//   3. Compare each chat-overlay message body against the returned
//      messages' BodyPreview. The OLDEST chat-overlay messages are most
//      likely to be indexed and least likely to collide with the user's
//      drafts. First match → take its ConversationId.
//   4. Of all the recent messages we just listed, pick the one with that
//      ConversationId AND the latest ReceivedDateTime. That's our reply
//      target.
async function semaiResolveMessageIdViaRest(threadMessages) {
  if (!semaiCachedOutlookToken) {
    throw new Error("No Outlook bearer token captured yet.");
  }

  // ----- Smoke test -----
  semaiDebugLine("REST: smoke test GET /me/messages?$top=1");
  const smoke = await semaiCallOutlookApi("https://outlook.office.com/api/v2.0/me/messages?$top=1", "GET");
  if (!smoke.ok) {
    semaiDebugLine(`REST: smoke FAILED status=${smoke.status}`);
    semaiDebugLine(`REST: smoke body: ${(smoke.error || smoke.body || "").slice(0, 400)}`);
    throw new Error(`Smoke test failed (${smoke.status}): ${(smoke.error || smoke.body || "").slice(0, 200)}`);
  }
  semaiDebugLine("REST: smoke OK ✓");

  // ----- List recent messages -----
  // 50 covers most active mailboxes' last few days. We deliberately do NOT
  // pass $select since this tenant rejects some fields (ConversationTopic)
  // and the default response includes everything we need.
  const url = `https://outlook.office.com/api/v2.0/me/messages?$orderby=${encodeURIComponent("ReceivedDateTime desc")}&$top=50`;
  semaiDebugLine("REST: listing 50 most recent messages…");
  const res = await semaiCallOutlookApi(url, "GET");
  if (!res.ok) {
    semaiDebugLine(`REST: list FAILED status=${res.status}`);
    semaiDebugLine(`REST: list body: ${(res.error || res.body || "").slice(0, 400)}`);
    throw new Error(`Recent-messages list failed (${res.status}).`);
  }
  let data = {};
  try { data = JSON.parse(res.body || "{}"); } catch (_) {}
  const recent = (data && data.value) || [];
  semaiDebugLine(`REST: got ${recent.length} recent messages`);
  if (recent.length === 0) {
    throw new Error("Recent-messages list returned 0 results.");
  }

  // ----- Match against chat-overlay messages -----
  // Iterate from the OLDEST chat-overlay messages forward — those are
  // guaranteed indexed, and they don't include the user's just-typed reply.
  const overlayMessages = Array.isArray(threadMessages) ? threadMessages : [];
  const norm = (s) => (s || "").replace(/\s+/g, " ").trim().toLowerCase();
  const recentNormPreviews = recent.map((m) => norm(m.BodyPreview));

  let conversationId = "";
  let pivotMatch = null;
  for (const overlayMsg of overlayMessages) {
    const snippet = semaiBodySnippet(overlayMsg.cleanHtml || overlayMsg.rawHtml || "");
    if (!snippet || snippet.length < 12) continue;
    const idx = recentNormPreviews.findIndex((preview) =>
      preview.includes(snippet) || snippet.includes(preview.slice(0, 30))
    );
    if (idx >= 0) {
      pivotMatch = recent[idx];
      conversationId = pivotMatch.ConversationId || "";
      semaiDebugLine(`REST: matched overlay msg "${snippet.slice(0, 30)}…" → convId=${(conversationId || "").slice(0, 16)}…`);
      break;
    }
  }

  if (!conversationId) {
    semaiDebugLine(`REST: no overlay message matched any of the ${recent.length} recent BodyPreviews`);
    throw new Error("Could not identify the active conversation from recent messages.");
  }

  // ----- Find the latest message in that conversation -----
  const inConversation = recent.filter((m) => m.ConversationId === conversationId && !m.IsDraft);
  if (inConversation.length === 0) {
    // Fallback: include drafts (shouldn't normally happen).
    inConversation.push(pivotMatch);
  }
  inConversation.sort((a, b) => (new Date(b.ReceivedDateTime || 0)) - (new Date(a.ReceivedDateTime || 0)));
  const latest = inConversation[0];
  semaiDebugLine(`REST: latest in convo → ${(latest.Id || "").slice(0, 24)}… received=${latest.ReceivedDateTime || ""}`);
  return latest.Id;
}

function semaiIsVisibleElement(el) {
  if (!(el instanceof Element)) return false;
  const rect = el.getBoundingClientRect();
  const style = window.getComputedStyle(el);
  return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
}

function semaiIsInsideRemouUi(el) {
  return el instanceof Element && !!el.closest("#semai-chat-overlay, #semai-panel, .semai-report-popover");
}

function semaiLooksLikeComposeElement(el) {
  if (!(el instanceof HTMLElement)) return false;
  if (!semaiIsVisibleElement(el)) return false;

  const ariaLabel = (el.getAttribute("aria-label") || "").toLowerCase();
  const role = (el.getAttribute("role") || "").toLowerCase();
  const isEditable = el.isContentEditable || el.getAttribute("contenteditable") === "true";
  const isTextbox = role === "textbox" || el.getAttribute("aria-multiline") === "true";
  const looksLikeMessageBody =
    ariaLabel.includes("message body") ||
    ariaLabel.includes("compose") ||
    ariaLabel.includes("reply");

  if (!isEditable) return false;
  if (looksLikeMessageBody) return true;
  if (isTextbox) return true;

  return !!el.closest('[aria-label*="Message body" i], [data-app-section="MailCompose"]');
}

function semaiGetComposeCandidates() {
  const selector = [
    'div[aria-label="Message body"][contenteditable="true"]',
    'div[aria-label*="Message body" i][contenteditable="true"]',
    'div[role="textbox"][contenteditable="true"]',
    '[aria-label*="compose" i][contenteditable="true"]',
    '[aria-label*="reply" i][contenteditable="true"]',
    '[contenteditable="true"][aria-multiline="true"]',
    '[data-contents="true"] [contenteditable="true"]',
    '[data-lexical-editor="true"][contenteditable="true"]'
  ].join(", ");

  return Array.from(document.querySelectorAll(selector)).filter(semaiLooksLikeComposeElement);
}

function semaiScoreComposeElement(el) {
  if (!(el instanceof HTMLElement)) return -1;

  let score = 0;
  if (el.matches('[data-lexical-editor="true"]')) score += 100;
  if (el.getAttribute("role") === "textbox") score += 25;
  if ((el.getAttribute("aria-label") || "").toLowerCase().includes("message body")) score += 20;
  if (el.closest('[data-app-section="MailCompose"]')) score += 20;
  if (el.closest('[data-contents="true"]')) score += 10;

  // Prefer the innermost actual editor root over larger wrapper elements.
  score += Math.min(el.querySelectorAll('[contenteditable="true"]').length, 5) * -5;

  return score;
}

function semaiPickBestComposeElement(candidates) {
  return [...candidates]
    .sort((left, right) => {
      const scoreDelta = semaiScoreComposeElement(right) - semaiScoreComposeElement(left);
      if (scoreDelta !== 0) return scoreDelta;

      const position = left.compareDocumentPosition(right);
      if (position & Node.DOCUMENT_POSITION_FOLLOWING) return 1;
      if (position & Node.DOCUMENT_POSITION_PRECEDING) return -1;
      return 0;
    })[0] || null;
}

// ===== UTIL: find the compose/body element =====
function getComposeElement() {
  const candidates = semaiGetComposeCandidates();
  if (candidates.length > 0) {
    return semaiPickBestComposeElement(candidates);
  }

  const allEditable = Array.from(document.querySelectorAll('[contenteditable="true"], [role="textbox"]'))
    .filter(semaiLooksLikeComposeElement);
  if (allEditable.length > 0) {
    return semaiPickBestComposeElement(allEditable);
  }

  return null;
}

function semaiWaitForComposeElement(timeoutMs = 12000) {
  return new Promise((resolve, reject) => {
    const startedAt = Date.now();
    let observer = null;
    let intervalId = null;

    const cleanup = () => {
      if (observer) {
        observer.disconnect();
        observer = null;
      }
      if (intervalId) {
        window.clearInterval(intervalId);
        intervalId = null;
      }
    };

    const check = () => {
      const composeEl = getComposeElement();
      if (composeEl) {
        cleanup();
        resolve(composeEl);
        return;
      }

      if (Date.now() - startedAt >= timeoutMs) {
        cleanup();
        reject(new Error("Outlook reply box did not open in time."));
        return;
      }
    };

    check();

    observer = new MutationObserver(check);
    observer.observe(document.body, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ["aria-label", "contenteditable", "role", "style", "class"]
    });

    intervalId = window.setInterval(check, 150);
  });
}

function semaiActivateElement(el) {
  if (!(el instanceof HTMLElement)) return;

  el.focus();

  ["pointerdown", "mousedown", "pointerup", "mouseup", "click"].forEach((eventName) => {
    el.dispatchEvent(new MouseEvent(eventName, {
      bubbles: true,
      cancelable: true,
      view: window
    }));
  });

  if (typeof el.click === "function") {
    el.click();
  }
}

function semaiGetElementActionText(el) {
  if (!(el instanceof Element)) return "";

  return [
    el.getAttribute("aria-label"),
    el.getAttribute("title"),
    el.getAttribute("name"),
    el.getAttribute("data-testid"),
    el.getAttribute("data-icon-name"),
    el.textContent
  ]
    .filter(Boolean)
    .join(" ")
    .replace(/\s+/g, " ")
    .trim();
}

function semaiFindVisibleActionElement(matcher) {
  const candidates = Array.from(document.querySelectorAll(`
    button,
    [role="button"],
    [role="menuitem"],
    [role="option"],
    [tabindex],
    [aria-label],
    [title],
    [data-testid],
    [data-icon-name]
  `))
    .filter(semaiIsVisibleElement)
    .filter((el) => !semaiIsInsideRemouUi(el))
    .filter((el) => matcher(semaiGetElementActionText(el), el));

  return candidates[candidates.length - 1] || null;
}

function semaiFindReplyAllButton() {
  return semaiFindVisibleActionElement((text) => (
    /\breply(?:\s+to)?\s+all\b/i.test(text) ||
    /replyall/i.test(text) ||
    /reply-all/i.test(text)
  ));
}

function semaiFindReplyButton() {
  return semaiFindVisibleActionElement((text) => (
    /\breply\b/i.test(text) &&
    !/\breply(?:\s+to)?\s+all\b/i.test(text) &&
    !/\bforward\b/i.test(text)
  ));
}

function semaiFindReplyAllModeSwitcher() {
  return semaiFindVisibleActionElement((text) => (
    /\breply(?:\s+to)?\s+all\b/i.test(text) ||
    /\brespond\b/i.test(text) ||
    /\bmore reply actions\b/i.test(text)
  ));
}

async function semaiEnsureReplyAllMode(timeoutMs = 2500) {
  const directReplyAll = semaiFindReplyAllButton();
  if (directReplyAll) {
    semaiActivateElement(directReplyAll);
    return true;
  }

  const switcher = semaiFindReplyAllModeSwitcher();
  if (switcher) {
    semaiActivateElement(switcher);
    const startedAt = Date.now();

    while (Date.now() - startedAt < timeoutMs) {
      await new Promise((resolve) => window.setTimeout(resolve, 120));
      const menuReplyAll = semaiFindReplyAllButton();
      if (menuReplyAll) {
        semaiActivateElement(menuReplyAll);
        return true;
      }
    }
  }

  return false;
}

function semaiGetComposeContainer(composeEl) {
  if (!(composeEl instanceof Element)) return null;

  const ancestors = [];
  let current = composeEl;
  while (current instanceof Element) {
    ancestors.push(current);
    current = current.parentElement;
  }

  return ancestors.find((el) => (
    el.matches?.('[data-app-section="MailCompose"]') ||
    el.querySelector?.(
      'button[aria-label="Send"], [role="button"][aria-label="Send"], button[title="Send"], [role="button"][title="Send"]'
    )
  )) || composeEl.parentElement || null;
}

function semaiFindSendButton(scopeEl = document) {
  const selector = [
    'button[aria-label="Send"]',
    '[role="button"][aria-label="Send"]',
    'button[title="Send"]',
    '[role="button"][title="Send"]',
    'button[aria-label*="Send" i]:not([aria-haspopup="menu"])',
    '[role="button"][aria-label*="Send" i]:not([aria-haspopup="menu"])',
    'button[title*="Send" i]:not([aria-haspopup="menu"])',
    '[role="button"][title*="Send" i]:not([aria-haspopup="menu"])',
    '[data-testid*="send" i]',
    '[name*="send" i]'
  ].join(", ");

  const root = scopeEl instanceof Element || scopeEl instanceof Document ? scopeEl : document;
  const matches = Array.from(root.querySelectorAll(selector))
    .filter(semaiIsVisibleElement)
    .filter((el) => !semaiIsInsideRemouUi(el))
    .filter((el) => {
      const label = el.getAttribute("aria-label") || "";
      const title = el.getAttribute("title") || "";
      return !/send to/i.test(label) && !/schedule/i.test(label) && !/schedule/i.test(title);
    });

  if (matches.length > 0) return matches[matches.length - 1];

  const textMatches = Array.from(root.querySelectorAll('button, [role="button"]'))
    .filter(semaiIsVisibleElement)
    .filter((el) => !semaiIsInsideRemouUi(el))
    .filter((el) => /^send$/i.test((el.getAttribute("aria-label") || el.textContent || "").trim()));

  return textMatches[textMatches.length - 1] || null;
}

function semaiTriggerComposeSend(composeEl) {
  composeEl.focus();

  const keyOptions = {
    key: "Enter",
    code: "Enter",
    keyCode: 13,
    which: 13,
    bubbles: true,
    cancelable: true,
    metaKey: true
  };

  composeEl.dispatchEvent(new KeyboardEvent("keydown", keyOptions));
  composeEl.dispatchEvent(new KeyboardEvent("keyup", keyOptions));
}

function semaiDescribeComposeElement(composeEl, index) {
  if (!(composeEl instanceof HTMLElement)) {
    return { index, connected: false };
  }

  const text = (composeEl.innerText || composeEl.textContent || "").replace(/\s+/g, " ").trim();
  return {
    index,
    connected: composeEl.isConnected,
    active: composeEl === getComposeElement(),
    ariaLabel: composeEl.getAttribute("aria-label") || "",
    textLength: text.length,
    rect: {
      width: Math.round(composeEl.getBoundingClientRect().width),
      height: Math.round(composeEl.getBoundingClientRect().height)
    }
  };
}

function semaiLogComposeSnapshot(stage, composeEl = null, extra = {}) {
  try {
    const composeEls = semaiGetComposeCandidates();
    const payload = {
      stage,
      composeCount: composeEls.length,
      composeElements: composeEls.map((el, index) => semaiDescribeComposeElement(el, index)),
      targetCompose: composeEl ? semaiDescribeComposeElement(composeEl, "target") : null,
      ...extra
    };
    semaiNativeLog(`[semai-debug] ${JSON.stringify(payload)}`);
    semaiAppendReplyDebugLine(payload);
  } catch (error) {
    console.log("[semai-debug] Failed to capture compose snapshot", error);
  }
}

function semaiAppendReplyDebugLine(payload) {
  const summary = [
    payload.stage,
    `compose=${payload.composeCount}`,
    payload.targetCompose?.textLength !== undefined ? `targetText=${payload.targetCompose.textLength}` : "",
    payload.sendButtonText ? `send="${payload.sendButtonText}"` : ""
  ]
    .filter(Boolean)
    .join(" | ");

  // Route through the sticky buffer so the line survives chat-overlay rebuilds.
  semaiDebugLine(summary);
}

// Sticky debug log buffer. Survives chat-overlay rebuilds (Outlook re-renders
// the overlay when the thread changes), so the user has time to read/copy
// what scrolled past. Capped to keep memory bounded.
const SEMAI_DEBUG_LOG_MAX = 500;
const semaiDebugLogBuffer = [];

function semaiDebugLine(text) {
  try {
    const line = String(text);
    semaiDebugLogBuffer.push(line);
    if (semaiDebugLogBuffer.length > SEMAI_DEBUG_LOG_MAX) {
      semaiDebugLogBuffer.splice(0, semaiDebugLogBuffer.length - SEMAI_DEBUG_LOG_MAX);
    }
    semaiRenderDebugLog();
  } catch (_) {}
}

function semaiRenderDebugLog() {
  try {
    const debugEl = document.getElementById("semai-chat-reply-debug");
    if (!debugEl) return;
    debugEl.textContent = semaiDebugLogBuffer.join("\n");
    debugEl.scrollTop = debugEl.scrollHeight;
  } catch (_) {}
}

// Replay buffered log lines whenever a new chat overlay appears so we don't
// lose history across Outlook re-renders.
(function semaiAttachDebugLogReplay() {
  if (typeof MutationObserver !== "function") return;
  const observer = new MutationObserver(() => {
    const debugEl = document.getElementById("semai-chat-reply-debug");
    if (debugEl && debugEl.textContent === "" && semaiDebugLogBuffer.length > 0) {
      semaiRenderDebugLog();
    }
  });
  // Wait for body to exist (we run at document_idle, so it should already).
  if (document.body) {
    observer.observe(document.body, { childList: true, subtree: true });
  } else {
    document.addEventListener("DOMContentLoaded", () => {
      observer.observe(document.body, { childList: true, subtree: true });
    }, { once: true });
  }
})();

function semaiGetComposeText(composeEl) {
  if (!(composeEl instanceof HTMLElement)) return "";
  return (composeEl.innerText || composeEl.textContent || "").replace(/\s+/g, " ").trim();
}

function semaiFindComposeCloseButton(scopeEl) {
  if (!(scopeEl instanceof Element) && !(scopeEl instanceof Document)) return null;

  const candidates = Array.from(scopeEl.querySelectorAll(`
    button,
    [role="button"],
    [role="menuitem"],
    [aria-label],
    [title],
    [data-testid],
    [data-icon-name]
  `))
    .filter(semaiIsVisibleElement)
    .filter((el) => !semaiIsInsideRemouUi(el));

  const discardButton = candidates.find((el) => /\bdiscard\b/i.test(semaiGetElementActionText(el)));
  if (discardButton) return discardButton;

  return candidates.find((el) => {
    const text = semaiGetElementActionText(el);
    return /\bclose\b/i.test(text) && /\bdraft\b/i.test(text);
  }) || null;
}

function semaiFindScopedActionElement(scopeEl, matcher) {
  if (!(scopeEl instanceof Element) && !(scopeEl instanceof Document)) return null;

  const candidates = Array.from(scopeEl.querySelectorAll(`
    button,
    [role="button"],
    [role="menuitem"],
    [aria-label],
    [title],
    [data-testid],
    [data-icon-name]
  `))
    .filter(semaiIsVisibleElement)
    .filter((el) => !semaiIsInsideRemouUi(el))
    .filter((el) => matcher(semaiGetElementActionText(el), el));

  return candidates[candidates.length - 1] || null;
}

function semaiDispatchEscape(el) {
  if (!(el instanceof HTMLElement)) return;
  el.focus();

  const eventOptions = {
    key: "Escape",
    code: "Escape",
    keyCode: 27,
    which: 27,
    bubbles: true,
    cancelable: true
  };

  el.dispatchEvent(new KeyboardEvent("keydown", eventOptions));
  el.dispatchEvent(new KeyboardEvent("keyup", eventOptions));
}

async function semaiDismissLeftoverEmptyCompose() {
  await new Promise((resolve) => window.setTimeout(resolve, 220));

  const composeEl = getComposeElement();
  if (!composeEl) return false;

  const composeText = semaiGetComposeText(composeEl);
  if (composeText.length > 0) {
    semaiLogComposeSnapshot("leftover_compose_not_empty", composeEl, { composeTextLength: composeText.length });
    return false;
  }

  semaiLogComposeSnapshot("leftover_empty_compose_detected", composeEl);
  const composeContainer = semaiGetComposeContainer(composeEl);
  semaiDispatchEscape(composeEl);
  await new Promise((resolve) => window.setTimeout(resolve, 220));

  const discardButton = semaiFindComposeCloseButton(document);
  if (discardButton) {
    semaiActivateElement(discardButton);
    await new Promise((resolve) => window.setTimeout(resolve, 220));
    semaiLogComposeSnapshot("leftover_empty_compose_discarded", composeEl, {
      buttonText: semaiGetElementActionText(discardButton)
    });
    return true;
  }

  const closeButton = semaiFindComposeCloseButton(composeContainer || document);
  if (closeButton) {
    semaiActivateElement(closeButton);
    await new Promise((resolve) => window.setTimeout(resolve, 220));
    semaiLogComposeSnapshot("leftover_empty_compose_closed_button", composeEl, {
      buttonText: semaiGetElementActionText(closeButton)
    });
    return true;
  }

  semaiLogComposeSnapshot("leftover_empty_compose_close_failed", composeEl);
  return false;
}

async function semaiDismissLeftoverEmptyComposeTwice() {
  const firstPass = await semaiDismissLeftoverEmptyCompose();
  await new Promise((resolve) => window.setTimeout(resolve, 700));
  const secondPass = await semaiDismissLeftoverEmptyCompose();
  return firstPass || secondPass;
}

function semaiDescribeThreadDraftRows() {
  const bodyNodes = Array.from(document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])'));

  return bodyNodes.slice(-8).map((bodyEl, index) => {
    const container = bodyEl.closest('[data-test-id="mailMessageBodyContainer"]')?.parentElement || bodyEl.parentElement;
    const scope = container || bodyEl;
    const scopeText = semaiFullText(scope).replace(/\s+/g, " ").trim();
    const actionCandidates = Array.from(scope.querySelectorAll(`
      button,
      [role="button"],
      [role="menuitem"],
      [aria-label],
      [title],
      [data-testid],
      [data-icon-name]
    `))
      .filter(semaiIsVisibleElement)
      .filter((el) => !semaiIsInsideRemouUi(el))
      .map((el) => semaiGetElementActionText(el))
      .filter(Boolean)
      .slice(0, 8);

    return {
      index,
      bodyTextLength: semaiFullText(bodyEl).replace(/\s+/g, " ").trim().length,
      scopeTextSample: scopeText.slice(0, 160),
      looksLikeDraft: /\bdraft\b/i.test(scopeText),
      actions: actionCandidates
    };
  });
}

function semaiLogThreadDraftSnapshot(stage) {
  try {
    const rows = semaiDescribeThreadDraftRows();
    semaiAppendReplyDebugLine({
      stage,
      composeCount: semaiGetComposeCandidates().length,
      targetCompose: null,
      rowSummary: rows.map((row) => `${row.index}:${row.looksLikeDraft ? "draft" : "msg"}:${row.bodyTextLength}`).join(", ")
    });
    semaiNativeLog(`[semai-thread-debug] ${JSON.stringify({ stage, rows })}`);
  } catch (error) {
    console.log("[semai-thread-debug] Failed to capture thread rows", error);
  }
}

function semaiComposeIsStillActive(composeEl) {
  return composeEl instanceof HTMLElement && composeEl.isConnected && semaiLooksLikeComposeElement(composeEl);
}

async function semaiWaitForComposeToClose(composeEl, timeoutMs = 5000) {
  const startedAt = Date.now();

  while (Date.now() - startedAt < timeoutMs) {
    if (!semaiComposeIsStillActive(composeEl)) {
      return true;
    }
    await new Promise((resolve) => window.setTimeout(resolve, 120));
  }

  return !semaiComposeIsStillActive(composeEl);
}

async function semaiSendCompose(composeEl) {
  const composeContainer = semaiGetComposeContainer(composeEl);
  const sendButton = semaiFindSendButton(composeContainer || composeEl.ownerDocument || document);
  semaiLogComposeSnapshot("before_send", composeEl, {
    sendButtonText: sendButton ? semaiGetElementActionText(sendButton) : ""
  });

  semaiTriggerComposeSend(composeEl);
  if (await semaiWaitForComposeToClose(composeEl, 4000)) {
    semaiLogComposeSnapshot("keyboard_send_closed", composeEl);
    return;
  }

  if (sendButton) {
    semaiActivateElement(sendButton);
    if (await semaiWaitForComposeToClose(composeEl, 4000)) {
      semaiLogComposeSnapshot("send_button_closed", composeEl);
      return;
    }
  }

  semaiLogComposeSnapshot("send_failed_still_open", composeEl);
  throw new Error("Reply all draft opened, but Outlook did not send it.");
}

async function semaiOpenReplyAllCompose() {
  let composeEl = getComposeElement();
  if (composeEl) return composeEl;

  const openedReplyAll = await semaiEnsureReplyAllMode();
  if (openedReplyAll) {
    composeEl = await semaiWaitForComposeElement();
    semaiLogComposeSnapshot("reply_all_opened_direct", composeEl);
    return composeEl;
  }

  const replyBtn = semaiFindReplyButton();
  if (!replyBtn) {
    throw new Error("Reply controls not found in Outlook.");
  }

  semaiActivateElement(replyBtn);
  composeEl = await semaiWaitForComposeElement();
  semaiLogComposeSnapshot("reply_opened_fallback", composeEl);

  // Some Outlook thread states only expose a generic Reply action until the compose UI opens.
  return composeEl;
}

async function semaiInsertComposeText(composeEl, text) {
  // Outlook's React editor initialises its internal state asynchronously after the
  // contenteditable element appears in the DOM. If we write content before that
  // completes, React resets the compose area and wipes our text. A short pause here
  // lets the framework settle before we touch the editor.
  await new Promise(resolve => window.setTimeout(resolve, 300));

  composeEl.focus();
  semaiLogComposeSnapshot("before_insert", composeEl, { draftLength: text.length });

  const selection = window.getSelection();
  if (selection) {
    const range = document.createRange();
    range.selectNodeContents(composeEl);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);
  }

  // Keep insertion scoped to the compose element instead of relying on a
  // document-wide selectAll, which can target the wrong Outlook surface.
  const inserted = document.execCommand("insertText", false, text);

  if (!inserted) {
    // Fallback: direct DOM manipulation when execCommand is unavailable.
    const lines = text.split(/\n/);
    const fragment = document.createDocumentFragment();
    lines.forEach((line) => {
      const block = document.createElement("div");
      if (line) {
        block.textContent = line;
      } else {
        block.appendChild(document.createElement("br"));
      }
      fragment.appendChild(block);
    });
    composeEl.innerHTML = "";
    composeEl.appendChild(fragment);
    composeEl.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: text }));
  }

  // Outlook can render the inserted text before its internal draft model is ready.
  // Give the compose pipeline a brief moment to commit the body before we try to send.
  await new Promise((resolve) => requestAnimationFrame(() => requestAnimationFrame(resolve)));
  await new Promise((resolve) => window.setTimeout(resolve, 180));
  semaiLogComposeSnapshot("after_insert", composeEl, { draftLength: text.length });
}

async function semaiDraftReplyAllFromChat() {
  const input = document.getElementById("semai-chat-reply-input");
  const draftBtn = document.getElementById("semai-chat-reply-draft-btn");
  const sendBtn = document.getElementById("semai-chat-reply-send-btn");
  const status = document.getElementById("semai-chat-reply-status");
  const draft = input?.value.trim() || "";

  if (!draft) {
    if (status) status.textContent = "Type a reply first.";
    return;
  }

  if (draftBtn) draftBtn.disabled = true;
  if (sendBtn) sendBtn.disabled = true;
  if (status) status.textContent = "Opening Reply all draft…";

  try {
    const composeEl = await semaiOpenReplyAllCompose();
    await semaiInsertComposeText(composeEl, draft);

    if (status) status.textContent = "Reply all draft inserted into Outlook.";
  } catch (err) {
    if (status) status.textContent = err.message;
  } finally {
    if (draftBtn) draftBtn.disabled = false;
    if (sendBtn) sendBtn.disabled = false;
  }
}

// Try the REST API path first. Returns true on success, false if we should
// fall back to the compose-UI path. Never throws — any error is logged and
// the caller proceeds with the fallback.
async function semaiTryReplyAllViaRestApi(draft) {
  if (!SEMAI_USE_REST_API_REPLY) {
    semaiDebugLine("REST: flag off → compose-UI fallback");
    semaiNativeLog("[semai-rest] feature flag off; using compose-UI fallback");
    return false;
  }
  if (!semaiCachedOutlookToken) {
    semaiDebugLine("REST: ✗ no token yet → compose-UI fallback");
    semaiNativeLog("[semai-rest] no token captured yet; using compose-UI fallback");
    return false;
  }

  semaiDebugLine(`REST: token OK (len=${semaiCachedOutlookToken.length})`);

  // Pick the last (most-recent) message in the active thread — that's what
  // "Reply all" targets in Outlook's UI.
  const overlay = document.getElementById("semai-chat-overlay");
  const messages = overlay?._semaiMessages || [];
  const lastMessage = messages[messages.length - 1];
  const bodyEl = lastMessage?.sourceBodyEl
    || document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])').item(
         document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])').length - 1
       );

  if (!(bodyEl instanceof Element)) {
    semaiDebugLine("REST: ✗ no body element → compose-UI fallback");
    semaiNativeLog("[semai-rest] no message body element found; using compose-UI fallback");
    return false;
  }

  // Step 1: try DOM extraction. Older Outlook builds stamped IDs into
  // attributes; modern builds keep them in React state, so this often misses.
  let messageId = semaiExtractMessageId(bodyEl);

  // Step 2: if DOM had nothing, ask the API. Pass the full thread messages
  // so the resolver can match against any of them — the OLDER messages are
  // most likely to already be indexed and least likely to collide with the
  // user's just-typed draft.
  if (!messageId) {
    try {
      semaiDebugLine(`REST: resolving msgId via API (using ${messages.length} thread messages)…`);
      messageId = await semaiResolveMessageIdViaRest(messages);
    } catch (err) {
      const msg = err && err.message ? err.message : String(err);
      semaiDebugLine(`REST: ✗ resolver failed: ${msg.slice(0, 100)}`);
      semaiNativeLog(`[semai-rest] resolver failed: ${msg}; using compose-UI fallback`);
      return false;
    }
  }

  if (!messageId) {
    semaiDebugLine("REST: ✗ no message ID after DOM + API resolve → fallback");
    semaiNativeLog("[semai-rest] could not extract or resolve message ID; using compose-UI fallback");
    return false;
  }

  semaiDebugLine(`REST: msgId=${messageId.slice(0, 20)}…`);

  try {
    semaiDebugLine("REST: POSTing replyAll…");
    semaiNativeLog(`[semai-rest] POSTing replyAll for messageId="${messageId.slice(0, 24)}…" draftLen=${draft.length}`);
    await semaiPostOutlookReplyAll(messageId, draft);
    semaiDebugLine("REST: ✓ sent via API — no draft created");
    semaiNativeLog("[semai-rest] replyAll succeeded via REST API (no draft created)");
    return true;
  } catch (err) {
    const msg = err && err.message ? err.message : String(err);
    semaiDebugLine(`REST: ✗ replyAll failed: ${msg.slice(0, 100)}`);
    semaiNativeLog(`[semai-rest] replyAll REST call failed: ${msg}; using compose-UI fallback`);
    return false;
  }
}

async function semaiSendReplyAllFromChat() {
  const input = document.getElementById("semai-chat-reply-input");
  const draftBtn = document.getElementById("semai-chat-reply-draft-btn");
  const sendBtn = document.getElementById("semai-chat-reply-send-btn");
  const status = document.getElementById("semai-chat-reply-status");
  const draft = input?.value.trim() || "";

  if (!draft) {
    if (status) status.textContent = "Type a reply first.";
    return;
  }

  if (draftBtn) draftBtn.disabled = true;
  if (sendBtn) sendBtn.disabled = true;
  if (status) status.textContent = "Sending Reply all…";

  try {
    // Attempt REST API path first — sends immediately without ever opening the
    // compose UI, so Outlook's autosave never creates an empty draft. If
    // anything goes wrong we silently fall back to the compose-UI flow.
    const restOk = await semaiTryReplyAllViaRestApi(draft);

    if (!restOk) {
      const composeEl = await semaiOpenReplyAllCompose();
      await semaiInsertComposeText(composeEl, draft);
      semaiLogComposeSnapshot("before_chat_send", composeEl, { draftLength: draft.length });
      await new Promise(resolve => window.setTimeout(resolve, 100));
      await semaiSendCompose(composeEl);
      // NOTE: previously we called semaiDismissLeftoverEmptyComposeTwice() here
      // to sweep any leftover empty compose panel. That cleanup was the source
      // of the "Delete this draft?" confirmation prompt — Outlook gates its
      // Discard button behind a confirm dialog. After Outlook's native Send,
      // the compose UI should tear itself down on its own, exactly as it does
      // for a manual click. We now trust that native flow and do nothing
      // extra; if anything Outlook-side truly needs to be closed, that is
      // Outlook's business to handle.
    }

    if (status) status.textContent = "Reply all sent.";
    if (input) input.value = "";
  } catch (err) {
    if (status) status.textContent = err.message;
  } finally {
    if (draftBtn) draftBtn.disabled = false;
    if (sendBtn) sendBtn.disabled = false;
  }
}

function semaiOpenOnboardingAppWindow() {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage({ type: "OPEN_ONBOARDING_APP" }, (response) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }

      if (!response?.ok) {
        reject(new Error(response?.error || "Could not open the Remou app."));
        return;
      }

      resolve();
    });
  });
}

function semaiTrackEvent(eventName, details = {}) {
  chrome.runtime.sendMessage({ type: "TRACK_EVENT", eventName, details }, () => {
    if (chrome.runtime.lastError) {
      console.error("[semai] Failed to track event", eventName, chrome.runtime.lastError.message);
    }
  });
}

const SEMAI_DEBUG = false;
const SEMAI_AI_AGENT_ENABLED = false;
const SEMAI_CALIBRATION_STORAGE_KEY = "semaiSenderCalibration";
const SEMAI_PANEL_POSITION_STORAGE_KEY = "semaiPanelPosition";
const SEMAI_CHAT_UI_SETTINGS_STORAGE_KEY = "semaiChatUiSettings";
const SEMAI_CHAT_FONT_SIZE_MIN = 11;
const SEMAI_CHAT_FONT_SIZE_MAX = 20;
const SEMAI_CHAT_FONT_SIZE_DEFAULT = 13;

let semaiSavedSelection = null;
let semaiCalibrationState = null;
let semaiCalibrationHoverEl = null;
let semaiAutoOpenSuppressedSignature = "";
let semaiPanelDragState = null;

// Sends a log message to the native host so it appears in the Xcode console.
// Falls back to console.log if the background relay is unavailable.
function semaiNativeLog(text) {
  try {
    browser.runtime.sendMessage({ type: "semaiLog", text });
  } catch (e) {
    console.log("[semai-sig]", text);
  }
}

function semaiLog(message, details) {
  if (!SEMAI_DEBUG) {
    return;
  }

  if (details === undefined) {
    console.log(message);
    return;
  }

  console.log(message, details);
}

function semaiSelectionIsInsideCompose(selection, bodyEl) {
  if (!selection || selection.rangeCount === 0 || !bodyEl) {
    return false;
  }

  const range = selection.getRangeAt(0);
  return bodyEl.contains(range.commonAncestorContainer) && selection.toString().trim().length > 0;
}

function semaiSaveSelectionFromCompose() {
  const bodyEl = getComposeElement();
  const selection = window.getSelection();

  if (!bodyEl || !semaiSelectionIsInsideCompose(selection, bodyEl)) {
    return;
  }

  semaiSavedSelection = {
    range: selection.getRangeAt(0).cloneRange(),
    text: selection.toString()
  };

  semaiLog("[semai] Saved compose selection", { text: semaiSavedSelection.text });
}

function semaiGetSelectionForRewrite(bodyEl) {
  const selection = window.getSelection();

  if (semaiSelectionIsInsideCompose(selection, bodyEl)) {
    return {
      range: selection.getRangeAt(0).cloneRange(),
      text: selection.toString()
    };
  }

  if (semaiSavedSelection && bodyEl.contains(semaiSavedSelection.range.commonAncestorContainer)) {
    return {
      range: semaiSavedSelection.range.cloneRange(),
      text: semaiSavedSelection.text
    };
  }

  return null;
}

// ===== OpenAI API call =====
async function semaiCallOpenAI(text, mode, customInstruction) {
  if (!SEMAI_AI_AGENT_ENABLED) {
    throw new Error("AI rewrite is disabled for now.");
  }

  if (!SEMAI_OPENAI_API_KEY) {
    throw new Error("No API key — paste your key into secrets.js first.");
  }

  const preset = SEMAI_PRESETS[mode] || SEMAI_PRESETS.custom;
  const userMessage = preset.userTemplate
    .replace("{{TEXT}}", text)
    .replace("{{INSTRUCTION}}", customInstruction || "");

  const resp = await fetch("https://api.openai.com/v1/chat/completions", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${SEMAI_OPENAI_API_KEY}`
    },
    body: JSON.stringify({
      model: SEMAI_MODEL,
      messages: [
        { role: "system", content: preset.system },
        { role: "user", content: userMessage }
      ],
      temperature: 0.7,
      max_tokens: 600
    })
  });

  if (!resp.ok) {
    const err = await resp.json().catch(() => ({}));
    throw new Error(err.error?.message || `OpenAI API error ${resp.status}`);
  }

  const data = await resp.json();
  // Strip surrounding quotes the model sometimes adds
  return (data.choices?.[0]?.message?.content ?? text)
    .trim()
    .replace(/^["""’’`]|["""’’`]$/g, "");
}

// ===== Loading state helpers =====
function semaiSetLoading(loading) {
  const panel = document.getElementById("semai-panel");
  if (!panel) return;
  panel.querySelectorAll(".semai-chip, .semai-apply-btn").forEach(btn => {
    btn.disabled = loading;
  });
  const applyBtn = panel.querySelector(".semai-apply-btn");
  if (applyBtn) {
    applyBtn.textContent = loading ? "Working…" : "Apply custom instruction";
  }
  if (loading) {
    panel.classList.add("semai-working");
  } else {
    panel.classList.remove("semai-working");
  }
}

// ===== Replace selected text within the compose element =====
async function rewriteSelectionInCompose(mode) {
  const bodyEl = getComposeElement();
  if (!bodyEl) {
    alert("semai couldn’t find the email body. Open a compose or reply window.");
    return;
  }

  const selectionState = semaiGetSelectionForRewrite(bodyEl);
  if (!selectionState) {
    alert("Highlight the text you want semai to help with, then try again.");
    return;
  }

  const { range, text: selectedText } = selectionState;
  if (!selectedText.trim()) {
    alert("No words selected. Highlight the part you want semai to rewrite.");
    return;
  }

  const customInput = document.getElementById("semai-custom-input");
  const customInstruction = customInput?.value.trim() ?? "";
  if (mode === "custom" && !customInstruction) {
    alert("Enter a custom instruction before applying it.");
    return;
  }

  semaiLog("[semai] Rewriting selection", { mode, selectedText, customInstruction });

  semaiSetLoading(true);
  try {
    const newText = await semaiCallOpenAI(selectedText, mode, customInstruction);
    const selection = window.getSelection();

    // Restore the saved range (user may have clicked elsewhere during await)
    selection.removeAllRanges();
    selection.addRange(range);

    range.deleteContents();
    const textNode = document.createTextNode(newText);
    range.insertNode(textNode);

    // Move caret after inserted text
    selection.removeAllRanges();
    const newRange = document.createRange();
    newRange.setStartAfter(textNode);
    newRange.collapse(true);
    selection.addRange(newRange);
    semaiSavedSelection = null;

    semaiLog("[semai] Rewrite complete");
  } catch (err) {
    alert(`semai: ${err.message}`);
    console.error("[semai] Rewrite error", err);
  } finally {
    semaiSetLoading(false);
  }
}

// ===== Toggle collapsed state =====
function toggleSemaiPanel() {
  const panel = document.getElementById("semai-panel");
  if (!panel) return;

  const toggleBtn = panel.querySelector(".semai-toggle-btn");
  const isCollapsed = panel.classList.toggle("semai-collapsed");

  if (toggleBtn) {
    toggleBtn.textContent = isCollapsed ? "▾" : "▴";
    toggleBtn.setAttribute(
      "aria-label",
      isCollapsed ? "Expand semai" : "Collapse semai"
    );
  }

  window.requestAnimationFrame(() => {
    semaiEnsurePanelVisible(panel);
  });
  semaiLog("[semai] Panel toggled", { collapsed: isCollapsed });
}

function semaiGetSavedPanelPosition() {
  try {
    const raw = window.localStorage.getItem(SEMAI_PANEL_POSITION_STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function semaiBuildPanelPositionSnapshot(panel) {
  const rect = panel.getBoundingClientRect();
  const anchorX = rect.left + rect.width / 2 >= window.innerWidth / 2 ? "right" : "left";
  const anchorY = rect.top + rect.height / 2 >= window.innerHeight / 2 ? "bottom" : "top";

  return {
    left: Math.round(rect.left),
    top: Math.round(rect.top),
    right: Math.round(window.innerWidth - rect.right),
    bottom: Math.round(window.innerHeight - rect.bottom),
    anchorX,
    anchorY
  };
}

function semaiSavePanelPosition(panel) {
  try {
    window.localStorage.setItem(
      SEMAI_PANEL_POSITION_STORAGE_KEY,
      JSON.stringify(semaiBuildPanelPositionSnapshot(panel))
    );
  } catch (e) {
    console.warn("[semai] Failed to save panel position", e);
  }
}

function semaiGetCurrentPanelPosition(panel) {
  const rect = panel.getBoundingClientRect();
  const saved = semaiGetSavedPanelPosition();
  return {
    left: Math.round(rect.left),
    top: Math.round(rect.top),
    right: Math.round(window.innerWidth - rect.right),
    bottom: Math.round(window.innerHeight - rect.bottom),
    anchorX: saved?.anchorX === "right" ? "right" : "left",
    anchorY: saved?.anchorY === "bottom" ? "bottom" : "top"
  };
}

function semaiApplyPanelPosition(panel, position) {
  const width = panel.offsetWidth;
  const height = panel.offsetHeight;
  const maxLeft = Math.max(8, window.innerWidth - width - 8);
  const maxTop = Math.max(8, window.innerHeight - height - 8);

  const anchorX = position?.anchorX === "right" ? "right" : "left";
  const anchorY = position?.anchorY === "bottom" ? "bottom" : "top";

  const rawLeft = anchorX === "right"
    ? window.innerWidth - width - (typeof position?.right === "number" ? position.right : 16)
    : (typeof position?.left === "number" ? position.left : 16);
  const rawTop = anchorY === "bottom"
    ? window.innerHeight - height - (typeof position?.bottom === "number" ? position.bottom : 16)
    : (typeof position?.top === "number" ? position.top : 16);

  const nextLeft = Math.min(Math.max(8, rawLeft), maxLeft);
  const nextTop = Math.min(Math.max(8, rawTop), maxTop);
  const nextRight = Math.max(8, window.innerWidth - nextLeft - width);
  const nextBottom = Math.max(8, window.innerHeight - nextTop - height);

  panel.style.left = anchorX === "left" ? `${nextLeft}px` : "auto";
  panel.style.right = anchorX === "right" ? `${nextRight}px` : "auto";
  panel.style.top = anchorY === "top" ? `${nextTop}px` : "auto";
  panel.style.bottom = anchorY === "bottom" ? `${nextBottom}px` : "auto";
}

function semaiEnsurePanelVisible(panel, persist = true) {
  if (!(panel instanceof HTMLElement)) return;

  semaiApplyPanelPosition(panel, semaiGetCurrentPanelPosition(panel));

  if (persist) {
    semaiSavePanelPosition(panel);
  }
}

function semaiRestorePanelPosition(panel) {
  const saved = semaiGetSavedPanelPosition();
  if (!saved) return;
  semaiApplyPanelPosition(panel, saved);
}

function semaiHandlePanelDragMove(event) {
  if (!semaiPanelDragState) return;

  const nextLeft = event.clientX - semaiPanelDragState.offsetX;
  const nextTop = event.clientY - semaiPanelDragState.offsetY;
  semaiApplyPanelPosition(semaiPanelDragState.panel, {
    left: nextLeft,
    top: nextTop,
    anchorX: "left",
    anchorY: "top"
  });
}

function semaiHandlePanelDragEnd() {
  if (!semaiPanelDragState) return;

  const panel = semaiPanelDragState.panel;
  const droppedPosition = semaiBuildPanelPositionSnapshot(panel);
  semaiApplyPanelPosition(panel, droppedPosition);
  semaiSavePanelPosition(panel);
  panel.classList.remove("semai-dragging");
  semaiPanelDragState = null;
  document.removeEventListener("pointermove", semaiHandlePanelDragMove);
  document.removeEventListener("pointerup", semaiHandlePanelDragEnd);
}

function semaiEnablePanelDragging(panel) {
  const handle = panel.querySelector(".semai-header");
  if (!handle) return;

  handle.addEventListener("pointerdown", (event) => {
    const target = event.target;
    if (target instanceof Element && target.closest("button, input, textarea")) return;

    const rect = panel.getBoundingClientRect();
    semaiPanelDragState = {
      panel,
      offsetX: event.clientX - rect.left,
      offsetY: event.clientY - rect.top
    };

    panel.classList.add("semai-dragging");
    semaiApplyPanelPosition(panel, {
      left: rect.left,
      top: rect.top,
      anchorX: "left",
      anchorY: "top"
    });

    document.addEventListener("pointermove", semaiHandlePanelDragMove);
    document.addEventListener("pointerup", semaiHandlePanelDragEnd);
  });
}

// ===== UI: create floating panel =====
function createPanel() {
  if (document.getElementById("semai-panel")) return;

  const panel = document.createElement("div");
  panel.id = "semai-panel";

  panel.innerHTML = `
    <div class="semai-header">
      <div class="semai-header-left">
        <div class="semai-logo-dot"></div>
        <div class="semai-title">REMOU</div>
      </div>
      <div class="semai-header-actions">
        <button
          class="semai-settings-btn"
          type="button"
          aria-label="Open Remou setup"
          title="Open Remou setup"
        >
          ⚙
        </button>
        <button
          class="semai-toggle-btn"
          type="button"
          aria-label="Collapse REMOU"
        >
          ▴
        </button>
      </div>
    </div>
    <div class="semai-body">
      <button class="semai-chat-toggle-btn" type="button" style="display:none">Turn on chat view</button>
      <button class="semai-calibrate-btn" type="button">Train sender detection</button>
      <p id="semai-calibration-status" class="semai-calibration-status"></p>
      ${SEMAI_AI_AGENT_ENABLED ? `
      <p class="semai-subtitle">
        Highlight text in your email and choose how you want it to sound.
      </p>
      <div class="semai-chip-row">
        <button class="semai-chip" data-mode="polite">
          Polite
          <span>softer tone</span>
        </button>
        <button class="semai-chip" data-mode="concise">
          Concise
          <span>short & clear</span>
        </button>
      </div>
      <input
        type="text"
        id="semai-custom-input"
        placeholder="Custom instruction (e.g. “more formal, but friendly”)"
      />
      <div class="semai-footer">
        <button class="semai-apply-btn" data-mode="custom">
          Apply custom instruction
        </button>
      </div>
      ` : ""}
    </div>
  `;

  panel.addEventListener("click", (e) => {
    const target = e.target;
    if (!(target instanceof HTMLButtonElement)) return;

    if (target.classList.contains("semai-settings-btn")) {
      semaiOpenOnboardingAppWindow().catch((error) => {
        console.error("[semai] Failed to open onboarding app", error);
        semaiShowOnboardingModal();
      });
      return;
    }

    // Handle collapse/expand toggle
    if (target.classList.contains("semai-toggle-btn")) {
      toggleSemaiPanel();
      return;
    }

    // Handle chat view toggle
    if (target.classList.contains("semai-chat-toggle-btn")) {
      if (semaiChatViewActive) {
        semaiDeactivateChatView();
      } else {
        semaiChatViewPinned = true;
        semaiActivateChatView();
      }
      return;
    }

    if (target.classList.contains("semai-calibrate-btn")) {
      semaiShowOnboardingModal();
      return;
    }

    const mode = target.dataset.mode;
    if (!mode) return;

    if (!SEMAI_AI_AGENT_ENABLED) return;

    console.log("[semai] Button clicked", { mode });

    // Polite / Concise chips → instant rewrite
    if (target.classList.contains("semai-chip")) {
      rewriteSelectionInCompose(mode);
      return;
    }

    // Custom → use text from input, but same rewriteSelectionInCompose path
    if (target.classList.contains("semai-apply-btn")) {
      rewriteSelectionInCompose("custom");
    }
  });

  document.body.appendChild(panel);
  semaiRestorePanelPosition(panel);
  semaiEnablePanelDragging(panel);
  window.requestAnimationFrame(() => {
    semaiEnsurePanelVisible(panel, false);
  });
  const calibration = semaiGetCalibration();
  const calibrateBtn = panel.querySelector(".semai-calibrate-btn");
  if (calibrateBtn) {
    calibrateBtn.textContent = calibration
      ? "Retrain sender detection"
      : "Set up Remou";
  }
  semaiUpdateCalibrationStatus(
    calibration
      ? "✓ Setup complete. You can retrain anytime."
      : "👆 Start here — tell Remou who you are.",
    calibration ? "success" : "neutral"
  );

  // Show first-run onboarding modal if never calibrated
  if (!calibration) {
    semaiShowOnboardingModal();
  }

  semaiLog("[semai] Panel created");
}

// ===== SIGNATURE STRIPPING (reading view) =====
// All signature detection and body cleaning logic is in semaiSigDetector.js.

// ===== CHAT VIEW ============================================================

let semaiChatViewActive = false;
let semaiChatViewActivationInProgress = false;
let semaiChatViewPinned = false;
let semaiCurrentUser = null; // { name, email, initials }
let semaiReportHoverRow = null;
let semaiReportModeOverlay = null;
let semaiReportPopoverEl = null;
let semaiReportMissedBodies = [];
let semaiReportHoverOriginalBody = null;
let semaiChatUiSettings = semaiGetChatUiSettings();

function semaiNormalizeChatUiSettings(settings) {
  const rawFontSize = Number(settings?.fontSize);
  const fontSize = Number.isFinite(rawFontSize)
    ? Math.min(SEMAI_CHAT_FONT_SIZE_MAX, Math.max(SEMAI_CHAT_FONT_SIZE_MIN, Math.round(rawFontSize)))
    : SEMAI_CHAT_FONT_SIZE_DEFAULT;

  return { fontSize };
}

function semaiGetChatUiSettings() {
  try {
    const raw = window.localStorage.getItem(SEMAI_CHAT_UI_SETTINGS_STORAGE_KEY);
    return semaiNormalizeChatUiSettings(raw ? JSON.parse(raw) : null);
  } catch (_) {
    return semaiNormalizeChatUiSettings(null);
  }
}

function semaiPersistChatUiSettings(settings) {
  try {
    window.localStorage.setItem(SEMAI_CHAT_UI_SETTINGS_STORAGE_KEY, JSON.stringify(settings));
  } catch (error) {
    console.warn("[semai] Failed to persist chat UI settings", error);
  }
}

function semaiApplyChatUiSettings(overlay) {
  if (!(overlay instanceof HTMLElement)) return;
  const settings = semaiNormalizeChatUiSettings(semaiChatUiSettings);
  overlay.style.setProperty("--semai-chat-font-size", `${settings.fontSize}px`);

  const slider = overlay.querySelector(".semai-chat-settings-slider");
  const value = overlay.querySelector(".semai-chat-settings-value");
  if (slider instanceof HTMLInputElement) {
    slider.value = String(settings.fontSize);
  }
  if (value) {
    value.textContent = `${settings.fontSize}px`;
  }
}

function semaiUpdateChatUiSettings(nextPartialSettings) {
  semaiChatUiSettings = semaiNormalizeChatUiSettings({
    ...semaiChatUiSettings,
    ...nextPartialSettings
  });
  semaiPersistChatUiSettings(semaiChatUiSettings);

  const overlay = document.getElementById("semai-chat-overlay");
  if (overlay) {
    semaiApplyChatUiSettings(overlay);
  }
}

function semaiCloseChatSettingsPopover(overlay) {
  if (!(overlay instanceof HTMLElement)) return;
  const popover = overlay.querySelector(".semai-chat-settings-popover");
  const toggleBtn = overlay.querySelector(".semai-chat-settings-toggle");
  if (!popover || !toggleBtn) return;

  popover.hidden = true;
  toggleBtn.setAttribute("aria-expanded", "false");
}

function semaiOpenChatSettingsPopover(overlay) {
  if (!(overlay instanceof HTMLElement)) return;
  const popover = overlay.querySelector(".semai-chat-settings-popover");
  const toggleBtn = overlay.querySelector(".semai-chat-settings-toggle");
  if (!popover || !toggleBtn) return;

  semaiApplyChatUiSettings(overlay);
  popover.hidden = false;
  toggleBtn.setAttribute("aria-expanded", "true");
}

function semaiToggleChatSettingsPopover(overlay) {
  if (!(overlay instanceof HTMLElement)) return;
  const popover = overlay.querySelector(".semai-chat-settings-popover");
  if (!popover) return;

  if (popover.hidden) {
    semaiOpenChatSettingsPopover(overlay);
  } else {
    semaiCloseChatSettingsPopover(overlay);
  }
}

// Deterministic avatar colour from name — 8-colour palette
const SEMAI_AVATAR_COLORS = [
  "#6366f1","#8b5cf6","#ec4899","#f59e0b","#10b981","#3b82f6","#ef4444","#14b8a6"
];
function semaiNameColor(name) {
  let h = 0;
  for (let i = 0; i < name.length; i++) h = name.charCodeAt(i) + ((h << 5) - h);
  return SEMAI_AVATAR_COLORS[Math.abs(h) % SEMAI_AVATAR_COLORS.length];
}

function semaiFirstNameFromDisplayName(displayName) {
  const name = (displayName || "")
    .replace(/^from:\s*/i, "")
    .replace(/<[^>]+>/g, "")
    .trim();

  if (!name) return "";

  if (name.includes(",")) {
    const afterComma = (name.split(/\s*,\s*/, 2)[1] || "")
      .split(/\s+/)
      .find((token) => token.length >= 2 && /^\p{L}/u.test(token));
    if (afterComma) {
      return afterComma;
    }
  }

  return (name.split(/[\s,<(@]+/)[0] || "").trim();
}

function semaiInitials(name) {
  const cleaned = (name || "")
    .replace(/^from:\s*/i, "")
    .replace(/<[^>]+>/g, "")
    .trim();

  if (!cleaned) return "?";

  const splitPart = (text) => text
    .split(/\s+/)
    .map((token) => token.replace(/^[^A-Za-zÀ-ÿ]+|[^A-Za-zÀ-ÿ]+$/g, ""))
    .filter((token) => token.length >= 2);

  let parts;
  if (cleaned.includes(",")) {
    const [last, rest] = cleaned.split(/\s*,\s*/, 2);
    parts = [...splitPart(rest || ""), ...splitPart(last || "")];
  } else {
    parts = splitPart(cleaned);
  }

  if (parts.length === 0) return "?";
  if (parts.length === 1) return parts[0][0].toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function semaiLooksLikeAttachmentLabel(text) {
  const value = (text || "").replace(/\s+/g, " ").trim();
  if (!value) return false;

  return (
    /\.(pdf|doc|docx|xls|xlsx|ppt|pptx|png|jpg|jpeg|gif|webp|heic|heif|mp4|mov|avi|mkv|webm|zip|csv|txt)\b/i.test(value) ||
    /\b(open|download|preview|attachment|attachments)\b/i.test(value)
  );
}

function semaiLooksLikeSenderName(text) {
  const value = (text || "").replace(/\s+/g, " ").trim();
  if (!value) return false;
  if (semaiLooksLikeAttachmentLabel(value)) return false;
  if (/^to:/i.test(value)) return false;
  return /[A-Za-z]/.test(value);
}

function semaiEscapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

const SEMAI_GITHUB_ISSUE_SECTION_LIMIT = 12000;
const SEMAI_GITHUB_ISSUE_BODY_LIMIT = 60000;

function semaiTrimForGitHubIssue(value, maxLength = SEMAI_GITHUB_ISSUE_SECTION_LIMIT) {
  const text = String(value || "");
  if (text.length <= maxLength) {
    return text;
  }

  return `${text.slice(0, maxLength)}\n\n[truncated ${text.length - maxLength} characters]`;
}

function semaiBuildGitHubIssueTitle(reason, subject) {
  const normalizedReason = String(reason || "").replace(/\s+/g, " ").trim();
  if (normalizedReason) {
    return normalizedReason.slice(0, 240);
  }

  return [
    "UI issue",
    subject || "Conversation"
  ]
    .map((part) => String(part || "").replace(/\s+/g, " ").trim())
    .filter(Boolean)
    .join(" | ")
    .slice(0, 240);
}

function semaiBuildFallbackGitHubIssueBody(message, subject, reason) {
  const senderName = message.sender?.name || "Unknown";
  const senderEmail = message.sender?.email || "Unknown";
  const timestamp = message.timestamp || "Unknown";
  const bubbleRole = message.isMe ? "me/right-aligned" : "them/left-aligned";
  const bubbleInitials = message.sender?.initials || semaiInitials(senderName);

  return [
    "## Reported from REMOU",
    "",
    `- Subject: ${subject || "Conversation"}`,
    `- Sender: ${senderName}`,
    `- Sender Email: ${senderEmail}`,
    `- Sender Initials: ${bubbleInitials}`,
    `- Chat Bubble: ${bubbleRole}`,
    `- Timestamp: ${timestamp}`,
    `- Page URL: ${window.location.href}`,
    "",
    "## Reason It's An Issue",
    "",
    reason || "No reason provided.",
    "",
    "## Note",
    "",
    "The richer HTML payload was rejected by GitHub validation, so this fallback report omits the raw HTML attachments."
  ].join("\n");
}

// Capture a sanitized snapshot of the reading pane for use as a test fixture.
// Strips: script tags, inline styles, GUIDs in attribute values, message body text.
// Keeps: element structure, class names, data-* attributes, ARIA roles.
function semaiCaptureFixtureHtml() {
  const CONTAINER_SELECTORS = [
    '#ReadingPaneContainerId',
    '[data-app-section="ConversationContainer"]',
    '[aria-label*="Reading Pane" i]',
    '.ReadingPane',
  ];

  let root = null;
  for (const sel of CONTAINER_SELECTORS) {
    root = document.querySelector(sel);
    if (root) break;
  }
  if (!root) return null;

  const clone = root.cloneNode(true);

  // 1. Remove script tags
  clone.querySelectorAll('script').forEach(el => el.remove());

  // 2. Remove inline styles (noisy, never needed for fixture assertions)
  clone.querySelectorAll('[style]').forEach(el => el.removeAttribute('style'));

  // 3. Redact GUIDs and base64-ish tokens in attribute values
  const GUID_RE = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi;
  const TOKEN_RE = /[A-Za-z0-9+/]{40,}={0,2}/g;
  function redactAttrs(el) {
    for (const attr of Array.from(el.attributes)) {
      let v = attr.value;
      v = v.replace(GUID_RE, 'GUID_REDACTED');
      v = v.replace(TOKEN_RE, 'TOKEN_REDACTED');
      if (v !== attr.value) el.setAttribute(attr.name, v);
    }
    for (const child of Array.from(el.children)) redactAttrs(child);
  }
  redactAttrs(clone);

  // 4. Replace actual email body text with a placeholder (keep DOM structure intact)
  clone.querySelectorAll('[aria-label="Message body"]:not([contenteditable])').forEach(body => {
    function scrubTextNodes(node) {
      for (const child of Array.from(node.childNodes)) {
        if (child.nodeType === Node.TEXT_NODE) {
          if (child.textContent.trim()) child.textContent = '[email content redacted]';
        } else if (child.nodeType === Node.ELEMENT_NODE) {
          scrubTextNodes(child);
        }
      }
    }
    scrubTextNodes(body);
  });

  return clone.outerHTML;
}

function semaiBuildGitHubIssueBody(message, subject, reason) {
  const senderName = message.sender?.name || "Unknown";
  const senderEmail = message.sender?.email || "Unknown";
  const timestamp = message.timestamp || "Unknown";
  const fixtureHtml = semaiCaptureFixtureHtml();
  const bubbleRole = message.isMe ? "me/right-aligned" : "them/left-aligned";
  const bubbleInitials = message.sender?.initials || semaiInitials(senderName);

  const parts = [
    "## Reported from REMOU",
    "",
    `- Subject: ${subject || "Conversation"}`,
    `- Sender: ${senderName}`,
    `- Sender Email: ${senderEmail}`,
    `- Sender Initials: ${bubbleInitials}`,
    `- Chat Bubble: ${bubbleRole}`,
    `- Timestamp: ${timestamp}`,
    `- Page URL: ${window.location.href}`,
    "",
    "## Reason It's An Issue",
    "",
    reason || "No reason provided.",
    "",
    "## Clean HTML",
    "",
    "```html",
    semaiTrimForGitHubIssue(message.cleanHtml),
    "```",
    "",
    "## Original HTML",
    "",
    "```html",
    semaiTrimForGitHubIssue(message.rawHtml),
    "```",
  ];

  if (fixtureHtml) {
    parts.push(
      "",
      "## Reading Pane Fixture HTML",
      "<!-- Sanitized snapshot for test fixtures. GUIDs, tokens, and email body text have been redacted. -->",
      "",
      "```html",
      semaiTrimForGitHubIssue(fixtureHtml),
      "```"
    );
  }

  return semaiTrimForGitHubIssue(parts.join("\n"), SEMAI_GITHUB_ISSUE_BODY_LIMIT);
}

async function semaiCreateGitHubIssue(message, subject, reason) {
  if (!REMOU_GITHUB_TOKEN) {
    throw new Error("Missing GitHub token in secrets.js.");
  }

  if (!REMOU_GITHUB_REPO) {
    throw new Error("Missing GitHub repo in secrets.js.");
  }

  const title = semaiBuildGitHubIssueTitle(reason, subject);
  const issueUrl = `https://api.github.com/repos/${REMOU_GITHUB_REPO}/issues`;
  const headers = {
    "Accept": "application/vnd.github+json",
    "Authorization": `Bearer ${REMOU_GITHUB_TOKEN}`,
    "Content-Type": "application/json"
  };
  const primaryBody = semaiBuildGitHubIssueBody(message, subject, reason);
  semaiNativeLog(`[semai-report] Creating GitHub issue for repo="${REMOU_GITHUB_REPO}" title="${title}" primaryBodyLength=${primaryBody.length}`);
  let response = await fetch(issueUrl, {
    method: "POST",
    headers,
    body: JSON.stringify({
      title,
      body: primaryBody
    })
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    semaiNativeLog(`[semai-report] Primary issue create failed status=${response.status} message="${errorData.message || "unknown"}" errors=${JSON.stringify(errorData.errors || [])}`);
    if (response.status === 422) {
      const fallbackBody = semaiBuildFallbackGitHubIssueBody(message, subject, reason);
      semaiNativeLog(`[semai-report] Retrying GitHub issue with fallback body length=${fallbackBody.length}`);
      response = await fetch(issueUrl, {
        method: "POST",
        headers,
        body: JSON.stringify({
          title,
          body: fallbackBody
        })
      });

      if (response.ok) {
        const fallbackIssue = await response.json();
        semaiNativeLog(`[semai-report] Fallback issue create succeeded issueNumber=${fallbackIssue.number || "unknown"} url="${fallbackIssue.html_url || ""}"`);
        return fallbackIssue;
      }

      const fallbackErrorData = await response.json().catch(() => ({}));
      semaiNativeLog(`[semai-report] Fallback issue create failed status=${response.status} message="${fallbackErrorData.message || "unknown"}" errors=${JSON.stringify(fallbackErrorData.errors || [])}`);
      const errorDetails = Array.isArray(fallbackErrorData.errors)
        ? fallbackErrorData.errors
          .map((error) => {
            if (typeof error === "string") return error;
            const field = error.field ? `${error.field}: ` : "";
            return `${field}${error.message || JSON.stringify(error)}`;
          })
          .join("; ")
        : "";
      const messageText = [fallbackErrorData.message, errorDetails].filter(Boolean).join(" — ");
      throw new Error(messageText || `GitHub API error ${response.status}`);
    }

    const errorDetails = Array.isArray(errorData.errors)
      ? errorData.errors
        .map((error) => {
          if (typeof error === "string") return error;
          const field = error.field ? `${error.field}: ` : "";
          return `${field}${error.message || JSON.stringify(error)}`;
        })
        .join("; ")
      : "";
    const messageText = [errorData.message, errorDetails].filter(Boolean).join(" — ");
    throw new Error(messageText || `GitHub API error ${response.status}`);
  }

  const createdIssue = await response.json();
  semaiNativeLog(`[semai-report] Issue create succeeded issueNumber=${createdIssue.number || "unknown"} url="${createdIssue.html_url || ""}"`);
  return createdIssue;
}

// ── Live fix preview ─────────────────────────────────────────────────────────
// Routes the Claude API call through background.js via browser.runtime.sendMessage.
// Content scripts can't call api.anthropic.com directly (CORS blocks it since
// the request origin is outlook.cloud.microsoft). background.js runs under the
// extension's origin and bypasses CORS.

// ── Claude API constants (content-script side — avoids service worker lifecycle issues) ──

const SEMAI_APPLY_FIX_TOOL = {
  name: 'apply_fix',
  description: 'Apply a CSS or JS patch to fix the reported Outlook rendering issue.',
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
        description: 'A regex pattern matching Outlook URLs where this patch should apply.',
      },
    },
    required: ['explanation', 'patchType', 'patchCode', 'urlPattern'],
  },
};

const SEMAI_PREVIEW_FIX_SYSTEM_PROMPT = [
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

// Pre-fetch sig detector source once (content script stays alive with the page).
let _semaiSigDetectorSource = null;
fetch(chrome.runtime.getURL('semaiSigDetector.js'))
  .then(r => r.text())
  .then(src => { _semaiSigDetectorSource = src; })
  .catch(() => {});

async function semaiRequestPreviewFix(message, subject, reason, conversationHistory = null) {
  const apiKey = typeof REMOU_ANTHROPIC_API_KEY !== "undefined" ? REMOU_ANTHROPIC_API_KEY : null;
  if (!apiKey) throw new Error("Anthropic API key not configured in secrets.js");

  if (conversationHistory?.length) {
    console.log("[semai-preview] Retrying — turns:", conversationHistory.length);
  }

  const cleanHtml = (message.cleanHtml || "").slice(0, 6000);
  const sigDetectorSource = _semaiSigDetectorSource;

  const userMessage = [
    '## Bug report',
    'The user reported an issue while viewing: ' + window.location.href,
    'Subject: ' + (subject || '(no subject)'),
    'Sender: ' + (message.sender?.name || 'Unknown') + ' <' + (message.sender?.email || 'unknown') + '>',
    '',
    '## User description',
    reason || '(no description)',
    '',
    ...(cleanHtml ? ['## Email HTML (clean, sig stripped)', '```html', cleanHtml, '```', ''] : []),
    ...(sigDetectorSource ? ['## Extension source: semaiSigDetector.js', '```javascript', sigDetectorSource, '```', ''] : []),
  ].join('\n');

  const messages = [
    { role: 'user', content: userMessage },
    ...(Array.isArray(conversationHistory) ? conversationHistory : []),
  ];

  const body = JSON.stringify({
    model: 'claude-sonnet-4-6',
    max_tokens: 8000,
    system: SEMAI_PREVIEW_FIX_SYSTEM_PROMPT,
    tools: [SEMAI_APPLY_FIX_TOOL],
    tool_choice: { type: 'tool', name: 'apply_fix' },
    messages,
  });

  console.log('[semai-preview] Calling Anthropic API — turns:', messages.length, 'body:', body.length, 'chars');

  const res = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'anthropic-dangerous-direct-browser-access': 'true',
      'content-type': 'application/json',
    },
    body,
  });

  if (!res.ok) {
    const t = await res.text();
    const parsed = (() => { try { return JSON.parse(t); } catch { return {}; } })();
    throw new Error(parsed.error?.message || 'HTTP ' + res.status);
  }

  const data = await res.json();
  console.log('[semai-preview] Response stop_reason:', data.stop_reason, 'input_tokens:', data.usage?.input_tokens);

  const toolUse = data.content?.find(b => b.type === 'tool_use' && b.name === 'apply_fix');
  if (!toolUse) throw new Error('Claude did not return a fix suggestion.');

  return {
    ok: true,
    toolUseId: toolUse.id,
    explanation: toolUse.input.explanation,
    patchType: toolUse.input.patchType,
    patchCode: toolUse.input.patchCode,
    urlPattern: toolUse.input.urlPattern,
  };
}

// Tracks the last injected CSS so it can be removed on reject.
// JS patches can't be undone without a page reload.
let _semaiPreviewCssCode = null;

function semaiInjectPreviewPatch(patchType, patchCode) {
  semaiRemovePreviewPatch();
  if (patchType === 'css') _semaiPreviewCssCode = patchCode;
  chrome.runtime.sendMessage(
    { type: 'INJECT_PATCH', payload: { patchType, patchCode } },
    (response) => {
      if (chrome.runtime.lastError) {
        semaiNativeLog(`[semai-preview] Inject failed: ${chrome.runtime.lastError.message}`);
      } else if (!response?.ok) {
        semaiNativeLog(`[semai-preview] Inject error: ${response?.error}`);
      } else {
        semaiNativeLog(`[semai-preview] Patch injected (${patchType})`);
      }
    }
  );
}

function semaiRemovePreviewPatch() {
  // Remove CSS injected via chrome.scripting
  if (_semaiPreviewCssCode) {
    const css = _semaiPreviewCssCode;
    _semaiPreviewCssCode = null;
    chrome.runtime.sendMessage({ type: 'REMOVE_CSS_PATCH', payload: { css } });
  }
  // JS patches can't be removed without a page reload — acceptable for preview
  semaiRemoveTestBanner();
}

let _semaiTestBannerRow = null;

function semaiInjectTestBanner(row) {
  semaiRemoveTestBanner();
  if (!row) return;
  _semaiTestBannerRow = row;
  const banner = document.createElement('div');
  banner.setAttribute('data-semai-test-banner', '1');
  banner.style.cssText = [
    'display:block',
    'margin:6px 12px',
    'padding:8px 12px',
    'background:#007aff',
    'color:#fff',
    'border-radius:8px',
    'font-size:13px',
    'font-weight:600',
    'font-family:system-ui,sans-serif',
    'z-index:9999',
  ].join(';');
  banner.textContent = 'Hello — Claude injected this ✓';
  row.prepend(banner);
}

function semaiRemoveTestBanner() {
  if (_semaiTestBannerRow) {
    const existing = _semaiTestBannerRow.querySelector('[data-semai-test-banner]');
    if (existing) existing.remove();
    _semaiTestBannerRow = null;
  }
}

function semaiBuildApprovedFixIssueBody(message, subject, reason, patch) {
  const baseBody = semaiBuildGitHubIssueBody(message, subject, reason);
  const approvedSection = [
    "",
    "## Approved Fix",
    "",
    `- Patch Type: ${patch.patchType}`,
    `- URL Pattern: \`${patch.urlPattern}\``,
    "",
    "### Explanation",
    patch.explanation,
    "",
    "### Patch Code",
    "",
    "```" + patch.patchType,
    patch.patchCode,
    "```",
    "",
    "<!-- SEMAI_APPROVED_PATCH",
    JSON.stringify({
      patchType: patch.patchType,
      patchCode: patch.patchCode,
      urlPattern: patch.urlPattern,
      explanation: patch.explanation
    }),
    "SEMAI_APPROVED_PATCH -->"
  ].join("\n");

  return semaiTrimForGitHubIssue(
    baseBody + approvedSection,
    SEMAI_GITHUB_ISSUE_BODY_LIMIT
  );
}

async function semaiCreateApprovedFixIssue(message, subject, reason, patch) {
  if (!REMOU_GITHUB_TOKEN) throw new Error("Missing GitHub token in secrets.js.");
  if (!REMOU_GITHUB_REPO) throw new Error("Missing GitHub repo in secrets.js.");

  const title = semaiBuildGitHubIssueTitle(reason, subject);
  const response = await fetch(`https://api.github.com/repos/${REMOU_GITHUB_REPO}/issues`, {
    method: "POST",
    headers: {
      "Accept": "application/vnd.github+json",
      "Authorization": `Bearer ${REMOU_GITHUB_TOKEN}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      title,
      body: semaiBuildApprovedFixIssueBody(message, subject, reason, patch),
      labels: ["auto-fix"]
    })
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    const errorDetails = Array.isArray(errorData.errors)
      ? errorData.errors.map(e => typeof e === "string" ? e : `${e.field || ""}: ${e.message || JSON.stringify(e)}`).join("; ")
      : "";
    throw new Error([errorData.message, errorDetails].filter(Boolean).join(" — ") || `GitHub API error ${response.status}`);
  }

  return response.json();
}

function semaiClearReportHover() {
  if (semaiReportHoverRow) {
    semaiReportHoverRow.classList.remove("semai-chat-row-report-hover");
    semaiReportHoverRow = null;
  }

  if (semaiReportHoverOriginalBody) {
    semaiReportHoverOriginalBody.classList.remove("semai-original-message-report-hover");
    semaiReportHoverOriginalBody = null;
  }
}

function semaiClearMissedOriginalHighlights() {
  semaiReportMissedBodies.forEach((bodyEl) => {
    bodyEl.classList.remove("semai-original-message-report-missed");
  });
  semaiReportMissedBodies = [];
}

function semaiHighlightMissedOriginalMessages(overlay) {
  semaiClearMissedOriginalHighlights();

  const messages = overlay?._semaiMessages || [];
  messages.forEach((message) => {
    if (message.cleanHtml) return;
    if (!(message.sourceBodyEl instanceof HTMLElement)) return;

    message.sourceBodyEl.classList.add("semai-original-message-report-missed");
    semaiReportMissedBodies.push(message.sourceBodyEl);
  });
}

function semaiSetReportModeStatus(overlay, message, tone = "neutral") {
  const status = overlay?.querySelector("#semai-chat-reply-status");
  if (!status) return;
  status.textContent = message;
  status.dataset.tone = tone;
}

function semaiHandleReportModeKeydown(event) {
  if (event.key !== "Escape") return;

  if (!semaiReportModeOverlay) return;

  event.preventDefault();
  semaiExitReportMode(semaiReportModeOverlay);
}

function semaiCloseReportPopover() {
  semaiReportPopoverEl?.remove();
  semaiReportPopoverEl = null;
}

function semaiOpenReportPopover(overlay, message, subject, clientX, clientY, reportRow = null) {
  semaiCloseReportPopover();

  const popover = document.createElement("div");
  popover.className = "semai-report-popover";
  popover.innerHTML = `
    <div class="semai-report-popover-title">Report issue</div>
    <textarea
      class="semai-report-popover-input"
      rows="4"
      placeholder="What is wrong with this email bubble?"
    ></textarea>
    <div class="semai-report-popover-actions">
      <button type="button" class="semai-report-popover-cancel">Cancel</button>
      <!-- <button type="button" class="semai-report-popover-preview">Preview</button> -->
      <button type="button" class="semai-report-popover-send">Report</button>
    </div>
    <div class="semai-report-popover-preview-result" style="display:none;">
      <details class="semai-report-popover-drawer">
        <summary class="semai-report-popover-drawer-summary">Claude's explanation</summary>
        <div class="semai-report-popover-explanation"></div>
      </details>
      <div class="semai-report-popover-preview-actions">
        <button type="button" class="semai-report-popover-reject" title="Reject — ask Claude for a different fix">✕</button>
        <button type="button" class="semai-report-popover-approve" title="Approve — report to GitHub with this fix">✓</button>
      </div>
    </div>
  `;

  const cancelBtn = popover.querySelector(".semai-report-popover-cancel");
  const previewBtn = popover.querySelector(".semai-report-popover-preview");
  const sendBtn = popover.querySelector(".semai-report-popover-send");
  const previewResult = popover.querySelector(".semai-report-popover-preview-result");
  const explanationEl = popover.querySelector(".semai-report-popover-explanation");
  const drawerEl = popover.querySelector(".semai-report-popover-drawer");
  const rejectBtn = popover.querySelector(".semai-report-popover-reject");
  const approveBtn = popover.querySelector(".semai-report-popover-approve");
  const input = popover.querySelector(".semai-report-popover-input");

  let currentPatch = null;
  let conversationHistory = []; // multi-turn retry history

  function setAllDisabled(disabled) {
    cancelBtn.disabled = disabled;
    if (previewBtn) previewBtn.disabled = disabled;
    sendBtn.disabled = disabled;
    input.disabled = disabled;
    rejectBtn.disabled = disabled;
    approveBtn.disabled = disabled;
  }

  cancelBtn.addEventListener("click", () => {
    semaiRemovePreviewPatch();
    semaiExitReportMode(overlay);
  });

  // ── Preview Fix: call Claude via background.js ──
  previewBtn?.addEventListener("click", async () => {
    const reason = input.value.trim();
    if (!reason) {
      input.focus();
      semaiSetReportModeStatus(overlay, "Type the reason before previewing a fix.", "error");
      return;
    }

    conversationHistory = []; // fresh request resets history
    setAllDisabled(true);
    previewResult.style.display = "none";
    semaiRemovePreviewPatch();
    semaiSetReportModeStatus(overlay, "Asking Claude for a fix…", "report");

    try {
      const patch = await semaiRequestPreviewFix(message, subject, reason, conversationHistory);
      semaiNativeLog(`[semai-preview] Patch received patchType=${patch.patchType} codeLength=${(patch.patchCode||'').length}`);
      currentPatch = patch;
      semaiInjectPreviewPatch(patch.patchType, patch.patchCode);
      semaiInjectTestBanner(reportRow);
      explanationEl.textContent = patch.explanation;
      if (drawerEl) drawerEl.open = true;
      previewResult.style.display = "block";
      setAllDisabled(false);
      semaiSetReportModeStatus(overlay, "Fix applied — ✓ approve or ✕ reject for a different fix.", "success");
    } catch (error) {
      setAllDisabled(false);
      semaiNativeLog(`[semai-preview] Preview fix failed: ${error.message}`);
      semaiSetReportModeStatus(overlay, error.message || "Preview fix failed.", "error");
    }
  });

  // ── Reject: add to conversation history and ask Claude for a different fix ──
  rejectBtn.addEventListener("click", async () => {
    if (!currentPatch) return;
    const rejectedPatch = currentPatch;

    // Build the next conversation turn
    conversationHistory.push({
      role: "assistant",
      content: [{ type: "tool_use", id: rejectedPatch.toolUseId, name: "apply_fix", input: {
        explanation: rejectedPatch.explanation,
        patchType: rejectedPatch.patchType,
        patchCode: rejectedPatch.patchCode,
        urlPattern: rejectedPatch.urlPattern,
      }}],
    });
    conversationHistory.push({
      role: "user",
      content: [{ type: "tool_result", tool_use_id: rejectedPatch.toolUseId,
        content: "The patch was rejected. The visual result did not look correct. Please suggest a different approach — try a different strategy or fix a different root cause." }],
    });

    semaiRemovePreviewPatch();
    currentPatch = null;
    previewResult.style.display = "none";
    setAllDisabled(true);
    semaiSetReportModeStatus(overlay, "Asking Claude for a different fix…", "report");

    try {
      const patch = await semaiRequestPreviewFix(message, subject, input.value.trim(), conversationHistory);
      semaiNativeLog(`[semai-preview] Retry patch received patchType=${patch.patchType}`);
      currentPatch = patch;
      semaiInjectPreviewPatch(patch.patchType, patch.patchCode);
      semaiInjectTestBanner(reportRow);
      explanationEl.textContent = patch.explanation;
      if (drawerEl) drawerEl.open = true;
      previewResult.style.display = "block";
      setAllDisabled(false);
      semaiSetReportModeStatus(overlay, "New fix applied — ✓ approve or ✕ reject for another.", "success");
    } catch (error) {
      setAllDisabled(false);
      semaiNativeLog(`[semai-preview] Retry failed: ${error.message}`);
      semaiSetReportModeStatus(overlay, error.message || "Retry failed.", "error");
    }
  });

  // ── Approve & Report: create issue with the working patch embedded ──
  approveBtn.addEventListener("click", async () => {
    if (!currentPatch) return;
    const reason = input.value.trim();

    setAllDisabled(true);
    semaiSetReportModeStatus(overlay, "Creating GitHub issue with approved fix…", "report");

    try {
      await semaiCreateApprovedFixIssue(message, subject, reason, currentPatch);
      semaiRemovePreviewPatch();
      semaiCloseReportPopover();
      semaiExitReportMode(overlay, "Issue reported with approved fix.", "success");
    } catch (error) {
      setAllDisabled(false);
      semaiSetReportModeStatus(overlay, error.message || "Failed to create GitHub issue.", "error");
    }
  });

  // ── Report Only: existing behavior (no fix preview) ──
  sendBtn.addEventListener("click", async () => {
    const reason = input.value.trim();
    if (!reason) {
      input.focus();
      semaiSetReportModeStatus(overlay, "Type the reason this message is misbehaving.", "error");
      return;
    }

    setAllDisabled(true);
    semaiRemovePreviewPatch();
    semaiSetReportModeStatus(overlay, "Creating GitHub issue…", "report");

    try {
      await semaiCreateGitHubIssue(message, subject, reason);
      semaiCloseReportPopover();
      semaiExitReportMode(overlay, "Issue has been reported successfully.", "success");
    } catch (error) {
      setAllDisabled(false);
      semaiSetReportModeStatus(overlay, error.message || "Failed to create GitHub issue.", "error");
    }
  });

  popover.addEventListener("click", (event) => {
    event.stopPropagation();
  });

  document.body.appendChild(popover);
  semaiReportPopoverEl = popover;

  const margin = 12;
  const rect = popover.getBoundingClientRect();
  const left = Math.min(clientX + 16, window.innerWidth - rect.width - margin);
  const top = Math.min(clientY + 16, window.innerHeight - rect.height - margin);
  popover.style.left = `${Math.max(margin, left)}px`;
  popover.style.top = `${Math.max(margin, top)}px`;

  // ── Drag to reposition ──
  const titleEl = popover.querySelector(".semai-report-popover-title");
  let dragStartX = 0, dragStartY = 0, dragOrigLeft = 0, dragOrigTop = 0;

  titleEl.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    e.preventDefault();
    dragStartX = e.clientX;
    dragStartY = e.clientY;
    dragOrigLeft = parseInt(popover.style.left, 10) || 0;
    dragOrigTop  = parseInt(popover.style.top,  10) || 0;

    function onMove(ev) {
      const dx = ev.clientX - dragStartX;
      const dy = ev.clientY - dragStartY;
      const newLeft = Math.max(margin, Math.min(window.innerWidth  - popover.offsetWidth  - margin, dragOrigLeft + dx));
      const newTop  = Math.max(margin, Math.min(window.innerHeight - popover.offsetHeight - margin, dragOrigTop  + dy));
      popover.style.left = `${newLeft}px`;
      popover.style.top  = `${newTop}px`;
    }

    function onUp() {
      document.removeEventListener("mousemove", onMove);
      document.removeEventListener("mouseup",   onUp);
    }

    document.addEventListener("mousemove", onMove);
    document.addEventListener("mouseup",   onUp);
  });

  semaiSetReportModeStatus(overlay, "Describe the issue, then report it.", "report");
  input.focus();
}

function semaiExitReportMode(overlay, statusMessage, tone = "neutral") {
  if (!overlay) return;

  overlay.classList.remove("semai-chat-report-mode");
  overlay.dataset.reportMode = "inactive";
  document.removeEventListener("keydown", semaiHandleReportModeKeydown, true);
  semaiReportModeOverlay = null;
  semaiClearReportHover();
  semaiClearMissedOriginalHighlights();
  semaiCloseReportPopover();

  const reportButton = overlay.querySelector("#semai-chat-report-issue-btn");
  if (reportButton) {
    reportButton.classList.remove("semai-chat-report-issue-btn-active");
    reportButton.disabled = false;
  }

  if (statusMessage) {
    semaiSetReportModeStatus(overlay, statusMessage, tone);
    return;
  }

  semaiUpdateOverlayViewToggle(overlay);
}

function semaiEnterReportMode(overlay) {
  if (!overlay) return;

  overlay.dataset.reportMode = "active";
  overlay.classList.add("semai-chat-report-mode");
  document.addEventListener("keydown", semaiHandleReportModeKeydown, true);
  semaiReportModeOverlay = overlay;

  const reportButton = overlay.querySelector("#semai-chat-report-issue-btn");
  if (reportButton) {
    reportButton.classList.add("semai-chat-report-issue-btn-active");
  }

  semaiHighlightMissedOriginalMessages(overlay);

  semaiSetReportModeStatus(
    overlay,
    "Hover an email, click to choose it, or press Esc to cancel.",
    "report"
  );
}

function semaiToggleReportMode(overlay) {
  if (!overlay) return;

  if (overlay.dataset.reportMode === "active") {
    semaiExitReportMode(
      overlay,
      overlay.dataset.viewMode === "real"
        ? "The original Outlook thread is visible above the reply box. Use the eye button to switch back to chat."
        : "Chat view is on. Use the eye button to switch back to regular Outlook."
    );
    return;
  }

  semaiEnterReportMode(overlay);
}

function semaiGetReportRowFromEventTarget(target) {
  if (!(target instanceof Element)) return null;
  return target.closest(".semai-chat-row[data-report-index]");
}

function semaiGetOriginalReportBodyFromEventTarget(target) {
  if (!(target instanceof Element) || semaiIsInsideRemouUi(target)) return null;
  return target.closest('[aria-label="Message body"]:not([contenteditable])');
}

function semaiGetMessageForOriginalBody(overlay, bodyEl) {
  const messages = overlay?._semaiMessages || [];
  return messages.find((message) => message.sourceBodyEl === bodyEl) || null;
}

function semaiHandleReportRowHover(event) {
  if (!semaiReportModeOverlay || semaiReportModeOverlay.dataset.reportMode !== "active") {
    semaiClearReportHover();
    return;
  }

  const row = semaiGetReportRowFromEventTarget(event.target);
  if (row === semaiReportHoverRow) return;

  semaiClearReportHover();
  if (!row) return;

  semaiReportHoverRow = row;
  semaiReportHoverRow.classList.add("semai-chat-row-report-hover");
}

function semaiHandleOriginalReportHover(event) {
  const overlay = semaiReportModeOverlay;
  if (!overlay || overlay.dataset.reportMode !== "active" || overlay.dataset.viewMode !== "real") {
    if (semaiReportHoverOriginalBody) {
      semaiReportHoverOriginalBody.classList.remove("semai-original-message-report-hover");
      semaiReportHoverOriginalBody = null;
    }
    return;
  }

  const bodyEl = semaiGetOriginalReportBodyFromEventTarget(event.target);
  const message = bodyEl ? semaiGetMessageForOriginalBody(overlay, bodyEl) : null;
  const nextBody = message?.sourceBodyEl || null;

  if (nextBody === semaiReportHoverOriginalBody) return;

  if (semaiReportHoverOriginalBody) {
    semaiReportHoverOriginalBody.classList.remove("semai-original-message-report-hover");
  }

  semaiReportHoverOriginalBody = nextBody;
  if (semaiReportHoverOriginalBody) {
    semaiReportHoverOriginalBody.classList.add("semai-original-message-report-hover");
  }
}

async function semaiHandleReportRowClick(event) {
  const overlay = semaiReportModeOverlay;
  if (!overlay || overlay.dataset.reportMode !== "active") return;

  const row = semaiGetReportRowFromEventTarget(event.target);
  if (!row) return;

  event.preventDefault();
  event.stopPropagation();

  const index = Number(row.dataset.reportIndex);
  const message = overlay._semaiMessages?.[index];
  const subject = overlay._semaiSubject || "Conversation";
  if (!message) return;

  semaiOpenReportPopover(overlay, message, subject, event.clientX, event.clientY, row);
}

function semaiHandleOriginalReportClick(event) {
  const overlay = semaiReportModeOverlay;
  if (!overlay || overlay.dataset.reportMode !== "active" || overlay.dataset.viewMode !== "real") return;

  const bodyEl = semaiGetOriginalReportBodyFromEventTarget(event.target);
  if (!bodyEl) return;

  const message = semaiGetMessageForOriginalBody(overlay, bodyEl);
  if (!message) return;

  event.preventDefault();
  event.stopPropagation();

  const subject = overlay._semaiSubject || "Conversation";
  semaiOpenReportPopover(overlay, message, subject, event.clientX, event.clientY, bodyEl);
}

function semaiGetCalibration() {
  try {
    const raw = window.localStorage.getItem(SEMAI_CALIBRATION_STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch (e) {
    return null;
  }
}

function semaiSetCalibration(calibration) {
  try {
    window.localStorage.setItem(SEMAI_CALIBRATION_STORAGE_KEY, JSON.stringify(calibration));
  } catch (e) {
    console.warn("[semai] Failed to persist calibration", e);
  }
}

function semaiBuildSenderSelector(el) {
  if (!(el instanceof Element)) return null;
  if (el.id) return `#${CSS.escape(el.id)}`;

  const stableClasses = Array.from(el.classList).filter(Boolean).slice(0, 2);
  if (stableClasses.length > 0) {
    return `${el.tagName.toLowerCase()}.${stableClasses.map((cls) => CSS.escape(cls)).join(".")}`;
  }

  return el.tagName.toLowerCase();
}

function semaiUpdateCalibrationStatus(message, tone = "neutral") {
  const status = document.getElementById("semai-calibration-status");
  if (!status) return;
  status.textContent = message;
  status.dataset.tone = tone;
}

function semaiClearCalibrationHover() {
  if (semaiCalibrationHoverEl) {
    semaiCalibrationHoverEl.classList.remove("semai-calibration-target");
    semaiCalibrationHoverEl = null;
  }
}

function semaiStopCalibration(message = "Calibration cancelled.", tone = "neutral") {
  semaiCalibrationState = null;
  semaiClearCalibrationHover();
  document.body.classList.remove("semai-calibrating");
  document.removeEventListener("mousemove", semaiHandleCalibrationHover, true);
  document.removeEventListener("click", semaiHandleCalibrationClick, true);
  document.removeEventListener("keydown", semaiHandleCalibrationKeydown, true);
  semaiUpdateCalibrationStatus(message, tone);
}

function semaiFindCalibrationTarget(startEl) {
  if (!(startEl instanceof Element)) return null;

  const candidate = startEl.closest(
    '.OZZZK, [data-testid="senderName"], [class*="senderName" i], [class*="sender-name" i], .ms-Persona-primaryText'
  );
  if (!candidate) return null;

  const text = (candidate.innerText || candidate.textContent || "").trim();
  if (!text || text.length > 160) return null;
  if (!semaiLooksLikeSenderName(text)) return null;

  return candidate;
}

function semaiHandleCalibrationHover(event) {
  if (!semaiCalibrationState) return;

  const nextTarget = semaiFindCalibrationTarget(event.target);
  if (nextTarget === semaiCalibrationHoverEl) return;

  semaiClearCalibrationHover();
  if (nextTarget) {
    semaiCalibrationHoverEl = nextTarget;
    semaiCalibrationHoverEl.classList.add("semai-calibration-target");
  }
}

function semaiFinishCalibration(selfLabel, otherLabel, selector) {
  const selfSender = semaiNormalizeSenderLabel(selfLabel);
  const otherSender = semaiNormalizeSenderLabel(otherLabel);
  const calibration = {
    senderSelector: selector || null,
    selfName: selfSender.name || null,
    selfEmail: selfSender.email || null,
    sampleOtherName: otherSender.name || null
  };

  semaiSetCalibration(calibration);
  semaiCurrentUser = null;
  semaiGetCurrentUser();
  semaiUpdateChatToggleVisibility();
  semaiUpdateChatToggleBtn();
  semaiDismissOnboardingModal();
  const calibrateBtn = document.querySelector(".semai-calibrate-btn");
  if (calibrateBtn) calibrateBtn.textContent = "Retrain sender detection";
  semaiStopCalibration(`Saved. Using "${selfSender.name}" as you.`, "success");
}

function semaiHandleCalibrationKeydown(event) {
  if (!semaiCalibrationState) return;
  if (event.key !== "Escape") return;

  event.preventDefault();
  event.stopPropagation();
  semaiStopCalibration();
}

function semaiHandleCalibrationClick(event) {
  if (!semaiCalibrationState) return;

  const target = semaiFindCalibrationTarget(event.target);
  if (!target) return;

  const text = (target.innerText || target.textContent || "").trim();

  event.preventDefault();
  event.stopPropagation();

  if (semaiCalibrationState.step === "self") {
    semaiCalibrationState.selfLabel = text;
    semaiCalibrationState.selector = semaiBuildSenderSelector(target);
    semaiCalibrationState.step = "other";
    semaiUpdateCalibrationStatus(
      `Step 2: Click on another sender (who is not you).`,
      "other"
    );
    return;
  }

  semaiFinishCalibration(
    semaiCalibrationState.selfLabel,
    text,
    semaiCalibrationState.selector || semaiBuildSenderSelector(target)
  );
}

function semaiStartCalibration() {
  semaiStopCalibration("Step 1: Click your sender name only. Do not click a To/recipient field. Step 2: Click another sender who is not you.", "neutral");
  document.removeEventListener("click", semaiHandleCalibrationClick, true);
  document.removeEventListener("mousemove", semaiHandleCalibrationHover, true);
  semaiClearCalibrationHover();
  semaiCalibrationState = {
    step: "self",
    selfLabel: "",
    selector: null
  };

  document.body.classList.add("semai-calibrating");
  semaiUpdateCalibrationStatus("Step 1: Click on your sender name only, not a To/recipient field.", "self");
  document.addEventListener("mousemove", semaiHandleCalibrationHover, true);
  document.addEventListener("click", semaiHandleCalibrationClick, true);
  document.addEventListener("keydown", semaiHandleCalibrationKeydown, true);
}

// ===== ONBOARDING MODAL =====

function semaiHandleOnboardingKeydown(event) {
  if (event.key !== "Escape") return;

  event.preventDefault();
  semaiDismissOnboardingModal();
}

function semaiShowOnboardingModal() {
  if (document.getElementById("semai-onboarding-modal")) return;

  const modal = document.createElement("div");
  modal.id = "semai-onboarding-modal";
  modal.innerHTML = `
    <div class="semai-onboarding-card">
      <button
        class="semai-onboarding-close"
        type="button"
        id="semai-onboarding-close-btn"
        aria-label="Close setup"
        title="Close"
      >
        ×
      </button>
      <div class="semai-onboarding-logo">
        <div class="semai-logo-dot" style="width:14px;height:14px;margin-right:8px;flex-shrink:0;"></div>
        <span style="font-size:14px;font-weight:700;letter-spacing:0.06em;text-transform:uppercase;color:#0f172a;">Remou</span>
      </div>
      <h2 class="semai-onboarding-headline">One quick setup before you start</h2>
      <p class="semai-onboarding-body">
        Remou needs to know who you are so it can tell your messages apart from others in chat view.
      </p>
      <ol class="semai-onboarding-steps">
        <li>Click "Start setup" below — the panel will enter setup mode.</li>
        <li>An email thread will appear highlighted — click on your name where it shows as the sender, not in any To/recipient line.</li>
        <li>Then click on any other person's name in a different message.</li>
      </ol>
      <p class="semai-onboarding-body" style="margin-top:0;">
        That's it. Remou will remember your identity for future sessions.
      </p>
      <button class="semai-onboarding-cta" type="button" id="semai-onboarding-cta-btn">Start setup →</button>
      <p class="semai-onboarding-note">You can redo this anytime from the Remou panel.</p>
    </div>
  `;

  modal.addEventListener("click", (event) => {
    if (event.target === modal) {
      semaiDismissOnboardingModal();
    }
  });
  modal.querySelector("#semai-onboarding-close-btn").addEventListener("click", () => {
    semaiDismissOnboardingModal();
  });
  modal.querySelector("#semai-onboarding-cta-btn").addEventListener("click", () => {
    semaiDismissOnboardingModal();
    semaiStartCalibration();
  });

  document.body.appendChild(modal);
  document.addEventListener("keydown", semaiHandleOnboardingKeydown, true);
}

function semaiDismissOnboardingModal() {
  const modal = document.getElementById("semai-onboarding-modal");
  if (modal) modal.remove();
  document.removeEventListener("keydown", semaiHandleOnboardingKeydown, true);
}

function semaiNodePrecedesBody(node, bodyEl) {
  return Boolean(node.compareDocumentPosition(bodyEl) & Node.DOCUMENT_POSITION_FOLLOWING);
}

function semaiPickClosestPrecedingMatch(container, bodyEl, selectors) {
  const matches = [];

  for (const sel of selectors) {
    try {
      container.querySelectorAll(sel).forEach((node) => {
        if (!(node instanceof Element)) return;
        if (node === bodyEl || node.contains(bodyEl) || bodyEl.contains(node)) return;
        if (!semaiNodePrecedesBody(node, bodyEl)) return;
        matches.push(node);
      });
    } catch (e) { /* skip invalid selector */ }
  }

  if (matches.length === 0) return null;
  return matches[matches.length - 1];
}

// Returns the full text content of an element, including text inside CSS-hidden child
// nodes (e.g. mark.js search-highlight spans). Unlike innerText, textContent is not
// affected by display/visibility styles, so highlighted words are never silently dropped.
function semaiFullText(el) {
  // textContent always includes text from every descendant node regardless of CSS
  // visibility, so highlighted search-term spans (mark.js) are never silently dropped.
  // We only trim edges; internal newlines are preserved so semaiNormalizeSenderLabel
  // can still split and filter lines correctly (e.g. avatar initials vs real name).
  return (el?.textContent || el?.innerText || "").trim();
}

function semaiGetSenderLabelNearBody(bodyEl) {
  const bodyContainer = bodyEl.closest('[data-test-id="mailMessageBodyContainer"]');
  if (!bodyContainer || !bodyContainer.parentElement) return null;
  const calibration = semaiGetCalibration();

  let sibling = bodyContainer.previousElementSibling;
  while (sibling) {
    if (calibration?.senderSelector) {
      if (sibling.matches?.(calibration.senderSelector)) {
        const text = semaiFullText(sibling);
        if (semaiLooksLikeSenderName(text)) return text;
      }

      const calibratedLabel = sibling.querySelector?.(calibration.senderSelector);
      if (calibratedLabel) {
        const text = semaiFullText(calibratedLabel);
        if (semaiLooksLikeSenderName(text)) return text;
      }
    }

    if (sibling.matches(".OZZZK")) {
      const text = semaiFullText(sibling);
      if (semaiLooksLikeSenderName(text)) return text;
    }

    const directLabel = sibling.querySelector?.(".OZZZK");
    if (directLabel) {
      const text = semaiFullText(directLabel);
      if (semaiLooksLikeSenderName(text)) return text;
    }

    const text = semaiFullText(sibling);
    if (text && text.length <= 120 && semaiLooksLikeSenderName(text)) {
      return text;
    }
    sibling = sibling.previousElementSibling;
  }

  return null;
}

function semaiNormalizeSenderLabel(rawLabel) {
  const raw = (rawLabel || "").trim();
  if (!raw) return { name: "", email: "" };

  const emailMatch = raw.match(/<?([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})>?/i);
  const email = emailMatch ? emailMatch[1].toLowerCase() : "";

  const candidateLines = raw
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);

  let name = "";

  for (const line of candidateLines) {
    if (/^to:/i.test(line)) continue;
    if (line === email) continue;

    const cleanedLine = line
      .replace(/^from:\s*/i, "")
      .replace(/<[^>]+@[^>]+>/g, "")
      .trim();

    if (!cleanedLine) continue;
    if (!/[A-Za-z]/.test(cleanedLine)) continue;
    if (cleanedLine.length > 80) continue;
    // Skip Outlook persona avatar initials (e.g. "AA", "SA") — all-caps 2–4 chars
    if (/^[A-Z]{2,4}$/.test(cleanedLine)) continue;

    name = cleanedLine;
    break;
  }

  if (!name) {
    name = raw
      .replace(/^from:\s*/i, "")
      .split(/[<(\n]/)[0]
      .trim();
  }

  return { name, email };
}

// Detect the logged-in Outlook user (cached once per session).
// Returns { name, nameLower, email, initials } or null.
function semaiGetCurrentUser() {
  if (semaiCurrentUser) return semaiCurrentUser;

  function trySet(name, email) {
    const n = (name || "").trim();
    const e = (email || "").trim().toLowerCase();
    if (n.length >= 2) {
      semaiCurrentUser = {
        name: n,
        nameLower: n.toLowerCase(),
        email: e || null,
        initials: semaiInitials(n)
      };
      console.log("[semai] Current user detected", semaiCurrentUser);
      return true;
    }
    return false;
  }

  const calibration = semaiGetCalibration();
  if (calibration?.selfName) {
    trySet(calibration.selfName, calibration.selfEmail || "");
    return semaiCurrentUser;
  }

  // ── Strategy 1: UI-based — try many selectors ──
  const uiSelectors = [
    '#mectrl_currentAccount_primary',
    '#mectrl_headerPicture',
    '#O365_MainLink_Me',
    'button[data-tid="AccountManagerButton"]',
    'button#mectrl_main_trigger',
    '#O365_HeaderRightRegion .ms-Persona-primaryText',
    '#meInitialsButton',
    '[class*="mectrl" i] [class*="name" i]',
    '[class*="mectrl" i] [class*="primary" i]',
    '[data-automationid="personaInfo"]',
  ];

  for (const sel of uiSelectors) {
    try {
      const el = document.querySelector(sel);
      if (!el) continue;
      const label = (el.getAttribute("aria-label") || "").replace(/^account\s+manager\s+for\s+/i, "").trim();
      if (label.length >= 2 && /[A-Za-z]/.test(label) && !/^(settings|search|help|reactions|show)/i.test(label)) {
        if (trySet(label)) return semaiCurrentUser;
      }
      const title = (el.getAttribute("title") || "").trim();
      if (title.length >= 2 && /[A-Za-z]/.test(title)) {
        if (trySet(title)) return semaiCurrentUser;
      }
      const text = (el.innerText || el.textContent || "").trim();
      if (text.length >= 2 && text.length <= 60 && /[A-Za-z]/.test(text)) {
        if (trySet(text)) return semaiCurrentUser;
      }
    } catch (e) { /* skip */ }
  }

  // ── Strategy 2: Email from header ──
  const emailSelectors = [
    '#mectrl_currentAccount_secondary',
    '#O365_HeaderRightRegion [class*="email" i]',
    '[class*="mectrl" i] [class*="secondary" i]',
  ];
  for (const sel of emailSelectors) {
    try {
      const el = document.querySelector(sel);
      if (!el) continue;
      const email = (el.innerText || el.textContent || "").trim();
      if (email.includes("@")) {
        const localPart = email.split("@")[0].replace(/[._]/g, " ");
        if (trySet(localPart, email)) return semaiCurrentUser;
      }
    } catch (e) { /* skip */ }
  }

  console.log("[semai] Current user detection failed — complete Remou setup to identify your account.");
  return null;
}

// Extract sender info from the message card surrounding a body element
function semaiGetMessageSender(bodyEl) {
  const nameSelectors = [
    '.OZZZK',
    '.ms-Persona-primaryText',
    '[class*="personaName" i]',
    '[data-testid="senderName"]',
    '[class*="senderName" i]',
    '[class*="sender-name" i]',
    '[class*="fromAddress" i]',
    // Sender contact button — must contain "mailto:" or be inside a persona
    'a[href^="mailto:"]',
    '[class*="Persona" i] button[aria-label]',
  ];
  const emailSelectors = [
    '[data-testid="senderEmail"]',
    '[class*="senderEmail" i]',
    '[class*="sender-email" i]',
    '.ms-Persona-secondaryText',
  ];

  let name = "Unknown";
  let email = "";

  const nearbySenderLabel = semaiGetSenderLabelNearBody(bodyEl);
  if (nearbySenderLabel) {
    const normalized = semaiNormalizeSenderLabel(nearbySenderLabel);
    if (normalized.name.length >= 2 && semaiLooksLikeSenderName(normalized.name)) {
      name = normalized.name;
    }
    if (normalized.email) {
      email = normalized.email;
    }
  }

  let ancestor = bodyEl.parentElement;
  for (let d = 0; d < 10 && ancestor; d++, ancestor = ancestor.parentElement) {
    if (name === "Unknown") {
      const found = semaiPickClosestPrecedingMatch(ancestor, bodyEl, nameSelectors);
      if (found) {
        // Outlook search highlighting strips the matched word from aria-label while
        // innerText still includes the full visible name. Compare both and use whichever
        // normalizes to a more complete (longer) name.
        const ariaLabel = (found.getAttribute("aria-label") || "").trim();
        // semaiFullText uses textContent so highlighted search-term spans (mark.js)
        // are always included, even when CSS excludes them from innerText.
        const fullText = semaiFullText(found);
        let raw;
        if (ariaLabel && fullText) {
          const fromAria = semaiNormalizeSenderLabel(ariaLabel);
          const fromText = semaiNormalizeSenderLabel(fullText);
          raw = fromText.name.length > fromAria.name.length ? fullText : ariaLabel;
        } else {
          raw = ariaLabel || fullText;
        }
        const normalized = semaiNormalizeSenderLabel(raw);
        if (normalized.name.length >= 2 && semaiLooksLikeSenderName(normalized.name)) {
          name = normalized.name;
        }
        if (!email && normalized.email) {
          email = normalized.email;
        }
      }
    }

    if (!email) {
      const found = semaiPickClosestPrecedingMatch(ancestor, bodyEl, emailSelectors);
      if (found) {
        const text = (found.innerText || found.textContent || "").trim();
        if (text.includes("@")) email = text;
      }
    }

    if (name !== "Unknown" && email) break;
  }

  return { name, email, initials: semaiInitials(name) };
}

// Extract timestamp from the message card
function semaiGetMessageTimestamp(bodyEl) {
  let ancestor = bodyEl.parentElement;
  for (let d = 0; d < 10 && ancestor; d++, ancestor = ancestor.parentElement) {
    const timeEl = ancestor.querySelector(
      'time, [data-testid="sentTime"], [class*="DateTimeSent" i], [class*="timestamp" i], [class*="date-time" i]'
    );
    if (timeEl && !timeEl.contains(bodyEl)) {
      return (timeEl.getAttribute("datetime") || timeEl.innerText || timeEl.textContent || "").trim();
    }
  }
  return "";
}

function semaiGetMessageCard(bodyEl) {
  const labeledMessageCard = bodyEl.closest(
    '[aria-label="Email message"], [aria-label*="Email message" i], [role="group"][aria-label*="message" i]'
  );
  if (labeledMessageCard) {
    return labeledMessageCard;
  }

  const bodyContainer = bodyEl.closest('[data-test-id="mailMessageBodyContainer"]');
  if (bodyContainer?.parentElement) {
    return bodyContainer.parentElement;
  }

  let ancestor = bodyEl.parentElement;
  for (let depth = 0; depth < 8 && ancestor; depth += 1, ancestor = ancestor.parentElement) {
    if (ancestor.querySelector('time, [data-testid="sentTime"], a[href^="mailto:"]')) {
      return ancestor;
    }
  }

  return bodyEl.parentElement;
}

function semaiMessageHasAttachmentBlock(bodyEl) {
  const scopes = [];
  const messageCard = semaiGetMessageCard(bodyEl);
  if (messageCard) {
    scopes.push(messageCard);
  }

  const bodyContainer = bodyEl.closest('[data-test-id="mailMessageBodyContainer"]');
  if (bodyContainer) {
    scopes.push(bodyContainer);

    let sibling = bodyContainer.previousElementSibling;
    while (sibling) {
      if (!scopes.includes(sibling)) {
        scopes.push(sibling);
      }
      sibling = sibling.previousElementSibling;
    }

    sibling = bodyContainer.nextElementSibling;
    while (sibling) {
      if (!scopes.includes(sibling)) {
        scopes.push(sibling);
      }
      sibling = sibling.nextElementSibling;
    }
  }

  let ancestor = bodyEl.parentElement;
  for (let depth = 0; depth < 8 && ancestor; depth += 1, ancestor = ancestor.parentElement) {
    if (!scopes.includes(ancestor)) {
      scopes.push(ancestor);
    }
  }

  const attachmentSelectors = [
    '[role="heading"][id$="_ATTACHMENTS"]',
    '[id$="_ATTACHMENTS"]',
    '[role="listbox"][aria-label="file attachments"]',
    '[role="listbox"][aria-label*="attachments" i]',
    '[role="option"][aria-label*=" open " i]',
    '.av-container [role="option"]'
  ];

  const attachmentLabelRe = /\.(pdf|doc|docx|xls|xlsx|ppt|pptx|png|jpg|jpeg|zip|csv|txt)\b.+\bopen\b/i;

  return scopes.some(scope => {
    if (!(scope instanceof Element)) return false;

    return attachmentSelectors.some(selector => {
      return Array.from(scope.querySelectorAll(selector)).some(match => {
        if (!(match instanceof Element) || bodyEl.contains(match)) {
          return false;
        }

        const label = (
          match.getAttribute("aria-label") ||
          match.getAttribute("title") ||
          match.textContent ||
          ""
        ).replace(/\s+/g, " ").trim();

        if (
          match.matches('[role="heading"][id$="_ATTACHMENTS"], [id$="_ATTACHMENTS"]') ||
          match.matches('[role="listbox"][aria-label*="attachments" i]')
        ) {
          return true;
        }

        return attachmentLabelRe.test(label);
      });
    });
  });
}

function semaiIsCurrentUser(senderName, senderEmail) {
  const user = semaiGetCurrentUser();
  if (!user) return false;

  // Email match (strongest signal)
  if (user.email && senderEmail && senderEmail.toLowerCase() === user.email) return true;

  const sLower = senderName.toLowerCase().trim();
  if (!sLower) return false;

  // Exact full name match
  if (sLower === user.nameLower) return true;

  // "Lastname(s), Firstname" format — handles both single and compound last names.
  // e.g. "Alvarez, Santiago" or "Arconada Alvarez, Santiago" (Spanish compound surnames).
  const userParts = user.nameLower.split(/\s+/);
  if (userParts.length >= 2 && sLower.includes(",")) {
    const commaIdx = sLower.indexOf(",");
    const beforeComma = sLower.substring(0, commaIdx).trim();
    const afterComma = sLower.substring(commaIdx + 1).trim();
    const userFirst = userParts[0];
    const userLastWords = userParts.slice(1);
    if (
      afterComma.split(/\s+/).includes(userFirst) &&
      userLastWords.some(w => beforeComma.split(/\s+/).includes(w))
    ) return true;
  }

  return false;
}

// Extract all messages in the thread, in document order (oldest first)
function semaiExtractThreadMessages() {
  const bodies = Array.from(
    document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])')
  );
  return bodies.map(bodyEl => {
    const sender = semaiGetMessageSender(bodyEl);
    const timestamp = semaiGetMessageTimestamp(bodyEl);
    const senderFirstName = semaiFirstNameFromDisplayName(sender.name);
    semaiNativeLog(`[semai-sig] extractThreadMessages: sender.name="${sender.name}" → senderFirstName="${senderFirstName}"`);
    const cleanHtml = semaiCleanBodyClone(bodyEl, senderFirstName);
    const rawHtml = bodyEl.dataset.semaiOriginalHtml || bodyEl.innerHTML;
    const isMe = semaiIsCurrentUser(sender.name, sender.email);
    const hasAttachment = semaiMessageHasAttachmentBlock(bodyEl);
    return { sender, timestamp, cleanHtml, rawHtml, isMe, hasAttachment, sourceBodyEl: bodyEl };
  });
}

// Get the conversation subject
function semaiGetThreadSubject() {
  const readingPane = semaiGetReadingPane();
  const searchRoots = [readingPane, document].filter(Boolean);
  const selectors = [
    '[data-testid="subjectLine"]',
    '[data-testid*="subject" i]',
    '[aria-label="Subject"]',
    '[aria-label^="Subject" i]',
    'h1[class*="subject" i]',
    'h2[class*="subject" i]',
    '[role="heading"][data-testid*="subject" i]',
    '[role="heading"][class*="subject" i]'
  ];

  const normalizeSubject = (value) => (value || "")
    .replace(/^subject:\s*/i, "")
    .replace(/\s+/g, " ")
    .trim();

  const isLikelySubject = (text, el) => {
    const normalized = normalizeSubject(text);
    if (!normalized) return false;
    if (normalized.length < 2) return false;

    const attrText = [
      el?.getAttribute?.("data-testid") || "",
      el?.getAttribute?.("aria-label") || "",
      el?.className || ""
    ].join(" ").toLowerCase();

    if (attrText.includes("subject")) return true;
    if (/^(from|to|cc|bcc):/i.test(normalized)) return false;
    if (/^[^@]+@[^@]+\.[^@]+$/.test(normalized)) return false;

    return normalized.split(/\s+/).length >= 2 || /[—\-:()[\]]/.test(normalized);
  };

  for (const root of searchRoots) {
    for (const sel of selectors) {
      const candidates = Array.from(root.querySelectorAll(sel));
      for (const el of candidates) {
        const text = normalizeSubject(
          el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement
            ? el.value
            : (el.innerText || el.textContent || "")
        );
        if (isLikelySubject(text, el)) {
          return text;
        }
      }
    }
  }

  const title = (document.title || "").trim();
  const titleCandidates = title
    .split(/\s+-\s+/)
    .map((part) => normalizeSubject(part))
    .filter(Boolean)
    .filter((part) => !/^(mail|outlook|inbox)$/i.test(part));

  return titleCandidates[0] || "Conversation";
}

// Build the chat overlay DOM
function semaiCreateChatOverlay(messages, subject) {
  const overlay = document.createElement("div");
  overlay.id = "semai-chat-overlay";
  overlay.dataset.viewMode = "chat";
  overlay.dataset.reportMode = "inactive";
  overlay._semaiMessages = messages;
  overlay._semaiSubject = subject;

  // Header bar
  const header = document.createElement("div");
  header.className = "semai-chat-header";
  header.innerHTML = `
    <span class="semai-chat-subject">${semaiEscapeHtml(subject)}</span>
    <div class="semai-chat-header-actions">
      <div class="semai-chat-settings">
        <button
          class="semai-chat-settings-toggle"
          type="button"
          aria-label="Open chat settings"
          aria-expanded="false"
          title="Chat settings"
        >⚙</button>
        <div class="semai-chat-settings-popover" hidden>
          <div class="semai-chat-settings-title">Settings</div>
          <label class="semai-chat-settings-control" for="semai-chat-font-size-slider">
            <span>Font size</span>
            <span class="semai-chat-settings-value">${semaiChatUiSettings.fontSize}px</span>
          </label>
          <input
            id="semai-chat-font-size-slider"
            class="semai-chat-settings-slider"
            type="range"
            min="${SEMAI_CHAT_FONT_SIZE_MIN}"
            max="${SEMAI_CHAT_FONT_SIZE_MAX}"
            step="1"
            value="${semaiChatUiSettings.fontSize}"
          />
        </div>
      </div>
      <button class="semai-chat-close" type="button">✕ Hide chat view</button>
    </div>
  `;
  header.querySelector(".semai-chat-close").addEventListener("click", semaiDeactivateChatView);
  overlay.appendChild(header);
  semaiApplyChatUiSettings(overlay);

  const settingsToggle = header.querySelector(".semai-chat-settings-toggle");
  const settingsPopover = header.querySelector(".semai-chat-settings-popover");
  const settingsSlider = header.querySelector(".semai-chat-settings-slider");

  settingsToggle?.addEventListener("click", (event) => {
    event.stopPropagation();
    semaiToggleChatSettingsPopover(overlay);
  });

  settingsPopover?.addEventListener("click", (event) => {
    event.stopPropagation();
  });

  settingsSlider?.addEventListener("input", (event) => {
    const target = event.target;
    if (!(target instanceof HTMLInputElement)) return;
    semaiUpdateChatUiSettings({ fontSize: target.value });
  });

  overlay.addEventListener("click", (event) => {
    const target = event.target;
    if (!(target instanceof Element)) return;
    if (target.closest(".semai-chat-settings")) return;
    semaiCloseChatSettingsPopover(overlay);
  });

  const content = document.createElement("div");
  content.className = "semai-chat-content";

  // Scrollable chat message area
  const chatScroll = document.createElement("div");
  chatScroll.className = "semai-chat-scroll";

  messages.forEach((msg, index) => {
    if (!msg.cleanHtml) return; // skip empty bodies

    const row = document.createElement("div");
    row.className = `semai-chat-row ${msg.isMe ? "semai-chat-me" : "semai-chat-them"}`;
    row.dataset.reportIndex = String(index);

    const avatar = document.createElement("div");
    avatar.className = "semai-chat-avatar";
    avatar.textContent = msg.sender.initials;
    avatar.style.background = semaiNameColor(msg.sender.name);
    avatar.title = msg.sender.name;

    const bubble = document.createElement("div");
    bubble.className = "semai-chat-bubble";

    // Body
    const body = document.createElement("div");
    body.className = "semai-chat-body";
    body.innerHTML = msg.cleanHtml;
    bubble.appendChild(body);

    const attachmentBadge = msg.hasAttachment
      ? Object.assign(document.createElement("span"), {
          className: "semai-chat-attachment-badge",
          textContent: "📎"
        })
      : null;

    if (msg.isMe) {
      if (attachmentBadge) {
        row.appendChild(attachmentBadge);
      }
      row.appendChild(bubble);
      row.appendChild(avatar);
    } else {
      row.appendChild(avatar);
      row.appendChild(bubble);
      if (attachmentBadge) {
        row.appendChild(attachmentBadge);
      }
    }
    chatScroll.appendChild(row);
  });

  const chatEndSpacer = document.createElement("div");
  chatEndSpacer.className = "semai-chat-end-spacer";
  chatScroll.appendChild(chatEndSpacer);

  content.appendChild(chatScroll);
  overlay.appendChild(content);

  const composer = document.createElement("div");
  composer.className = "semai-chat-composer";
  composer.innerHTML = `
    <textarea
      id="semai-chat-reply-input"
      class="semai-chat-reply-input"
      rows="2"
      placeholder="Type a reply-all response to the latest message…"
    ></textarea>
    <div class="semai-chat-composer-footer">
      <div id="semai-chat-reply-status" class="semai-chat-reply-status">
        Chat view is on. Use the eye button to switch back to regular Outlook.
      </div>
      <pre id="semai-chat-reply-debug" class="semai-chat-reply-debug"></pre>
      <div class="semai-chat-reply-actions">
        <button
          id="semai-chat-report-issue-btn"
          class="semai-chat-report-issue-btn"
          type="button"
        >
          Report issue
        </button>
        <button
          id="semai-chat-view-toggle-btn"
          class="semai-chat-view-toggle-btn"
          type="button"
          aria-label="Show real thread above the reply box"
          title="Show real thread above the reply box"
        >
          <span class="semai-chat-view-toggle-icon" aria-hidden="true">
            <svg viewBox="0 0 24 24" focusable="false">
              <path class="semai-eye-shape" d="M12 5C6.5 5 2.1 8.4 1 12c1.1 3.6 5.5 7 11 7s9.9-3.4 11-7c-1.1-3.6-5.5-7-11-7Zm0 11.2A4.2 4.2 0 1 1 12 7.8a4.2 4.2 0 0 1 0 8.4Zm0-2.1a2.1 2.1 0 1 0 0-4.2 2.1 2.1 0 0 0 0 4.2Z"></path>
              <path class="semai-eye-slash" d="M5.1 3.7 20.3 18.9l-1.4 1.4L3.7 5.1l1.4-1.4Z"></path>
            </svg>
          </span>
        </button>
        <button id="semai-chat-reply-send-btn" class="semai-chat-reply-btn" type="button">
          Reply all
        </button>
      </div>
    </div>
  `;

  const replyInput = composer.querySelector("#semai-chat-reply-input");
  const reportIssueBtn = composer.querySelector("#semai-chat-report-issue-btn");
  const viewToggleBtn = composer.querySelector("#semai-chat-view-toggle-btn");
  const replyBtn = composer.querySelector("#semai-chat-reply-send-btn");

  reportIssueBtn.addEventListener("click", () => {
    semaiToggleReportMode(overlay);
  });
  viewToggleBtn.addEventListener("click", () => {
    semaiToggleOverlayView(overlay);
  });
  replyBtn.addEventListener("click", semaiSendReplyAllFromChat);
  replyInput.addEventListener("input", () => {
    if (overlay._semaiReadingPane) {
      requestAnimationFrame(() => semaiUpdateReadingPaneBottomClearance(overlay._semaiReadingPane, overlay));
    }
  });
  replyInput.addEventListener("keydown", (event) => {
    if (event.key === "Enter" && (event.metaKey || event.ctrlKey)) {
      event.preventDefault();
      semaiSendReplyAllFromChat();
    }
  });
  chatScroll.addEventListener("mousemove", semaiHandleReportRowHover);
  chatScroll.addEventListener("mouseleave", () => {
    semaiClearReportHover();
  });
  chatScroll.addEventListener("click", semaiHandleReportRowClick, true);

  overlay.appendChild(composer);
  semaiUpdateOverlayViewToggle(overlay);
  return overlay;
}

function semaiUpdateOverlayViewToggle(overlay) {
  const toggleBtn = overlay.querySelector("#semai-chat-view-toggle-btn");
  const status = overlay.querySelector("#semai-chat-reply-status");
  if (!toggleBtn) return;

  const isChatView = overlay.dataset.viewMode !== "real";
  toggleBtn.classList.toggle("semai-chat-view-toggle-active", !isChatView);
  toggleBtn.classList.toggle("semai-chat-view-toggle-chat", isChatView);
  toggleBtn.setAttribute("aria-label", isChatView ? "Show real thread above the reply box" : "Show chat thread above the reply box");
  toggleBtn.setAttribute("title", isChatView ? "Show real thread above the reply box" : "Show chat thread above the reply box");

  if (status && overlay.dataset.reportMode !== "active") {
    status.textContent = isChatView
      ? "Chat view is on. Use the eye button to switch back to regular Outlook."
      : "The original Outlook thread is visible above the reply box. Use the eye button to switch back to chat.";
    delete status.dataset.tone;
  }

  if (overlay._semaiReadingPane) {
    semaiUpdateReadingPaneBottomClearance(overlay._semaiReadingPane, overlay);
  }
}

function semaiToggleOverlayView(overlay) {
  if (!overlay) return;

  semaiCloseChatSettingsPopover(overlay);
  overlay.dataset.viewMode = overlay.dataset.viewMode === "real" ? "chat" : "real";
  semaiUpdateOverlayViewToggle(overlay);
}

function semaiUpdateReadingPaneBottomClearance(readingPane, overlay) {
  if (!readingPane || !overlay) return;

  let spacer = readingPane.querySelector(":scope > .semai-reading-pane-spacer");
  if (!spacer) {
    spacer = document.createElement("div");
    spacer.className = "semai-reading-pane-spacer";
    readingPane.appendChild(spacer);
  }

  const composer = overlay.querySelector(".semai-chat-composer");
  const header = overlay.querySelector(".semai-chat-header");
  const composerHeight = composer?.getBoundingClientRect().height || 0;
  const headerHeight = header?.getBoundingClientRect().height || 0;
  const isRealView = overlay.dataset.viewMode === "real";
  if (isRealView) {
    spacer.style.height = "20px";
    return;
  }

  const extraClearance = 320;
  spacer.style.height = `${Math.ceil(composerHeight + headerHeight + extraClearance)}px`;
}

function semaiRemoveReadingPaneBottomClearance(readingPane) {
  readingPane?.querySelector(":scope > .semai-reading-pane-spacer")?.remove();
}

function semaiIsLargePaneCandidate(el) {
  if (!(el instanceof Element)) return false;
  const rect = el.getBoundingClientRect();
  return rect.width >= 420 && rect.height >= 320;
}

// Find the Outlook reading pane container and prefer the broadest reader surface.
function semaiGetReadingPane() {
  const firstBody = document.querySelector('[aria-label="Message body"]:not([contenteditable])');
  if (!firstBody) return null;

  const priorityPane =
    firstBody.closest('[data-app-section="MailReadCompose"]') ||
    firstBody.closest('[data-app-section="ConversationContainer"]') ||
    firstBody.closest('[role="main"]') ||
    firstBody.closest('[role="complementary"]') ||
    firstBody.closest('[data-app-section-name]');

  if (priorityPane && semaiIsLargePaneCandidate(priorityPane)) {
    return priorityPane;
  }

  let bestPane = null;
  let el = firstBody.parentElement;
  for (let d = 0; d < 25 && el; d++, el = el.parentElement) {
    if (el === document.body || el === document.documentElement) break;
    if (!semaiIsLargePaneCandidate(el)) continue;

    bestPane = el;
    const style = window.getComputedStyle(el);
    const overflow = style.overflowY;
    if ((overflow === "auto" || overflow === "scroll") && el.clientHeight > 200) {
      return bestPane;
    }
  }

  return bestPane || null;
}

function semaiActivateChatView() {
  if (semaiChatViewActive || semaiChatViewActivationInProgress) return;
  if (!semaiGetCalibration()?.senderSelector) {
    semaiUpdateCalibrationStatus(
      "Train sender detection before turning on chat view.",
      "neutral"
    );
    return;
  }
  semaiChatViewActivationInProgress = true;
  semaiChatViewPinned = true;

  try {
    // Ensure we have the current user
    if (!semaiGetCurrentUser()) {
      alert("Remou couldn't identify your account.\nComplete setup first by clicking 'Train sender detection' in the Remou panel.");
      return;
    }

    const messages = semaiExtractThreadMessages();
    if (messages.length < 2) {
      alert("semai chat view needs a thread with at least 2 messages.");
      return;
    }

    const subject = semaiGetThreadSubject();
    const overlay = semaiCreateChatOverlay(messages, subject);

    semaiChatViewActive = true;
    semaiUpdateChatToggleBtn();
    semaiTrackEvent("chat_on", {
      page_url: window.location.href,
      message_count: messages.length
    });

    // Contain the overlay within the reading pane
    const readingPane = semaiGetReadingPane();
    if (readingPane) {
      const rpStyle = window.getComputedStyle(readingPane);
      if (rpStyle.position === "static") readingPane.style.position = "relative";
      readingPane.appendChild(overlay);
      overlay._semaiReadingPane = readingPane;
      semaiUpdateReadingPaneBottomClearance(readingPane, overlay);
      if (window.ResizeObserver) {
        const resizeObserver = new ResizeObserver(() => {
          semaiUpdateReadingPaneBottomClearance(readingPane, overlay);
        });
        const composer = overlay.querySelector(".semai-chat-composer");
        const header = overlay.querySelector(".semai-chat-header");
        if (composer) resizeObserver.observe(composer);
        if (header) resizeObserver.observe(header);
        overlay._semaiResizeObserver = resizeObserver;
      }
    } else {
      overlay.classList.add("semai-chat-overlay-fixed");
      document.body.appendChild(overlay);
    }

    // Scroll to bottom after the overlay is in the DOM and painted
    const scrollEl = overlay.querySelector(".semai-chat-scroll");
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        scrollEl.scrollTop = scrollEl.scrollHeight;
      });
    });
 
    semaiLog("[semai] Chat view activated", { messageCount: messages.length });
  } finally {
    semaiChatViewActivationInProgress = false;
  }
}

function semaiDeactivateChatView(preservePinned = false) {
  const overlay = document.getElementById("semai-chat-overlay");
  const readingPane = overlay?.parentElement;
  document.removeEventListener("keydown", semaiHandleReportModeKeydown, true);
  semaiReportModeOverlay = null;
  semaiCloseReportPopover();
  semaiClearReportHover();
  overlay?._semaiResizeObserver?.disconnect();
  if (overlay) overlay.remove();
  semaiRemoveReadingPaneBottomClearance(readingPane);
  semaiChatViewActive = false;
  if (!preservePinned) {
    semaiChatViewPinned = false;
  }
  const bodies = document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])');
  semaiAutoOpenSuppressedSignature = Array.from(bodies).map(b => b.dataset.semaiSigStripped || "").join("|");
  semaiUpdateChatToggleBtn();
  semaiTrackEvent("chat_off", {
    page_url: window.location.href
  });
  semaiLog("[semai] Chat view deactivated");
}

function semaiUpdateChatToggleBtn() {
  const btn = document.querySelector(".semai-chat-toggle-btn");
  if (!btn) return;
  const isCalibrated = Boolean(semaiGetCalibration()?.senderSelector);
  btn.textContent = semaiChatViewActive ? "Hide chat view" : "Turn on chat view";
  btn.disabled = !semaiChatViewActive && !isCalibrated;
  btn.title = !semaiChatViewActive && !isCalibrated
    ? "Train sender detection before turning on chat view."
    : "";
}

// Show/hide the chat toggle based on whether we're looking at a thread
function semaiUpdateChatToggleVisibility() {
  const btn = document.querySelector(".semai-chat-toggle-btn");
  if (!btn) return;
  const bodies = document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])');
  btn.style.display = bodies.length >= 2 ? "" : "none";
  semaiUpdateChatToggleBtn();
}

// Auto-deactivate when Outlook navigates to a different email
let semaiLastReadingPaneSignature = "";
function semaiWatchForNavigation() {
  let checkTimer = null;

  const check = () => {
    checkTimer = null;
    const bodies = document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])');
    const sig = Array.from(bodies).map(b => b.dataset.semaiSigStripped || "").join("|");
    if (semaiChatViewActive && sig !== semaiLastReadingPaneSignature) {
      semaiDeactivateChatView(true);
    }
    if (sig !== semaiLastReadingPaneSignature) {
      semaiAutoOpenSuppressedSignature = "";
    }
    semaiLastReadingPaneSignature = sig;
    semaiUpdateChatToggleVisibility();
    if (
      semaiChatViewPinned &&
      !semaiChatViewActive &&
      !semaiChatViewActivationInProgress &&
      !semaiCalibrationState &&
      semaiGetCalibration()?.senderSelector &&
      bodies.length >= 2 &&
      sig
    ) {
      semaiActivateChatView();
    }
  };

  const scheduleCheck = () => {
    if (checkTimer) {
      window.clearTimeout(checkTimer);
    }
    checkTimer = window.setTimeout(check, 180);
  };

  const obs = new MutationObserver(scheduleCheck);
  obs.observe(document.body, { childList: true, subtree: true });
  check();
}

function semaiObserveReadingBodies() {
  const selector = '[aria-label="Message body"]:not([contenteditable])';

  const observer = new MutationObserver(() => {
    document.querySelectorAll(selector).forEach(body => {
      semaiStripSignature(body);
    });
  });

  observer.observe(document.body, { childList: true, subtree: true });

  // Handle already-rendered emails on load
  document.querySelectorAll(selector).forEach(body => semaiStripSignature(body));
}

// ===== INIT =====
function setupWhenReady() {
  if (!document.body || !document.documentElement) {
    window.setTimeout(setupWhenReady, 150);
    return;
  }

  createPanel();
  semaiObserveReadingBodies();
  semaiGetCurrentUser();
  semaiWatchForNavigation();
  document.addEventListener("mousemove", semaiHandleOriginalReportHover, true);
  document.addEventListener("click", semaiHandleOriginalReportClick, true);
  document.addEventListener("selectionchange", semaiSaveSelectionFromCompose);
  window.addEventListener("resize", () => {
    const panel = document.getElementById("semai-panel");
    if (panel) semaiEnsurePanelVisible(panel, false);
  });

  // First-run modal is shown by createPanel() when calibration is missing.
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", setupWhenReady, { once: true });
} else {
  setupWhenReady();
}
