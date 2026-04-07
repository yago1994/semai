/**
 * Signature cleanup logic test — Node.js / jsdom
 *
 * Tests semaiCleanBodyClone against two real-world HTML cases:
 *   1. Daniel Drane / Emory contact card
 *   2. Mandi Schmitt / UserTesting outreach signature
 *
 * Run with:
 *   node docs/test-signatures.js
 */

import { JSDOM } from "jsdom";

// ─── Shim the DOM environment ───────────────────────────────────────────────
// We need a shared `document` / `window` so the functions that reference them
// (getBoundingClientRect etc.) don't crash.  We create one minimal global DOM
// and then re-use it across tests.
const { window: globalWindow } = new JSDOM("<!DOCTYPE html><html><body></body></html>");
const { document: globalDoc } = globalWindow;

// Patch globals so the ported functions can call document.createElement etc.
global.window     = globalWindow;
global.document   = globalDoc;
global.Element    = globalWindow.Element;
global.HTMLElement = globalWindow.HTMLElement;
global.NodeFilter  = globalWindow.NodeFilter;

// ─── Paste in all the constants and helper functions from contentScript.js ──
// (Minimal subset — only what semaiCleanBodyClone and its callees need.)

const SEMAI_DEBUG = false;
const SEMAI_SIG_SHORT_LINE_MAX = 60;
const SEMAI_SIG_MIN_LINES      = 5;
const SEMAI_PHONE_RE    = /\+?[\d][\d\s\-\.\(\)]{6,}/;
const SEMAI_EMAIL_RE    = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i;
const SEMAI_URL_RE      = /https?:\/\/|www\./i;
const SEMAI_SOCIAL_RE   = /linkedin\.com|twitter\.com|facebook\.com|instagram\.com/i;
const SEMAI_SIGNATURE_TITLE_RE   = /\b(professor|director|chair|vice chair|clinical research|neurology|pediatrics|medicine|program|department|affiliate professor|research|health)\b/i;
const SEMAI_SIGNATURE_ADDRESS_RE = /\b(suite|mailstop|building|circle|road|street|avenue|drive|lane|boulevard|floor|room|atlanta|georgia|\d{5}(?:-\d{4})?)\b/i;

const SEMAI_BLOCK_TAGS = new Set([
  "P","DIV","BLOCKQUOTE","LI","H1","H2","H3","H4","H5","H6",
  "SECTION","ARTICLE","HEADER","FOOTER","MAIN","TABLE","TR","TD","TH"
]);

const SEMAI_CLOSING_RE = /^(best|regards|thanks|thank you|cheers|sincerely|warmly|warm regards|kind regards|best regards|yours|cordially|take care|many thanks|with gratitude|respectfully|yours truly|talk soon|looking forward)[,.]?\s*$/i;

function semaiLog() {}  // no-op in tests

// ── Walk text nodes to find the nearest block-level element matching pattern ──
function semaiFirstSeparatorEl(container, pattern) {
  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
  let node;
  while ((node = walker.nextNode())) {
    if (pattern.test(node.textContent.trim())) {
      let el = node.parentElement;
      while (el && el !== container && !SEMAI_BLOCK_TAGS.has(el.tagName)) {
        el = el.parentElement;
      }
      if (el && el !== container) return el;
    }
  }
  return null;
}

// ── Strip quoted-reply headers ────────────────────────────────────────────────
function semaiStripQuotedReplyHeaders(container) {
  const HEADER_LINE_RE = /^(from|date|sent|to|cc|subject)\s*:/i;
  const HEADER_PAIR_RE = /(from|date|sent|to|cc|subject)\s*:.*\n.*(from|date|sent|to|cc|subject)\s*:/is;
  const WROTE_LINE_RE  = /^on .+wrote:\s*$/i;

  function removeFromAndAfter(el) {
    let sib = el;
    while (sib) { const next = sib.nextElementSibling; sib.remove(); sib = next; }
  }

  const blocks = Array.from(container.querySelectorAll("div, p, blockquote, section, article, td"));
  for (const block of blocks) {
    const blockText = (block.textContent || "").trim();
    if (!blockText) continue;
    const lines = blockText.split(/\n+/).map(l => l.trim()).filter(Boolean);
    const headerLines = lines.filter(l => HEADER_LINE_RE.test(l));
    if (headerLines.length >= 2 || HEADER_PAIR_RE.test(blockText)) { removeFromAndAfter(block); return; }
    if (WROTE_LINE_RE.test(blockText)) { removeFromAndAfter(block); return; }
  }

  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
  let textNode;
  while ((textNode = walker.nextNode())) {
    const text = (textNode.textContent || "").trim();
    if (!HEADER_LINE_RE.test(text) && !WROTE_LINE_RE.test(text)) continue;
    let block = textNode.parentElement;
    while (block && block !== container && !SEMAI_BLOCK_TAGS.has(block.tagName)) block = block.parentElement;
    if (!block || block === container) continue;
    const blockText = (block.textContent || "").trim();
    const lines = blockText.split(/\n+/).map(l => l.trim()).filter(Boolean);
    const headerLines = lines.filter(l => HEADER_LINE_RE.test(l));
    if (headerLines.length >= 1 || lines.some(l => WROTE_LINE_RE.test(l))) { removeFromAndAfter(block); return; }
  }

  const quoteBlocks = Array.from(container.querySelectorAll("blockquote"));
  for (const quoteBlock of quoteBlocks) {
    const previous = quoteBlock.previousElementSibling;
    const previousText = (previous?.textContent || "").trim();
    if (previous && WROTE_LINE_RE.test(previousText)) { previous.remove(); quoteBlock.remove(); return; }
    const parentText = (quoteBlock.parentElement?.textContent || "").trim();
    if (WROTE_LINE_RE.test(parentText.split(/\n+/)[0] || "")) { quoteBlock.parentElement?.remove(); return; }
  }
}

// ── Looks-like-sig heuristic ──────────────────────────────────────────────────
function semaiLooksLikeSig(el) {
  const text = el.textContent || "";
  const lines = text.split("\n").map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length < SEMAI_SIG_MIN_LINES) return false;
  const shortLines = lines.filter(l => l.length <= SEMAI_SIG_SHORT_LINE_MAX).length;
  if (shortLines / lines.length < 0.8) return false;
  return SEMAI_PHONE_RE.test(text) || SEMAI_URL_RE.test(text) || SEMAI_SOCIAL_RE.test(text);
}

// ── Nested-div ltr signature detection ───────────────────────────────────────
function semaiLooksLikeNestedDivSignature(el) {
  if (!(el instanceof HTMLElement)) return false;
  if (el.tagName !== "DIV" || el.getAttribute("dir")?.toLowerCase() !== "ltr") return false;
  const childDivs = Array.from(el.children).filter(c => c.tagName === "DIV");
  if (childDivs.length < 4 || childDivs.length !== el.children.length) return false;
  const lines = (el.textContent || "").split(/\r?\n|\r/).map(l => l.trim()).filter(Boolean);
  if (lines.length < 6) return false;
  const hasPhone    = lines.some(l => SEMAI_PHONE_RE.test(l));
  const hasEmailOrUrl = lines.some(l => /@/.test(l) || /\b(?:[a-z0-9-]+\.)+[a-z]{2,}(?:\/[^\s]*)?\b/i.test(l));
  const hasOrgLine  = lines.some(l => /\b(university|school|medicine|program|director|professor|chair|clinical research|department|health)\b/i.test(l));
  return hasPhone && hasEmailOrUrl && hasOrgLine;
}

function semaiFindNestedDivSignature(container) {
  const candidates = Array.from(container.querySelectorAll('div[dir="ltr"]'));
  for (let i = candidates.length - 1; i >= 0; i--) {
    if (semaiLooksLikeNestedDivSignature(candidates[i])) return candidates[i];
  }
  return null;
}

// ── Name helpers ──────────────────────────────────────────────────────────────
function semaiNormalizeNameLine(text) {
  return (text || "").replace(/[^\p{L}\s.'-]/gu, " ").replace(/\s+/g, " ").trim().toLowerCase();
}

function semaiGetNameTokens(text) {
  return semaiNormalizeNameLine(text).split(/\s+/).filter(t => t.length >= 2);
}

function semaiLooksLikeNameLine(text, senderFirstName) {
  const raw = (text || "").trim();
  if (!raw || raw.length > 60) return false;
  const tokens = raw.split(/\s+/).filter(Boolean);
  if (tokens.length === 0 || tokens.length > 5) return false;
  const nameish = tokens.every(t => /^[A-Z][A-Za-z.'-]*$/.test(t) || /^[A-Z]\.$/.test(t));
  if (nameish) return true;
  return !!senderFirstName && raw.toLowerCase() === senderFirstName;
}

function semaiCountContactSignals(text, containerEl) {
  const signalKinds = new Set();
  // Build the full scan string: visible text + any href values from <a> tags
  let raw = text || "";
  if (containerEl && typeof containerEl.querySelectorAll === "function") {
    const hrefs = Array.from(containerEl.querySelectorAll("a"))
      .map(a => a.getAttribute("href") || "")
      .filter(Boolean)
      .join(" ");
    if (hrefs) raw = raw + " " + hrefs;
  }
  if (SEMAI_PHONE_RE.test(raw))                           signalKinds.add("phone");
  if (SEMAI_EMAIL_RE.test(raw))                           signalKinds.add("email");
  if (SEMAI_URL_RE.test(raw) || SEMAI_SOCIAL_RE.test(raw)) signalKinds.add("url");
  if (SEMAI_SIGNATURE_TITLE_RE.test(raw))                 signalKinds.add("title");
  if (SEMAI_SIGNATURE_ADDRESS_RE.test(raw))               signalKinds.add("address");
  return signalKinds.size;
}

function semaiLooksLikeCompactBlock(el) {
  const text = (el?.textContent || "").trim();
  if (!text) return false;
  const lines = text.split(/\n+/).map(l => l.trim()).filter(Boolean);
  if (lines.length > 18) return false;
  return text.length <= 600;
}

function semaiLooksLikeContactCardBlock(el) {
  if (!(el instanceof HTMLElement)) return false;
  const text = (el.textContent || "").trim();
  if (!text) return false;
  const lines = text.split(/\n+/).map(l => l.trim()).filter(Boolean);
  const shortLines = lines.filter(l => l.length <= 90).length;
  const compactRatio = lines.length > 0 ? shortLines / lines.length : 0;
  const signalCount = semaiCountContactSignals(text, el);
  const hasTableOrRichLinks =
    el.tagName === "TABLE" ||
    !!el.querySelector("table, img, a[href^='mailto:'], a[href*='linkedin'], a[href*='calendar'], a[href*='opt_out']");
  return (
    (signalCount >= 2 && lines.length >= 3 && compactRatio >= 0.6) ||
    (signalCount >= 1 && hasTableOrRichLinks)
  );
}

// ── Repeated-name anchor detection ───────────────────────────────────────────
// Detect when the entire body is a compact branded signature with no actual message.
// Returns true if the content starts with the sender's name and has URL/image signals
// but no long-sentence body text — i.e., the whole email IS the signature.
function semaiIsEntireBodySignature(clone, senderFirstName) {
  if (!senderFirstName) return false;

  // Normalize to array of tokens for matching
  const nameTokens = Array.isArray(senderFirstName)
    ? senderFirstName
    : [senderFirstName.toLowerCase()];
  if (nameTokens.length === 0) return false;

  const text = (clone.textContent || "").replace(/\s+/g, " ").trim();
  if (!text) return false;

  // Must be short overall — a real message body would be longer
  if (text.length > 800) return false;

  // Sender's first name must appear within the first three non-empty lines.
  // Checking only the first line is too strict when a company name or logo
  // text precedes the person's name (e.g. "Acme Corp\nLeah Ekube\nRenewal…").
  const lines = text.split(/(?<=[.!?])\s+|\n+/).map(l => l.trim()).filter(Boolean);
  if (lines.length === 0) return false;
  const firstThreeText = lines.slice(0, 3).join(" ").toLowerCase();
  const sigNameTokens = Array.isArray(senderFirstName) ? senderFirstName : [senderFirstName];
  if (!sigNameTokens.some(t => firstThreeText.includes(t.toLowerCase()))) return false;

  // Must have at least one URL/image signal in the raw HTML
  const hasUrl = SEMAI_URL_RE.test(text) || !!clone.querySelector("img, a[href]");
  if (!hasUrl) return false;

  // Must NOT have any long sentence (real message body would have sentences > 60 chars)
  const hasLongSentence = lines.some(l => l.length > 60 && /\s/.test(l));
  if (hasLongSentence) return false;

  return true;
}

function semaiFindRepeatedNameSignatureAnchor(container, senderFirstNameOrTokens) {
  // Accept either a single string (backward compat) or an array of tokens
  const nameTokens = Array.isArray(senderFirstNameOrTokens)
    ? senderFirstNameOrTokens
    : (senderFirstNameOrTokens ? [senderFirstNameOrTokens.toLowerCase()] : []);
  const senderFirstName = nameTokens.length > 0 ? nameTokens[0] : null;

  const blocks = Array.from(container.querySelectorAll("p, div, table, td"))
    .filter(el => semaiLooksLikeCompactBlock(el));
  if (blocks.length < 3) return null;

  const startAt = Math.max(0, blocks.length - 24);
  for (let i = startAt; i < blocks.length; i++) {
    const currentText = (blocks[i].textContent || "").trim();
    // Check if name line matches ANY token
    const matchesAnyToken = nameTokens.length > 0
      ? nameTokens.some(token => semaiLooksLikeNameLine(currentText, token))
      : semaiLooksLikeNameLine(currentText, null);
    if (!matchesAnyToken) continue;

    const currentTokens = semaiGetNameTokens(currentText);
    if (currentTokens.length === 0) continue;

    for (let j = i + 1; j < Math.min(blocks.length, i + 8); j++) {
      const nextEl   = blocks[j];
      const nextText = (nextEl.textContent || "").trim();
      if (!nextText) continue;

      const nextTokens   = semaiGetNameTokens(nextText);
      const repeatedName = currentTokens.some(t => nextTokens.includes(t));
      if (!repeatedName) continue;

      let signalText     = nextText;
      let signalEls      = [nextEl];
      let contactCardSeen = semaiLooksLikeContactCardBlock(nextEl);
      for (let k = j + 1; k < Math.min(blocks.length, j + 8); k++) {
        const blockText = (blocks[k].textContent || "").trim();
        if (!blockText) continue;
        signalText += "\n" + blockText;
        signalEls.push(blocks[k]);
        if (semaiLooksLikeContactCardBlock(blocks[k])) contactCardSeen = true;
      }

      // Build a temporary wrapper so semaiCountContactSignals can scan hrefs too
      const signalWrapper = document.createElement("div");
      signalEls.forEach(e => signalWrapper.appendChild(e.cloneNode(true)));

      if (semaiCountContactSignals(signalText, signalWrapper) >= 2 && contactCardSeen) {
        // Return the ancestor of nextEl whose parent is also an ancestor of blocks[i],
        // so that removeFromAndAfter cuts at the right level (removing the signature
        // block and everything after it, not just the name line inside the email body).
        const nameLineAncestors = new Set();
        let anc = blocks[i];
        while (anc && anc !== container) { nameLineAncestors.add(anc); anc = anc.parentElement; }

        let cutEl = nextEl;
        let walker = nextEl;
        while (walker && walker !== container) {
          if (nameLineAncestors.has(walker.parentElement)) { cutEl = walker; break; }
          walker = walker.parentElement;
        }
        return cutEl;
      }
    }
  }
  return null;
}

// ── Standalone contact-card detection (no preceding sign-off required) ────────
const SEMAI_CREDENTIALS_RE = /\b(ph\.?d|m\.?d|m\.?s|m\.?a|m\.?b\.?a|j\.?d|ed\.?d|d\.?o|r\.?n|b\.?s|b\.?a|abpp|faes|faan|fana|facs|frcp|lcsw|lpc|lmft|cpa|esq|pe)\b/i;

function semaiFindStandaloneContactCard(container, senderFirstNameOrTokens) {
  // Accept either a single string (backward compat) or an array of tokens
  const nameTokens = Array.isArray(senderFirstNameOrTokens)
    ? senderFirstNameOrTokens
    : (senderFirstNameOrTokens ? [senderFirstNameOrTokens.toLowerCase()] : []);
  if (nameTokens.length === 0) return null;

  const blocks = Array.from(container.querySelectorAll("p, div, table, td"))
    .filter(el => {
      const text = (el.textContent || "").trim();
      return text && text.length <= 600;
    });

  if (blocks.length < 2) return null;

  for (let i = 0; i < blocks.length; i++) {
    const el = blocks[i];
    const text = (el.textContent || "").trim();
    if (!text) continue;

    const lines = text.split(/\r?\n|\r/).map(l => l.trim()).filter(Boolean);
    if (lines.length === 0) continue;

    const firstLine = lines[0];

    // First line must start with ANY sender name token (case insensitive)
    const firstLineLower = firstLine.toLowerCase();
    const matchedToken = nameTokens.find(token => firstLineLower.startsWith(token.toLowerCase()));
    if (!matchedToken) continue;
    const charAfter = firstLine[matchedToken.length];
    if (charAfter && !/[\s,.]/.test(charAfter)) continue;

    const hasCredentials = SEMAI_CREDENTIALS_RE.test(firstLine);
    const hasTitle = lines.length >= 2 && SEMAI_SIGNATURE_TITLE_RE.test(lines.slice(0, 3).join(" "));
    if (!hasCredentials && !hasTitle) continue;

    let signalText = text;
    let signalEls = [el];
    for (let k = i + 1; k < Math.min(blocks.length, i + 10); k++) {
      const sibText = (blocks[k].textContent || "").trim();
      if (!sibText) continue;
      if (!container.contains(blocks[k])) break;
      signalText += "\n" + sibText;
      signalEls.push(blocks[k]);
    }

    const signalWrapper = document.createElement("div");
    signalEls.forEach(e => signalWrapper.appendChild(e.cloneNode(true)));

    const signalCount = semaiCountContactSignals(signalText, signalWrapper);
    if (signalCount < 2) continue;

    let cutEl = el;
    let walker = el;
    while (walker && walker !== container) {
      const parent = walker.parentElement;
      if (!parent || parent === container) break;
      const siblings = Array.from(parent.children);
      const hasBodyTextSibling = siblings.some(sib => {
        if (sib === walker) return false;
        const sibText = (sib.textContent || "").trim();
        return sibText.length > 0 && sibText !== "\u00a0";
      });
      if (hasBodyTextSibling) { cutEl = walker; break; }
      walker = parent;
    }

    return cutEl;
  }

  return null;
}

// ── Mobile auto-signature patterns ───────────────────────────────────────────
const SEMAI_MOBILE_SIG_RE = /^(sent from my (iphone|ipad|android)|sent from mobile|sent from outlook for (ios|android)|get outlook for (ios|android))$/i;

function semaiIsMobileSigEl(el) {
  const text = (el.textContent || "").trim();
  return SEMAI_MOBILE_SIG_RE.test(text);
}

// ── Primary sender-name anchor ────────────────────────────────────────────────
function semaiFindSenderAnchor(body, senderName) {
  let scope = body;
  while (scope.children.length === 1 && SEMAI_BLOCK_TAGS.has(scope.children[0].tagName)) {
    scope = scope.children[0];
  }
  const kids   = Array.from(scope.children);
  const startAt = Math.max(0, kids.length - 8);

  for (let i = kids.length - 1; i >= startAt; i--) {
    const raw   = (kids[i].textContent || "").trim();
    if (!raw) continue;
    const lines = raw.split(/\r?\n|\r/).map(l => l.trim()).filter(Boolean);
    const firstLine = lines[0] || "";

    // Pattern A: closing + name in one element via <br>
    if (lines.length === 2 && SEMAI_CLOSING_RE.test(lines[0])) {
      const nameWords  = lines[1].split(/\s+/);
      const nameOk     = nameWords.length <= 3 && nameWords.every(w => /^[A-Z]/.test(w));
      const nameMatches = !senderName || lines[1].toLowerCase().startsWith(senderName);
      if (nameOk && nameMatches && i + 1 < kids.length) return kids[i + 1];
    }

    // Pattern B: name element immediately preceded by closing sibling
    if (lines.length === 1) {
      const words       = raw.split(/\s+/);
      const isShortName = words.length >= 1 && words.length <= 3 && raw.length <= 30 && words.every(w => /^[A-Z]/.test(w));
      const nameMatches = !senderName || raw.toLowerCase().startsWith(senderName);
      if (isShortName && nameMatches && i > 0) {
        const prevRaw = (kids[i - 1].textContent || "").trim();
        if (SEMAI_CLOSING_RE.test(prevRaw)) {
          return i + 1 < kids.length ? kids[i + 1] : null;
        }
      }
    }

    // Pattern C: sender's full name first line of a contact block
    if (!senderName) continue;
    if (firstLine.toLowerCase().startsWith(senderName)) {
      const charAfter = firstLine[senderName.length];
      if (!charAfter || /[\s,.]/.test(charAfter)) {
        const longLines = lines.filter(l => l.length > 120).length;
        if (lines.length >= 2 && longLines <= 2) return kids[i];
      }
    }
  }
  return null;
}

// ── Strip trailing sign-off ────────────────────────────────────────────────────
function semaiStripTrailingSignOff(container, senderFirstName) {
  let scope = container;
  while (scope.children.length === 1 && SEMAI_BLOCK_TAGS.has(scope.children[0].tagName)) {
    scope = scope.children[0];
  }
  const kids = Array.from(scope.children);
  if (kids.length === 0) return;

  let last = kids.length - 1;
  while (last >= 0 && !(kids[last].textContent || "").trim()) last--;
  if (last < 0) return;

  const lastRaw   = (kids[last].textContent || "").trim();
  const lastLines = lastRaw.split(/\r?\n|\r/).map(l => l.trim()).filter(Boolean);

  // Single element: "Best,\nMichael"
  if (lastLines.length === 2 && SEMAI_CLOSING_RE.test(lastLines[0])) {
    const nameWords = lastLines[1].split(/\s+/);
    if (nameWords.length <= 3 && nameWords.every(w => /^[A-Z]/.test(w))) {
      if (!senderFirstName || lastLines[1].toLowerCase().startsWith(senderFirstName)) {
        kids[last].remove(); return;
      }
    }
  }

  // "Michael" preceded by "Best,"
  if (lastLines.length === 1 && last > 0) {
    const words = lastRaw.split(/\s+/);
    if (words.length <= 3 && lastRaw.length <= 30 && words.every(w => /^[A-Z]/.test(w))) {
      if (!senderFirstName || lastRaw.toLowerCase().startsWith(senderFirstName)) {
        const prevRaw = (kids[last - 1].textContent || "").trim();
        if (SEMAI_CLOSING_RE.test(prevRaw)) {
          kids[last].remove();
          kids[last - 1].remove();
          return;
        }
      }
    }
  }

  // Just a closing phrase at the end
  if (lastLines.length === 1 && SEMAI_CLOSING_RE.test(lastRaw)) {
    kids[last].remove();
  }
}

// ── Main clone-cleaner (mirrors the one in contentScript.js) ──────────────────
function semaiCleanBodyClone(bodyEl, senderFirstName) {
  const clone = document.createElement("div");
  if (bodyEl.dataset.semaiOriginalHtml) {
    clone.innerHTML = bodyEl.dataset.semaiOriginalHtml;
  } else {
    clone.innerHTML = bodyEl.innerHTML;
  }

  // 0a. Short-circuit: if the entire body is a compact branded signature, return empty
  if (semaiIsEntireBodySignature(clone, senderFirstName)) {
    return "";
  }

  // 0. Remove Outlook "external sender" warning banners
  //    These are injected by Exchange/Outlook and contain links to aka.ms/LearnAboutSenderIdentification
  clone.querySelectorAll('a[href*="LearnAboutSenderIdentification"]').forEach(a => {
    // Walk up to the nearest block ancestor (table, div, td) and remove it
    let el = a;
    while (el && el !== clone) {
      const tag = el.tagName;
      if (tag === "TABLE" || (tag === "DIV" && el.parentElement !== clone)) {
        el.remove();
        break;
      }
      el = el.parentElement;
    }
  });

  // 1. Remove Outlook reply/forward header blocks
  clone.querySelectorAll(
    '#divRplyFwdMsg, div[id*="divRplyFwdMsg"], div[id*="appendonsend"]'
  ).forEach(el => { let sib = el; while (sib) { const next = sib.nextElementSibling; sib.remove(); sib = next; } });

  // 2. Remove Outlook mobile reference messages
  clone.querySelectorAll(
    '#mail-editor-reference-message-container, [class*="reference-message" i]'
  ).forEach(el => el.remove());

  // 3. Remove signature wrapper divs
  clone.querySelectorAll('[id*="signature" i], [class*="signature" i]').forEach(el => el.remove());

  // 4. Remove quoted-reply header blocks
  semaiStripQuotedReplyHeaders(clone);

  // 5. Strip separator lines and everything after
  function removeFromAndAfter(el) {
    let sib = el;
    while (sib) { const next = sib.nextElementSibling; sib.remove(); sib = next; }
  }
  const hrEl = clone.querySelector("hr");
  if (hrEl) removeFromAndAfter(hrEl);

  const dashSep = semaiFirstSeparatorEl(clone, /^--\s*$/);
  if (dashSep) removeFromAndAfter(dashSep);

  const underscoreSep = semaiFirstSeparatorEl(clone, /^_{4,}$/);
  if (underscoreSep) removeFromAndAfter(underscoreSep);

  // 6. Strip sign-off + sender name contact block
  const anchor = semaiFindSenderAnchor(clone, senderFirstName);
  if (anchor) removeFromAndAfter(anchor);

  // 6b. Strip repeated-name contact cards
  // Support both string and array of tokens for name detection
  const nameTokensForClone = Array.isArray(senderFirstName) ? senderFirstName
    : (senderFirstName ? [senderFirstName] : []);
  const cloneNameArg = nameTokensForClone.length > 0 ? nameTokensForClone : senderFirstName;
  const repeatedNameAnchor = semaiFindRepeatedNameSignatureAnchor(clone, cloneNameArg);
  if (repeatedNameAnchor) removeFromAndAfter(repeatedNameAnchor);

  // 6c. Strip standalone contact cards (name + credentials, no sign-off)
  const standaloneCard = semaiFindStandaloneContactCard(clone, cloneNameArg);
  if (standaloneCard) removeFromAndAfter(standaloneCard);

  // 7. Strip specific nested-div contact-card signatures
  const nestedDivSig = semaiFindNestedDivSignature(clone);
  if (nestedDivSig) nestedDivSig.remove();

  // 8. Strip closing phrase + name even when no contact block follows
  semaiStripTrailingSignOff(clone, senderFirstName);

  // 9. Remove trailing empty elements
  while (clone.lastElementChild) {
    const last = clone.lastElementChild;
    if (!(last.textContent || "").trim()) last.remove();
    else break;
  }

  // 10. Remove mobile auto-signature lines
  clone.querySelectorAll("p, div, span, td").forEach(el => {
    if (semaiIsMobileSigEl(el)) el.remove();
  });

  return clone.innerHTML.trim();
}

// ─── Test helper ─────────────────────────────────────────────────────────────
function extractText(html) {
  const div = document.createElement("div");
  div.innerHTML = html;
  return (div.textContent || "").replace(/\s+/g, " ").trim();
}

function runTest(label, originalHtml, senderFirstName, mustKeep, mustRemove) {
  console.log(`\n${"=".repeat(60)}`);
  console.log(`TEST: ${label}`);
  console.log("=".repeat(60));

  // Build a fake bodyEl with dataset.semaiOriginalHtml set
  const bodyEl = document.createElement("div");
  bodyEl.dataset.semaiOriginalHtml = originalHtml;

  const cleanHtml  = semaiCleanBodyClone(bodyEl, senderFirstName);
  const cleanText  = extractText(cleanHtml);

  console.log("\n--- Clean text output ---");
  console.log(cleanText);
  console.log("---");

  let passed = true;

  for (const phrase of mustKeep) {
    if (!cleanText.includes(phrase)) {
      console.log(`  FAIL: expected to keep "${phrase}" but it was removed`);
      passed = false;
    } else {
      console.log(`  OK  : kept "${phrase}"`);
    }
  }

  for (const phrase of mustRemove) {
    if (cleanText.includes(phrase)) {
      console.log(`  FAIL: expected to remove "${phrase}" but it was kept`);
      passed = false;
    } else {
      console.log(`  OK  : removed "${phrase}"`);
    }
  }

  console.log(`\nRESULT: ${passed ? "PASS" : "FAIL"}`);
  return passed;
}

// ─── Test case 1: Daniel Drane / Emory contact card ──────────────────────────
// Source: data-semai-original-html decoded from the HTML in the docs file.
const danielOriginalHtml = `<div visibility="hidden"><div>
<div dir="ltr">
<div lang="en-US" style="word-wrap:break-word;">
<div>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Let's take a look at the summary statements and see what kind of feedback we're getting.&nbsp; I am happy that you got the K99/R00 back. That softens the blow a bit.&nbsp; </p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">This was a tough cycle, as there were a lot of backed up grants from the fall…I got one discussed but I had a lot of friends not get discussed at all. </p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">daniel</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>
<div>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Daniel L. Drane, Ph.D., ABPP(CN)</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Professor of Neurology and Pediatrics</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Emory University School of Medicine</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Woodruff Memorial Research Building</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">101 Woodruff Circle, Suite 6111</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Mailstop: 1930-001-1AN</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Atlanta, Georgia 30322</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">(404)727-2844 (office)</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">ddrane@emory.edu</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Affiliate Professor of Neurology</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">University of Washington</p></div></div></div></div></div>
</div>`;

const danielPassed = runTest(
  "Daniel Drane / Emory contact card",
  danielOriginalHtml,
  "daniel",
  // Must KEEP
  [
    "Let's take a look at the summary statements",
    "This was a tough cycle"
  ],
  // Must REMOVE
  [
    "Daniel L. Drane, Ph.D.",
    "Professor of Neurology and Pediatrics",
    "Emory University School of Medicine",
    "(404)727-2844",
    "ddrane@emory.edu"
  ]
);

// ─── Test case 2: Mandi Schmitt / UserTesting outreach signature ──────────────
// Full decoded HTML from data-semai-original-html attribute.
const mandiOriginalHtml = `<div visibility="hidden"><div>
<div dir="ltr">
<div><table align="left" border="0" cellspacing="0" cellpadding="0" style="width:100%;">
<tbody><tr>
<td valign="middle" bgcolor="#A6A6A6" style="width:0;padding:7px 2px;"></td>
<td valign="middle" bgcolor="#EAEAEA" style="font-size:12px;font-family:wf_segoe-ui_normal,Segoe UI,Tahoma,Arial,sans-serif;width:100%;padding:7px 5px 7px 15px;">
<div>You don't often get email from mschmitt@usertesting.com. <a href="https://aka.ms/LearnAboutSenderIdentification" target="_blank" rel="noopener noreferrer">Learn why this is important</a> </div></td>
<td align="left" valign="middle" bgcolor="#EAEAEA" style="font-size:12px;font-family:wf_segoe-ui_normal,Segoe UI,Tahoma,Arial,sans-serif;width:75px;padding:7px 5px;"></td></tr></tbody></table>
<div>
<div style="font-size:12px;font-family:sans-serif;">
<div style="font-size:12px;font-family:sans-serif;">
<div style="font-size:9pt;font-family:Arial,Helvetica,sans-serif;">
<div>Hi Santiago,</div>
<div style="font-size:9pt;font-family:Arial,Helvetica,sans-serif;">&nbsp;</div>
<div>I'm reaching out to ensure we're fully supporting Emory National Primate Research Center ahead of your renewal on June 17, 2026.</div>
<div style="font-size:9pt;font-family:Arial,Helvetica,sans-serif;">&nbsp;</div>
<div style="font-size:9pt;font-family:Arial,Helvetica,sans-serif;">If there's anything standing in the way of finalizing the next phase of our partnership—whether budget, scope, or internal priorities—we're happy to help work through it.</div>
<div style="font-size:9pt;font-family:Arial,Helvetica,sans-serif;">&nbsp;</div>
<div>We've really enjoyed partnering with your team and are excited about what's next. Let us know how we can help move things forward.</div>
<div style="font-size:9pt;font-family:Arial,Helvetica,sans-serif;">&nbsp;</div>
<div>Best,</div>
<div style="font-size:9pt;font-family:Arial,Helvetica,sans-serif;">Mandi</div></div>
<div>
<div style="font-size:9pt;font-family:sans-serif;">
<div><table border="0" cellspacing="0" cellpadding="0" style="height:138px;">
<tbody><tr style="height:138px;">
<td valign="top" width="92">
  <img src="https://example.com/headshot.jpg" width="82" height="82" alt="Mandi Schmitt headshot" />
</td>
<td valign="top" width="488">
  <div><strong>Mandi Schmitt</strong></div>
  <div>Senior Account Executive</div>
  <div>UserTesting</div>
  <div><a href="mailto:mschmitt@usertesting.com">mschmitt@usertesting.com</a></div>
  <div><a href="https://www.linkedin.com/in/mandi-schmitt">LinkedIn</a></div>
  <div><a href="https://calendar.usertesting.com/mschmitt">Schedule a meeting</a></div>
</td></tr></tbody></table></div>
<div style="font-size:9pt;font-family:Arial;line-height:16px;">&nbsp;</div></div></div><br>
<br>
<div><a href="https://hello.userinterviews.com/t/113544/opt_out/4b07c105-7177-481c-b916-81114b6a189d">Would you like to opt out?</a></div></div></div></div>
</div></div></div>`;

const mandiPassed = runTest(
  "Mandi Schmitt / UserTesting outreach signature",
  mandiOriginalHtml,
  "mandi",
  // Must KEEP
  [
    "Hi Santiago,",
    "I'm reaching out to ensure we're fully supporting",
    "We've really enjoyed partnering"
  ],
  // Must REMOVE
  [
    "Senior Account Executive",
    "mschmitt@usertesting.com",
    "Would you like to opt out?"
  ]
);

// ─── Test case 3: Leah Ekube / Pendo compact branded signature ───────────────
// The ENTIRE email is a signature — no message body. Expected: empty output.
// Source: data-semai-original-html decoded from the HTML in the docs file.
const leahOriginalHtml = `<div visibility="hidden"><div>
<div dir="ltr">
<div dir="ltr">
<p style="color:black;font-family:Arial;margin:0;line-height:normal;font-stretch:normal;font-size-adjust:none;"><b>Leah Ekube<i>&nbsp;</i></b><i>(she/her)</i><i></i></p>
<p style="color:black;font-family:Arial;margin:0;line-height:normal;font-stretch:normal;font-size-adjust:none;"><span>Renewal Specialist&nbsp;&nbsp;</span>|<span>&nbsp; </span><a href="https://nam11.safelinks.protection.outlook.com/?url=http%3A%2F%2Fwww.pendo.io%2F" target="_blank" rel="noopener noreferrer"><span style="color:#103CC0;">pendo.io</span></a><span>&nbsp; </span>|<span>&nbsp;&nbsp;</span></p>
<p style="color:#D1D3DF;font-size:11px;font-family:Arial;margin:0;line-height:normal;font-stretch:normal;font-size-adjust:none;">————</p>
<p style="color:#D1D3DF;font-size:11px;font-family:Arial;margin:0;line-height:normal;font-stretch:normal;font-size-adjust:none;"><img data-imagetype="External" src="https://ci3.googleusercontent.com/mail-sig/AIorK4xDj2uZt40t0dFiwuMNlX6Yz_XkAo-922MqPlFHSY67LhVyOWr0rGkRxeV5XK0eNExm6ewXXOCMFJ0T" width="420" height="56"></p></div></div></div></div>
</div>`;

const leahPassed = runTest(
  "Leah Ekube / Pendo compact branded signature",
  leahOriginalHtml,
  "leah",
  // Must KEEP — nothing, the entire email is a signature
  [],
  // Must REMOVE — the signature content should all be gone
  [
    "Leah Ekube",
    "Renewal Specialist",
    "pendo.io"
  ]
);

// ─── Test case 4: Wilbur Lam / last-name-first display + long dept lines ───────
// Sender displayed as "Lam, Wilbur" in Outlook.  The signature starts with
// "Wilbur A. Lam, MD, PhD" and contains a line exceeding 80 chars.
// senderFirstName should resolve to "wilbur" (not "lam" or "lam,").
const lamOriginalHtml = `<div visibility="hidden"><div>
<div dir="ltr">
<div dir="ltr">
<div lang="en-US" style="word-wrap:break-word;">
<div>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:11pt;font-family:Calibri,sans-serif;">Let's change my stuff so that it's more catered to the DigitalStudio. See below and feel free to edit</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:11pt;font-family:Calibri,sans-serif;">&nbsp;</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:11pt;font-family:Calibri,sans-serif;">Dr. Lam is a Professor in the Department of Pediatrics, Division of Pediatric Hematology/Oncology and Dept of Biomedical Engineering, Emory University and Georgia Institute of Technology. His unique background as a physician-scientist-engineer trained in clinical pediatric hematology/oncology as well as bioengineering makes him an ideal Director of the DigitalStudio.</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:11pt;font-family:Calibri,sans-serif;">&nbsp;</span></p>
<div>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:9pt;">Wilbur A. Lam, MD, PhD</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:9pt;">Professor and W. Paul Bowers Research Chair</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:9pt;">Department of Pediatrics and the Wallace H. Coulter Department of Biomedical Engineering</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:9pt;">Aflac Cancer and Blood Disorders Center at Children's Healthcare of Atlanta</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:9pt;">Director, Center for the Advancement of Diagnostics for a Just Society (ADJUST Center)</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:9pt;">Co-Director, Pediatric Technology Center</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:9pt;">Associate Dean of Innovation, Emory University School of Medicine</span></p>
<p style="font-size:12pt;font-family:Aptos,sans-serif;margin:0;"><span style="font-size:9pt;">Vice Provost for Entrepreneurship, Emory University</span></p>
</div></div></div></div></div></div>
</div>`;

const lamPassed = runTest(
  "Wilbur Lam / last-name-first display name",
  lamOriginalHtml,
  "wilbur",   // semaiFirstNameFromDisplayName("Lam, Wilbur") → "wilbur"
  // Must KEEP
  [
    "Let's change my stuff so that it's more catered to the DigitalStudio",
    "Dr. Lam is a Professor in the Department of Pediatrics"
  ],
  // Must REMOVE
  [
    "Wilbur A. Lam, MD, PhD",
    "Professor and W. Paul Bowers Research Chair",
    "Wallace H. Coulter Department of Biomedical Engineering",
    "Vice Provost for Entrepreneurship"
  ]
);

// ─── Test case 5: Daniel Drane no-signoff variant ──────────────────────────────
// Short body "All great ideas!" followed directly by contact card — NO lowercase
// "daniel" sign-off before the contact card.
const danielNoSignoffHtml = `<div visibility="hidden"><div>
<div lang="en-US" style="word-wrap:break-word;">
<div>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">All great ideas!</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>
<div>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Daniel L. Drane, Ph.D., ABPP(CN)</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Professor of Neurology and Pediatrics</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Emory University School of Medicine</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Woodruff Memorial Research Building</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">101 Woodruff Circle, Suite 6111</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Mailstop: 1930-001-1AN</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Atlanta, Georgia 30322</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">(404)727-2844 (office)</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">ddrane@emory.edu</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Affiliate Professor of Neurology</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">University of Washington</p></div></div></div></div>
</div>`;

const danielNoSignoffPassed = runTest(
  "Daniel Drane no-signoff variant",
  danielNoSignoffHtml,
  "daniel",
  // Must KEEP
  [
    "All great ideas!"
  ],
  // Must REMOVE
  [
    "Daniel L. Drane, Ph.D.",
    "Professor of Neurology and Pediatrics",
    "Emory University School of Medicine",
    "(404)727-2844",
    "ddrane@emory.edu"
  ]
);

// ─── Test case 5: Daniel Drane no-signoff, Last-First sender format ─────────
// Same contact card as case 4, but sender name is "Drane, Daniel L" (last-first
// format from Outlook). semaiGetSenderNameTokens would return ["drane", "daniel"].
// The standalone contact card detector must match "Daniel" at the start of the
// first line using any of the name tokens.
const danielLastFirstHtml = `<div visibility="hidden"><div>
<div lang="en-US" style="word-wrap:break-word;">
<div>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">All great ideas!</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>
<div>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Daniel L. Drane, Ph.D., ABPP(CN)</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Professor of Neurology and Pediatrics</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Emory University School of Medicine</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Woodruff Memorial Research Building</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">101 Woodruff Circle, Suite 6111</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Mailstop: 1930-001-1AN</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Atlanta, Georgia 30322</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">(404)727-2844 (office)</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">ddrane@emory.edu</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">&nbsp;</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">Affiliate Professor of Neurology</p>
<p style="font-size:11pt;font-family:Calibri,sans-serif;margin:0;">University of Washington</p></div></div></div></div>
</div>`;

const danielLastFirstPassed = runTest(
  "Daniel Drane no-signoff, Last-First sender format",
  danielLastFirstHtml,
  ["drane", "daniel"],  // simulates semaiGetSenderNameTokens for "Drane, Daniel L"
  // Must KEEP
  [
    "All great ideas!"
  ],
  // Must REMOVE
  [
    "Daniel L. Drane, Ph.D.",
    "Professor of Neurology and Pediatrics",
    "Emory University School of Medicine",
    "(404)727-2844",
    "ddrane@emory.edu"
  ]
);

// ─── Test case 6: "Sent from my iPhone" mobile auto-signature ────────────────
// A simple email body ending with just "Sent from my iPhone". The phrase should
// be stripped; the preceding message text must be kept.
const mobileAutoSigHtml = `<div>
<p>Hey, just checking in. Are we still on for Thursday?</p>
<p>&nbsp;</p>
<p>Sent from my iPhone</p>
</div>`;

const mobileAutoSigPassed = runTest(
  "Sent from my iPhone mobile auto-signature",
  mobileAutoSigHtml,
  null,
  // Must KEEP
  [
    "Hey, just checking in. Are we still on for Thursday?"
  ],
  // Must REMOVE
  [
    "Sent from my iPhone"
  ]
);

// ─── Summary ──────────────────────────────────────────────────────────────────
console.log(`\n${"=".repeat(60)}`);
console.log("SUMMARY");
console.log("=".repeat(60));
console.log(`  Daniel Drane case       : ${danielPassed ? "PASS" : "FAIL"}`);
console.log(`  Daniel no-signoff case  : ${danielNoSignoffPassed ? "PASS" : "FAIL"}`);
console.log(`  Daniel Last-First case  : ${danielLastFirstPassed ? "PASS" : "FAIL"}`);
console.log(`  Leah Ekube case         : ${leahPassed ? "PASS" : "FAIL"}`);
console.log(`  Mandi Schmitt case      : ${mandiPassed ? "PASS" : "FAIL"}`);
console.log(`  Wilbur Lam case         : ${lamPassed ? "PASS" : "FAIL"}`);
console.log(`  Mobile auto-sig case    : ${mobileAutoSigPassed ? "PASS" : "FAIL"}`);
console.log("=".repeat(60));

process.exit(danielPassed && danielNoSignoffPassed && danielLastFirstPassed && leahPassed && mandiPassed && lamPassed && mobileAutoSigPassed ? 0 : 1);
