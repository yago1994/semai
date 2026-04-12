// semaiSigDetector.js — Signature detection and body cleaning
//
// Extracted from contentScript.js so this logic can be read in isolation
// by the live-fix Claude API call without including the full UI layer.
//
// Depends on: semaiNativeLog, semaiLog (defined in contentScript.js, same world)
// Loaded before contentScript.js in manifest.json.

// ── Signature detection constants ────────────────────────────────────────────

const SEMAI_SIG_SHORT_LINE_MAX = 60;
const SEMAI_SIG_MIN_LINES = 5;
const SEMAI_PHONE_RE = /\+?[\d][\d\s\-\.\(\)]{6,}/;
const SEMAI_EMAIL_RE = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/i;
const SEMAI_URL_RE = /https?:\/\/|www\./i;
const SEMAI_SOCIAL_RE = /linkedin\.com|twitter\.com|facebook\.com|instagram\.com/i;
const SEMAI_SIGNATURE_TITLE_RE = /\b(professor|director|chair|vice chair|clinical research|neurology|pediatrics|medicine|program|department|affiliate professor|research|health)\b/i;
const SEMAI_SIGNATURE_ADDRESS_RE = /\b(suite|mailstop|building|circle|road|street|avenue|drive|lane|boulevard|floor|room|atlanta|georgia|\d{5}(?:-\d{4})?)\b/i;
const SEMAI_BLOCK_TAGS = new Set([
  "P","DIV","BLOCKQUOTE","LI","H1","H2","H3","H4","H5","H6",
  "SECTION","ARTICLE","HEADER","FOOTER","MAIN","TABLE","TR","TD","TH"
]);
const SEMAI_CLOSING_RE = /^(best|regards|thanks|thank you|cheers|sincerely|warmly|warm regards|kind regards|best regards|yours|cordially|take care|many thanks|with gratitude|respectfully|yours truly|talk soon|looking forward)[,.]?\s*$/i;
const SEMAI_CREDENTIALS_RE = /\b(ph\.?d|m\.?d|m\.?s|m\.?a|m\.?b\.?a|j\.?d|ed\.?d|d\.?o|r\.?n|b\.?s|b\.?a|abpp|faes|faan|fana|facs|frcp|lcsw|lpc|lmft|cpa|esq|pe)\b/i;
const SEMAI_MOBILE_SIG_RE = /^(sent from my (iphone|ipad|android)|sent from mobile|sent from outlook for (ios|android)|get outlook for (ios|android))$/i;

// ── Name normalization helpers ────────────────────────────────────────────────

function semaiNormalizeNameLine(text) {
  return (text || "")
    .replace(/[^\p{L}\s.'-]/gu, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function semaiFoldNameForMatch(text) {
  return semaiNormalizeNameLine(text)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "");
}

function semaiNameMatchRemainder(text, nameToken) {
  const foldedText = semaiFoldNameForMatch((text || "").replace(/^[^\p{L}]+/gu, ""));
  const foldedToken = semaiFoldNameForMatch(nameToken);
  if (!foldedText.startsWith(foldedToken)) return null;
  return foldedText.slice(foldedToken.length);
}

function semaiGetNameTokens(text) {
  return semaiNormalizeNameLine(text)
    .split(/\s+/)
    .filter((token) => token.length >= 2);
}

function semaiLooksLikeNameLine(text, senderFirstName) {
  const raw = (text || "").trim();
  if (!raw || raw.length > 60) return false;

  const tokens = raw.split(/\s+/).filter(Boolean);
  if (tokens.length === 0 || tokens.length > 5) return false;

  const nameish = tokens.every((token) => /^[A-Z][A-Za-z.'-]*$/.test(token) || /^[A-Z]\.$/.test(token));
  if (nameish) return true;

  return !!senderFirstName && semaiFoldNameForMatch(raw) === semaiFoldNameForMatch(senderFirstName);
}

// ── Contact signal counting ───────────────────────────────────────────────────

function semaiCountContactSignals(text, containerEl) {
  const signalKinds = new Set();
  let raw = text || "";
  if (containerEl && typeof containerEl.querySelectorAll === "function") {
    const hrefs = Array.from(containerEl.querySelectorAll("a"))
      .map(a => a.getAttribute("href") || "")
      .filter(Boolean)
      .join(" ");
    if (hrefs) raw = raw + " " + hrefs;
  }

  if (SEMAI_PHONE_RE.test(raw)) signalKinds.add("phone");
  if (SEMAI_EMAIL_RE.test(raw)) signalKinds.add("email");
  if (SEMAI_URL_RE.test(raw) || SEMAI_SOCIAL_RE.test(raw)) signalKinds.add("url");
  if (SEMAI_SIGNATURE_TITLE_RE.test(raw)) signalKinds.add("title");
  if (SEMAI_SIGNATURE_ADDRESS_RE.test(raw)) signalKinds.add("address");

  return signalKinds.size;
}

// ── Block-shape detection ─────────────────────────────────────────────────────

function semaiLooksLikeCompactBlock(el) {
  const text = (el?.innerText || el?.textContent || "").trim();
  if (!text) return false;
  const lines = text.split(/\n+/).map((line) => line.trim()).filter(Boolean);
  if (lines.length > 18) return false;
  return text.length <= 600;
}

function semaiLooksLikeContactCardBlock(el) {
  if (!(el instanceof HTMLElement)) return false;
  const text = (el.innerText || el.textContent || "").trim();
  if (!text) return false;
  const lines = text.split(/\n+/).map((line) => line.trim()).filter(Boolean);
  const shortLines = lines.filter((line) => line.length <= 90).length;
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

function semaiLooksLikeSig(el) {
  const text = el.innerText || el.textContent || "";
  const lines = text.split("\n").map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length < SEMAI_SIG_MIN_LINES) return false;
  const shortLines = lines.filter(l => l.length <= SEMAI_SIG_SHORT_LINE_MAX).length;
  if (shortLines / lines.length < 0.8) return false;
  return SEMAI_PHONE_RE.test(text) || SEMAI_URL_RE.test(text) || SEMAI_SOCIAL_RE.test(text);
}

// ── Nested-div signature detection ───────────────────────────────────────────

function semaiLooksLikeNestedDivSignature(el) {
  if (!(el instanceof HTMLElement)) return false;
  if (el.tagName !== "DIV" || el.getAttribute("dir")?.toLowerCase() !== "ltr") return false;

  const childDivs = Array.from(el.children).filter((child) => child.tagName === "DIV");
  if (childDivs.length < 4 || childDivs.length !== el.children.length) return false;

  const lines = (el.innerText || el.textContent || "")
    .split(/\r?\n|\r/)
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length < 6) return false;

  const hasPhone = lines.some((line) => SEMAI_PHONE_RE.test(line));
  const hasEmailOrUrl = lines.some((line) => /@/.test(line) || /\b(?:[a-z0-9-]+\.)+[a-z]{2,}(?:\/[^\s]*)?\b/i.test(line));
  const hasOrgLine = lines.some((line) => /\b(university|school|medicine|program|director|professor|chair|clinical research|department|health)\b/i.test(line));

  return hasPhone && hasEmailOrUrl && hasOrgLine;
}

function semaiFindNestedDivSignature(container) {
  const candidates = Array.from(container.querySelectorAll('div[dir="ltr"]'));
  for (let i = candidates.length - 1; i >= 0; i--) {
    if (semaiLooksLikeNestedDivSignature(candidates[i])) {
      return candidates[i];
    }
  }
  return null;
}

function semaiIsVisuallyEmptyElement(el) {
  if (!(el instanceof HTMLElement)) return false;
  if (el.querySelector("img, video, audio, iframe, object, embed, svg, canvas, hr")) return false;
  const text = (el.innerText || el.textContent || "").replace(/\u00a0/g, " ").trim();
  return text.length === 0;
}

function semaiTrimTrailingEmptyBlocks(container) {
  if (!(container instanceof HTMLElement)) return;

  while (container.lastElementChild) {
    const last = container.lastElementChild;
    semaiTrimTrailingEmptyBlocks(last);
    if (!semaiIsVisuallyEmptyElement(last)) break;
    last.remove();
  }
}

// ── Entire-body signature detection ──────────────────────────────────────────

// Detect when the entire body is a compact branded signature with no actual message.
function semaiIsEntireBodySignature(clone, senderFirstName) {
  if (!senderFirstName) return false;
  const text = (clone.textContent || "").replace(/\s+/g, " ").trim();
  if (!text) return false;
  if (text.length > 800) return false;

  const lines = text.split(/(?<=[.!?])\s+|\n+/).map(l => l.trim()).filter(Boolean);
  if (lines.length === 0) return false;
  const firstThreeText = semaiFoldNameForMatch(lines.slice(0, 3).join(" "));
  if (!firstThreeText.includes(semaiFoldNameForMatch(senderFirstName))) return false;

  const hasUrl = SEMAI_URL_RE.test(text) || !!clone.querySelector("img, a[href]");
  if (!hasUrl) return false;

  const hasLongSentence = lines.some(l => l.length > 60 && /\s/.test(l));
  if (hasLongSentence) return false;

  // Must NOT have a line that looks like a complete prose sentence:
  // ≥ 4 words, > 25 chars, ending with sentence punctuation (. ? !)
  const hasProseSentence = lines.some(l =>
    l.length > 25 && /[.?!]$/.test(l) && l.split(/\s+/).length >= 4
  );
  if (hasProseSentence) return false;

  return true;
}

// ── Repeated-name contact card detection ─────────────────────────────────────

function semaiFindRepeatedNameSignatureAnchor(container, senderFirstNameOrTokens) {
  const nameTokens = Array.isArray(senderFirstNameOrTokens)
    ? senderFirstNameOrTokens
    : (senderFirstNameOrTokens ? [senderFirstNameOrTokens.toLowerCase()] : []);
  const senderFirstName = nameTokens.length > 0 ? nameTokens[0] : null;

  const blocks = Array.from(container.querySelectorAll("p, div, table, td"))
    .filter((el) => semaiLooksLikeCompactBlock(el));
  if (blocks.length < 3) return null;

  const startAt = Math.max(0, blocks.length - 24);
  for (let i = startAt; i < blocks.length; i++) {
    const currentText = (blocks[i].innerText || blocks[i].textContent || "").trim();
    const matchesAnyToken = nameTokens.length > 0
      ? nameTokens.some(token => semaiLooksLikeNameLine(currentText, token))
      : semaiLooksLikeNameLine(currentText, null);
    if (!matchesAnyToken) continue;

    const currentTokens = semaiGetNameTokens(currentText);
    if (currentTokens.length === 0) continue;

    for (let j = i + 1; j < Math.min(blocks.length, i + 8); j++) {
      const nextEl = blocks[j];
      const nextText = (nextEl.innerText || nextEl.textContent || "").trim();
      if (!nextText) continue;

      const nextTokens = semaiGetNameTokens(nextText);
      const repeatedName = currentTokens.some((token) => nextTokens.includes(token));
      if (!repeatedName) continue;

      let signalText = nextText;
      let signalEls = [nextEl];
      let contactCardSeen = semaiLooksLikeContactCardBlock(nextEl);
      for (let k = j + 1; k < Math.min(blocks.length, j + 8); k++) {
        const blockText = (blocks[k].innerText || blocks[k].textContent || "").trim();
        if (!blockText) continue;
        signalText += "\n" + blockText;
        signalEls.push(blocks[k]);
        if (semaiLooksLikeContactCardBlock(blocks[k])) contactCardSeen = true;
      }

      const signalWrapper = document.createElement("div");
      signalEls.forEach(e => signalWrapper.appendChild(e.cloneNode(true)));

      if (semaiCountContactSignals(signalText, signalWrapper) >= 2 && contactCardSeen) {
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

// ── Standalone contact-card detection ────────────────────────────────────────
// Detects a contact card block by its own signals: sender's first name with
// credentials (Ph.D., M.D., etc.), plus contact signals (phone, email, address,
// title, institution). Used when there is no lowercase sign-off before the card.

function semaiFindStandaloneContactCard(container, senderFirstNameOrTokens) {
  const nameTokens = Array.isArray(senderFirstNameOrTokens)
    ? senderFirstNameOrTokens
    : (senderFirstNameOrTokens ? [senderFirstNameOrTokens.toLowerCase()] : []);
  if (nameTokens.length === 0) return null;

  const blocks = Array.from(container.querySelectorAll("p, div, table, td"))
    .filter(el => {
      const text = (el.innerText || el.textContent || "").trim();
      return text && text.length <= 600;
    });
  if (blocks.length < 2) return null;

  for (let i = 0; i < blocks.length; i++) {
    const el = blocks[i];
    const text = (el.innerText || el.textContent || "").trim();
    if (!text) continue;

    const lines = text.split(/\r?\n|\r/).map(l => l.trim()).filter(Boolean);
    if (lines.length === 0) continue;

    const firstLine = lines[0];
    // Strip leading non-letter chars (e.g. dashes like "-Gaëlle Sabben, M.P.H.")
    const firstLineStripped = firstLine.replace(/^[^\p{L}]+/u, "").trim();

    const matchedToken = nameTokens.find((token) => semaiNameMatchRemainder(firstLineStripped, token) !== null);
    if (!matchedToken) continue;
    const remainder = semaiNameMatchRemainder(firstLineStripped, matchedToken);
    const charAfter = remainder?.[0];
    if (charAfter && !/[\s,.]/.test(charAfter)) continue;

    const hasCredentials = SEMAI_CREDENTIALS_RE.test(firstLine);
    const hasTitle = lines.length >= 2 && SEMAI_SIGNATURE_TITLE_RE.test(lines.slice(0, 3).join(" "));
    if (!hasCredentials && !hasTitle) continue;

    let signalText = text;
    let signalEls = [el];
    for (let k = i + 1; k < Math.min(blocks.length, i + 10); k++) {
      const sibText = (blocks[k].innerText || blocks[k].textContent || "").trim();
      if (!sibText) continue;
      if (!el.parentElement || !el.parentElement.contains(blocks[k])) {
        if (!container.contains(blocks[k])) break;
      }
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
        const sibText = (sib.innerText || sib.textContent || "").trim();
        return sibText.length > 0 && sibText !== "\u00a0";
      });
      if (hasBodyTextSibling) { cutEl = walker; break; }
      walker = parent;
    }

    return cutEl;
  }

  return null;
}

function semaiFindClosingNameSignatureAnchor(container, senderFirstNameOrTokens) {
  const nameTokens = Array.isArray(senderFirstNameOrTokens)
    ? senderFirstNameOrTokens
    : (senderFirstNameOrTokens ? [senderFirstNameOrTokens.toLowerCase()] : []);
  const blocks = Array.from(container.querySelectorAll("p, div, td"))
    .filter((el) => semaiLooksLikeCompactBlock(el));

  if (blocks.length < 4) return null;

  const startAt = Math.max(0, blocks.length - 20);
  for (let i = startAt; i < blocks.length - 2; i++) {
    const closingText = (blocks[i].innerText || blocks[i].textContent || "").trim();
    if (!SEMAI_CLOSING_RE.test(closingText)) continue;

    const nameText = (blocks[i + 1].innerText || blocks[i + 1].textContent || "").trim();
    if (!nameText) continue;

    const nameMatches = semaiLooksLikeNameLine(nameText, nameTokens[0] || null) ||
      (nameTokens.length > 0 && nameTokens.some((token) => semaiFoldNameForMatch(nameText).startsWith(semaiFoldNameForMatch(token))));
    if (!nameMatches) continue;

    let professionalLineCount = 0;
    let sawOrgOrTitleSignal = false;
    for (let j = i + 2; j < Math.min(blocks.length, i + 7); j++) {
      const lineText = (blocks[j].innerText || blocks[j].textContent || "").trim();
      if (!lineText) continue;
      if (lineText.length > 140) break;
      if (SEMAI_CLOSING_RE.test(lineText)) break;

      professionalLineCount++;
      if (SEMAI_SIGNATURE_TITLE_RE.test(lineText) || /\b(university|school|alliance|program|department|institute|center|office|clinical|research|medicine|health)\b/i.test(lineText)) {
        sawOrgOrTitleSignal = true;
      }
    }

    if (professionalLineCount < 2 || !sawOrgOrTitleSignal) continue;

    const signatureBlocks = blocks.slice(i, Math.min(blocks.length, i + 2 + professionalLineCount));
    const signatureAncestors = new Set();
    let anc = blocks[i];
    while (anc && anc !== container) { signatureAncestors.add(anc); anc = anc.parentElement; }

    let cutEl = blocks[i];
    let walker = blocks[i];
    while (walker && walker !== container) {
      const parent = walker.parentElement;
      if (!parent || parent === container) break;
      const containsAllSignatureBlocks = signatureBlocks.every((block) => parent.contains(block));
      if (!containsAllSignatureBlocks) break;

      const siblings = Array.from(parent.children);
      const hasBodyTextSibling = siblings.some((sib) => {
        if (sib === walker) return false;
        const sibText = (sib.innerText || sib.textContent || "").replace(/\u00a0/g, " ").trim();
        return sibText.length > 0;
      });
      if (hasBodyTextSibling) {
        cutEl = walker;
        break;
      }
      cutEl = parent;
      walker = parent;
    }

    return cutEl;
  }

  return null;
}

// ── Table-based signature anchor detection ────────────────────────────────────
// Detects signatures whose contact info is buried inside deeply nested tables
// (common with email-signature generators like Woodruff Center's Exclaimer).
// Strategy: scan the last 40% of the body's direct children; if a child
// contains a <table> with mailto:/tel: links or a font-size:1px wrapper
// (classic sig-generator fingerprint), and that table block appears in the
// last 40% of the total text content, treat it as a signature block.
// If a short sign-off (≤ 4 words) immediately precedes it, collapse from
// that sign-off; otherwise collapse from the table block itself.

function semaiFindTableSignatureAnchor(bodyEl, senderNameTokens) {
  // Unwrap single-child wrapper divs to reach the actual content container
  let scope = bodyEl;
  while (scope.children.length === 1 && SEMAI_BLOCK_TAGS.has(scope.children[0].tagName)) {
    scope = scope.children[0];
  }

  const children = Array.from(scope.children);
  if (children.length === 0) return null;

  const startIdx = Math.floor(children.length * 0.6);

  // Total text length for the 40% position safety guard
  const totalText = (scope.innerText || scope.textContent || "");
  const totalLen = totalText.length;

  function hasContactTable(el) {
    // Check for mailto:/tel: links anywhere in the subtree
    if (el.querySelector('a[href^="mailto:"], a[href^="tel:"], area[href^="mailto:"], area[href^="tel:"]')) {
      return true;
    }
    // Check for font-size:1px wrapper (signature generator fingerprint)
    const allEls = el.querySelectorAll ? Array.from(el.querySelectorAll("*")) : [];
    for (const sub of allEls) {
      const fs = sub.style && sub.style.fontSize;
      if (fs === "1px") return true;
    }
    return false;
  }

  for (let i = startIdx; i < children.length; i++) {
    const child = children[i];
    const childHasTable = child.tagName === "TABLE" || !!child.querySelector("table");
    if (!childHasTable) continue;
    if (!hasContactTable(child)) continue;

    // Safety guard: the table block must not appear in the first 40% of the text.
    // This protects against tables embedded mid-email (e.g. proposal tables).
    // We only apply this guard when total content is large enough to be meaningful
    // (short emails with a single sentence before the sig would fail a strict check).
    if (totalLen > 300) {
      const childText = (child.innerText || child.textContent || "");
      const childPos = totalText.indexOf(childText.slice(0, 30));
      if (childPos > 0 && childPos / totalLen < 0.6) continue;
    }

    // Check for a short sign-off immediately before this block
    if (i > 0) {
      const prevChild = children[i - 1];
      const prevText = (prevChild.innerText || prevChild.textContent || "").trim();
      const prevWords = prevText.split(/\s+/).filter(Boolean);
      if (prevWords.length >= 1 && prevWords.length <= 4 && prevText.length <= 40) {
        return prevChild;
      }
    }

    return child;
  }

  return null;
}

// ── Collapse / hide helpers ───────────────────────────────────────────────────

function semaiMakeSigToggle(wrapper) {
  const btn = document.createElement("button");
  btn.className = "semai-sig-toggle";
  btn.textContent = "Show signature";
  btn.addEventListener("click", () => {
    const nowHidden = wrapper.style.display === "none";
    wrapper.style.display = nowHidden ? "" : "none";
    btn.textContent = nowHidden ? "Hide signature" : "Show signature";
  });
  return btn;
}

// Wraps startEl and all its following siblings in a hidden div with a toggle button.
function semaiCollapseFrom(startEl) {
  const container = startEl.parentElement;
  if (!container) return;

  const toHide = [];
  let curr = startEl;
  while (curr) {
    toHide.push(curr);
    curr = curr.nextElementSibling;
  }
  if (toHide.length === 0) return;

  const wrapper = document.createElement("div");
  wrapper.style.display = "none";
  container.insertBefore(wrapper, startEl);
  toHide.forEach(el => wrapper.appendChild(el));
  container.insertBefore(semaiMakeSigToggle(wrapper), wrapper);
}

function semaiHideEl(el) {
  if (el.dataset.semaiSigHidden) return;
  el.dataset.semaiSigHidden = "true";
  el.style.display = "none";
  el.insertAdjacentElement("beforebegin", semaiMakeSigToggle(el));
}

// Walk text nodes and return the nearest block-level element whose trimmed
// text matches pattern.
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

// ── Sender name extraction ────────────────────────────────────────────────────

// Extracts the true first name from a display name in either "First Last" or
// "Last, First" format (Outlook often stores contacts as "Sabben, Gaelle L.").
function semaiFirstNameFromDisplayName(displayName) {
  const name = (displayName || "").trim();
  if (/,/.test(name)) {
    // "Last, First [Middle]" → take the token right after the comma
    const afterComma = (name.split(/\s*,\s*/)[1] || "").split(/\s+/)[0];
    if (afterComma && afterComma.length >= 2 && /^\p{L}/u.test(afterComma)) {
      const foldedAfterComma = semaiFoldNameForMatch(afterComma);
      semaiNativeLog(`[semai-sig] firstNameFromDisplay: "${name}" → "Last, First" format → "${foldedAfterComma}"`);
      return foldedAfterComma;
    }
  }
  // "First [Middle] Last" → take the first whitespace-separated token
  const first = semaiFoldNameForMatch(name.split(/[\s,<(@]+/)[0] || "");
  semaiNativeLog(`[semai-sig] firstNameFromDisplay: "${name}" → "First Last" format → "${first}"`);
  return first;
}

// Try to extract the sender's first name from Outlook's email header UI.
function semaiGetSenderFirstName(bodyEl) {
  const selectors = [
    '.OZZZK',
    '[aria-label^="From:"]',
    '.ms-Persona-primaryText',
    '[class*="personaName" i]',
    '[class*="persona-name" i]',
    '[data-testid="senderName"]',
    '[data-testid="sender-name"]',
    '[class*="senderName" i]',
    '[class*="sender-name" i]',
    '[class*="fromAddress" i]',
    'button[aria-label]:not([aria-label=""])',
  ];

  let ancestor = bodyEl.parentElement;
  for (let d = 0; d < 12 && ancestor; d++, ancestor = ancestor.parentElement) {
    for (const sel of selectors) {
      try {
        const found = ancestor.querySelector(sel);
        if (!found || found.contains(bodyEl)) continue;
        const raw = (found.getAttribute("aria-label") || found.innerText || found.textContent || "").trim();
        const cleaned = raw.replace(/^from[:\s]+/i, "");
        const firstName = semaiFirstNameFromDisplayName(cleaned);
        semaiNativeLog(`[semai-sig] getSenderFirstName: selector="${sel}" raw="${raw}" → firstName="${firstName}"`);
        if (firstName && firstName.length >= 2 && /^\p{L}/u.test(firstName)) {
          semaiNativeLog(`[semai-sig] getSenderFirstName: resolved to "${firstName}"`);
          return firstName;
        }
      } catch (e) { /* ignore invalid selectors */ }
    }
  }
  return null;
}

// Extract ALL meaningful name tokens from the sender info.
// For "Drane, Daniel L" returns ["drane", "daniel"].
// For "Daniel Drane" returns ["daniel", "drane"].
// Strips single-letter initials.
function semaiGetSenderNameTokens(bodyEl) {
  const selectors = [
    '.OZZZK',
    '[aria-label^="From:"]',
    '.ms-Persona-primaryText',
    '[class*="personaName" i]',
    '[class*="persona-name" i]',
    '[data-testid="senderName"]',
    '[data-testid="sender-name"]',
    '[class*="senderName" i]',
    '[class*="sender-name" i]',
    '[class*="fromAddress" i]',
    'button[aria-label]:not([aria-label=""])',
  ];

  let ancestor = bodyEl.parentElement;
  for (let d = 0; d < 12 && ancestor; d++, ancestor = ancestor.parentElement) {
    for (const sel of selectors) {
      try {
        const found = ancestor.querySelector(sel);
        if (!found || found.contains(bodyEl)) continue;
        const raw = (found.getAttribute("aria-label") || found.innerText || found.textContent || "").trim();
        const cleaned = raw.replace(/^from[:\s]+/i, "");
        const tokens = cleaned.split(/[\s,<(@]+/)
          .map((t) => semaiFoldNameForMatch(t.replace(/[^\p{L}]/gu, "")))
          .filter(t => t.length >= 2);
        if (tokens.length > 0) return tokens;
      } catch (e) { /* ignore invalid selectors */ }
    }
  }
  return [];
}

// ── Sender-anchor sign-off detection ─────────────────────────────────────────
//
// Scans the last portion of the email body bottom→top to find the sign-off
// and contact block. Returns the element to start collapsing from, or null.
//
// Patterns (checked bottom→top, first match wins):
//   A — "Best,\nMichael" in one element  → collapse from element AFTER
//   B — name element preceded by closing  → collapse from element AFTER name
//   C — "Michael T. Treadway, PhD\n..."  → collapse FROM this element
//       (only fires when all lines are short — not a paragraph starting with name)

function semaiFindSenderAnchor(body, senderName) {
  let scope = body;
  while (scope.children.length === 1 && SEMAI_BLOCK_TAGS.has(scope.children[0].tagName)) {
    scope = scope.children[0];
  }
  const kids = Array.from(scope.children);
  const startAt = Math.max(0, kids.length - 8);

  for (let i = kids.length - 1; i >= startAt; i--) {
    const raw = (kids[i].innerText || kids[i].textContent || "").trim();
    if (!raw) continue;
    const lines = raw.split(/\r?\n|\r/).map(l => l.trim()).filter(Boolean);
    const firstLine = lines[0] || "";

    // Pattern A: closing + name packed into one element via <br>
    if (lines.length === 2 && SEMAI_CLOSING_RE.test(lines[0])) {
      const nameWords = lines[1].split(/\s+/);
      const nameOk = nameWords.length <= 3 && nameWords.every(w => /^[A-Z]/.test(w));
      const nameMatches = !senderName || semaiFoldNameForMatch(lines[1]).startsWith(semaiFoldNameForMatch(senderName));
      if (nameOk && nameMatches && i + 1 < kids.length) {
        return kids[i + 1];
      }
    }

    // Pattern B: name element immediately preceded by a closing sibling
    if (lines.length === 1) {
      const words = raw.split(/\s+/);
      const isShortName = words.length >= 1 && words.length <= 3 && raw.length <= 30 && words.every(w => /^[A-Z]/.test(w));
      const nameMatches = !senderName || semaiFoldNameForMatch(raw).startsWith(semaiFoldNameForMatch(senderName));
      if (isShortName && nameMatches && i > 0) {
        const prevRaw = (kids[i - 1].innerText || kids[i - 1].textContent || "").trim();
        if (SEMAI_CLOSING_RE.test(prevRaw)) {
          return i + 1 < kids.length ? kids[i + 1] : null;
        }
      }
    }

    // Pattern C: sender's full name is the first line of a contact block
    if (!senderName) continue;
    const firstLineNameRemainder = semaiNameMatchRemainder(firstLine, senderName);
    if (firstLineNameRemainder !== null) {
      const charAfter = firstLineNameRemainder[0];
      if (!charAfter || /[\s,.]/.test(charAfter)) {
        const longLines = lines.filter(l => l.length > 120).length;
        semaiNativeLog(`[semai-sig] PatternC candidate: firstLine="${firstLine}" senderName="${senderName}" lines=${lines.length} longLines=${longLines}`);
        if (lines.length >= 2 && longLines <= 2) {
          semaiNativeLog(`[semai-sig] PatternC MATCH → collapsing from this element`);
          return kids[i];
        } else {
          semaiNativeLog(`[semai-sig] PatternC SKIP: lines=${lines.length} longLines=${longLines}`);
        }
      }
    }
  }
  return null;
}

// ── Mobile auto-signature detection ──────────────────────────────────────────

function semaiIsMobileSigEl(el) {
  const text = (el.innerText || el.textContent || "").trim();
  return SEMAI_MOBILE_SIG_RE.test(text);
}

// ── Main signature stripping entry point ─────────────────────────────────────
//
// Mutates the live reading-pane body element — collapses/hides the signature
// in place and adds a "Show signature" toggle button.
// Strategy order: branded-sig → Outlook sig div → sender anchor → RFC delimiter
//   → underscore separator → nested-div → repeated-name card → standalone card
//   → heuristic → mobile auto-sig

function semaiStripSignature(body) {
  if (body.dataset.semaiSigStripped) return;
  body.dataset.semaiOriginalHtml = body.innerHTML;
  body.dataset.semaiSigStripped = "true";

  const senderName = semaiGetSenderFirstName(body);
  const senderNameTokens = semaiGetSenderNameTokens(body);
  semaiNativeLog(`[semai-sig] stripSignature: senderName="${senderName}" tokens=${JSON.stringify(senderNameTokens)}`);

  // Strategy 0: Entire body is a compact branded signature
  if (semaiIsEntireBodySignature(body, senderName)) {
    semaiNativeLog(`[semai-sig] Strategy 0 (compact branded sig): collapsing entire body`);
    const firstChild = body.firstElementChild;
    if (firstChild) semaiCollapseFrom(firstChild);
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 0 skipped (body has real content)`);

  // Strategy 1: Outlook's labelled signature div
  const outlookSig = body.querySelector(
    '[id="Signature"], [id*="signature" i], [class*="signature" i]'
  );
  if (outlookSig && body.contains(outlookSig)) {
    semaiNativeLog(`[semai-sig] Strategy 1 (Outlook sig div): found element [${(outlookSig).tagName || ""}${(outlookSig).className ? "." + (outlookSig).className.toString().split(" ")[0] : ""}]`);
    const sepInSig =
      semaiFirstSeparatorEl(outlookSig, /^_{4,}$/) ||
      semaiFirstSeparatorEl(outlookSig, /^--\s*$/) ||
      outlookSig.querySelector("hr");
    if (sepInSig) {
      semaiCollapseFrom(sepInSig);
    } else {
      const anchor = semaiFindSenderAnchor(outlookSig, senderName);
      if (anchor) {
        semaiCollapseFrom(anchor);
      } else {
        semaiHideEl(outlookSig);
      }
    }
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 1 skipped (no id/class="signature" element)`);

  // Strategy 2: Sender name anchor (primary heuristic)
  const anchor = semaiFindSenderAnchor(body, senderName);
  if (anchor) {
    semaiNativeLog(`[semai-sig] Strategy 2 (sender anchor): collapsing from [${(anchor).tagName || ""}${(anchor).className ? "." + (anchor).className.toString().split(" ")[0] : ""}]`);
    semaiCollapseFrom(anchor);
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 2 skipped (no sender anchor found for "${senderName}")`);

  // Strategy 3: RFC "-- " delimiter
  const dashEl = semaiFirstSeparatorEl(body, /^--\s*$/);
  if (dashEl) {
    semaiNativeLog(`[semai-sig] Strategy 3 (-- delimiter): collapsing from [${(dashEl).tagName || ""}${(dashEl).className ? "." + (dashEl).className.toString().split(" ")[0] : ""}]`);
    semaiCollapseFrom(dashEl);
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 3 skipped (no -- delimiter)`);

  // Strategy 4: Outlook underscore separator
  const underscoreEl = semaiFirstSeparatorEl(body, /^_{4,}$/);
  if (underscoreEl) {
    semaiNativeLog(`[semai-sig] Strategy 4 (____ separator): collapsing from [${(underscoreEl).tagName || ""}${(underscoreEl).className ? "." + (underscoreEl).className.toString().split(" ")[0] : ""}]`);
    semaiCollapseFrom(underscoreEl);
    return;
  }

  // Strategy 5: Specific nested-div contact-card signature
  const nestedDivSig = semaiFindNestedDivSignature(body);
  if (nestedDivSig) {
    semaiNativeLog(`[semai-sig] Strategy 5 (nested div[dir=ltr]): hiding [${(nestedDivSig).tagName || ""}${(nestedDivSig).className ? "." + (nestedDivSig).className.toString().split(" ")[0] : ""}]`);
    semaiHideEl(nestedDivSig);
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 5 skipped (no nested-div signature found)`);

  // Strategy 5b: Repeated-name contact cards
  const nameTokensForDetection = senderNameTokens.length > 0 ? senderNameTokens : senderName;
  const repeatedNameAnchor = semaiFindRepeatedNameSignatureAnchor(body, nameTokensForDetection);
  if (repeatedNameAnchor) {
    semaiCollapseFrom(repeatedNameAnchor);
    semaiLog("[semai] Signature hidden via repeated-name anchor");
    return;
  }

  // Strategy 5c: Standalone contact card (name + credentials)
  const standaloneCard = semaiFindStandaloneContactCard(body, nameTokensForDetection);
  if (standaloneCard) {
    semaiCollapseFrom(standaloneCard);
    semaiLog("[semai] Signature hidden via standalone contact card");
    return;
  }

  // Strategy 5d: Table-based signature anchor (deeply nested contact tables)
  const tableAnchor = semaiFindTableSignatureAnchor(body, senderNameTokens);
  if (tableAnchor) {
    semaiNativeLog(`[semai-sig] Strategy 5d (table anchor): collapsing from [${(tableAnchor).tagName || ""}${(tableAnchor).className ? "." + (tableAnchor).className.toString().split(" ")[0] : ""}]`);
    semaiCollapseFrom(tableAnchor);
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 5d skipped (no table signature anchor found)`);

  // Strategy 6: Heuristic — contact-card block near the bottom
  const children = Array.from(body.children);
  semaiNativeLog(`[semai-sig] Strategy 6: checking last ${Math.min(6, children.length)} of ${children.length} body children`);
  let strategy6Hit = false;
  for (let i = children.length - 1; i >= Math.max(0, children.length - 6); i--) {
    if (semaiLooksLikeSig(children[i])) {
      semaiNativeLog(`[semai-sig] Strategy 6 (heuristic): collapsing from child[${i}] [${(children[i]).tagName || ""}${(children[i]).className ? "." + (children[i]).className.toString().split(" ")[0] : ""}]`);
      semaiCollapseFrom(children[i]);
      strategy6Hit = true;
      break;
    }
  }
  if (!strategy6Hit) semaiNativeLog(`[semai-sig] Strategy 6 skipped (no heuristic match)`);

  // Strategy 7: Mobile auto-signature lines
  const allChildren = Array.from(body.children);
  for (let i = allChildren.length - 1; i >= Math.max(0, allChildren.length - 5); i--) {
    if (semaiIsMobileSigEl(allChildren[i])) {
      semaiNativeLog(`[semai-sig] Strategy 7 (mobile auto-sig): collapsing from child[${i}]`);
      semaiCollapseFrom(allChildren[i]);
      break;
    }
  }
}

// ── Inline image resolution ───────────────────────────────────────────────────

function semaiIsTransparentPlaceholderImageSrc(src) {
  return typeof src === "string" && src.startsWith("data:image/gif;base64,R0lGODlhAQABAIA");
}

function semaiCloneImageNeedsResolution(img) {
  if (!(img instanceof HTMLImageElement)) return false;
  const imageType = img.getAttribute("data-imagetype") || "";
  const originalSrc = img.getAttribute("originalsrc") || "";
  const src = img.getAttribute("src") || "";
  return (
    /AttachmentByCid/i.test(imageType) ||
    /^cid:/i.test(originalSrc) ||
    semaiIsTransparentPlaceholderImageSrc(src)
  );
}

function semaiLiveImageHasRenderableSource(img) {
  if (!(img instanceof HTMLImageElement)) return false;
  const src = img.currentSrc || img.getAttribute("src") || "";
  if (!src || semaiIsTransparentPlaceholderImageSrc(src)) return false;
  return !/^cid:/i.test(src);
}

function semaiIsImageOnlyBlock(el) {
  if (!(el instanceof Element)) return false;
  const meaningfulChildren = Array.from(el.childNodes).filter(node => {
    if (node.nodeType === Node.TEXT_NODE) {
      return (node.textContent || "").replace(/\u00a0/g, " ").trim().length > 0;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return false;
    const child = node;
    if (child.tagName === "BR") return false;
    return true;
  });
  return meaningfulChildren.length > 0 && meaningfulChildren.every(node => (
    node.nodeType === Node.ELEMENT_NODE && node.tagName === "IMG"
  ));
}

function semaiRemoveUnresolvedImageBlock(img) {
  if (!(img instanceof HTMLImageElement)) return;
  let candidate = img;
  while (candidate.parentElement && semaiIsImageOnlyBlock(candidate.parentElement)) {
    candidate = candidate.parentElement;
  }
  candidate.remove();
}

function semaiResolveInlineImagesFromLiveDom(clone, bodyEl) {
  const cloneImages = Array.from(clone.querySelectorAll("img"));
  const liveImages = Array.from(bodyEl.querySelectorAll("img"));
  const imageCount = Math.min(cloneImages.length, liveImages.length);

  for (let index = 0; index < imageCount; index += 1) {
    const cloneImg = cloneImages[index];
    const liveImg = liveImages[index];

    if (!semaiCloneImageNeedsResolution(cloneImg)) continue;

    if (semaiLiveImageHasRenderableSource(liveImg)) {
      const resolvedSrc = liveImg.currentSrc || liveImg.getAttribute("src");
      cloneImg.setAttribute("src", resolvedSrc);
      const liveSrcset = liveImg.getAttribute("srcset");
      if (liveSrcset) {
        cloneImg.setAttribute("srcset", liveSrcset);
      } else {
        cloneImg.removeAttribute("srcset");
      }
      cloneImg.removeAttribute("originalsrc");
      continue;
    }

    semaiRemoveUnresolvedImageBlock(cloneImg);
  }
}

// ── Body clone cleaning ───────────────────────────────────────────────────────
//
// Returns a cleaned HTML string with signatures and quoted-reply blocks removed.
// senderFirstName: lowercase first name used to identify sign-off lines.

function semaiCleanBodyClone(bodyEl, senderFirstName) {
  const clone = document.createElement("div");
  if (bodyEl.dataset.semaiOriginalHtml) {
    clone.innerHTML = bodyEl.dataset.semaiOriginalHtml;
  } else {
    clone.innerHTML = bodyEl.innerHTML;
  }

  semaiResolveInlineImagesFromLiveDom(clone, bodyEl);

  semaiNativeLog(`[semai-sig] cleanBodyClone: senderFirstName="${senderFirstName}" bodyLen=${clone.textContent.length}`);
  if (semaiIsEntireBodySignature(clone, senderFirstName)) {
    semaiNativeLog(`[semai-sig] cleanBodyClone: Step 0a → entire body is a signature, returning empty`);
    return "";
  }

  // 0. Remove Outlook "external sender" warning banners
  clone.querySelectorAll('a[href*="LearnAboutSenderIdentification"]').forEach(a => {
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
  ).forEach(el => {
    let sib = el;
    while (sib) { const next = sib.nextElementSibling; sib.remove(); sib = next; }
  });

  // 2. Remove Outlook mobile reference messages (quoted replies)
  clone.querySelectorAll(
    '#mail-editor-reference-message-container, [class*="reference-message" i]'
  ).forEach(el => el.remove());

  // 3. Remove signature wrapper divs
  clone.querySelectorAll(
    '[id*="signature" i], [class*="signature" i]'
  ).forEach(el => el.remove());

  // 4. Remove quoted-reply header blocks (From / Date / Sent / To / Subject)
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
  if (anchor) {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 6 (sender anchor): removing from [${(anchor).tagName || ""}${(anchor).className ? "." + (anchor).className.toString().split(" ")[0] : ""}]`);
    removeFromAndAfter(anchor);
  } else {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 6: no sender anchor for "${senderFirstName}"`);
  }

  // 6b. Strip repeated-name contact cards
  const nameTokensForClone = Array.isArray(senderFirstName) ? senderFirstName
    : (senderFirstName ? [senderFirstName] : []);
  const cloneNameArg = nameTokensForClone.length > 0 ? nameTokensForClone : senderFirstName;
  const repeatedNameAnchor = semaiFindRepeatedNameSignatureAnchor(clone, cloneNameArg);
  if (repeatedNameAnchor) {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 6b (repeated-name anchor): removing from [${(repeatedNameAnchor).tagName || ""}]`);
    removeFromAndAfter(repeatedNameAnchor);
  } else {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 6b: no repeated-name anchor found`);
  }

  // 6c. Strip standalone contact cards (name + credentials, no sign-off)
  const standaloneCard = semaiFindStandaloneContactCard(clone, cloneNameArg);
  if (standaloneCard) removeFromAndAfter(standaloneCard);

  // 6c2. Strip closing + name + professional lines signatures with no contact fields
  const closingNameAnchor = semaiFindClosingNameSignatureAnchor(clone, cloneNameArg);
  if (closingNameAnchor) removeFromAndAfter(closingNameAnchor);

  // 6d. Strip deeply nested table-based signatures
  const tableAnchorClone = semaiFindTableSignatureAnchor(clone, cloneNameArg);
  if (tableAnchorClone) {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 6d (table anchor): removing from [${(tableAnchorClone).tagName || ""}]`);
    removeFromAndAfter(tableAnchorClone);
  }

  // 7. Strip specific nested-div contact-card signatures
  const nestedDivSig = semaiFindNestedDivSignature(clone);
  if (nestedDivSig) {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 7 (nested div[dir=ltr]): removing [${(nestedDivSig).tagName || ""}${(nestedDivSig).className ? "." + (nestedDivSig).className.toString().split(" ")[0] : ""}]`);
    nestedDivSig.remove();
  } else {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 7: no nested-div signature found`);
  }

  // 8. Strip closing phrase + name even when no contact block follows
  semaiStripTrailingSignOff(clone, senderFirstName);

  // 8b. Remove trailing legal disclaimers
  semaiStripTrailingDisclaimers(clone);

  // 10. Remove mobile auto-signature lines
  clone.querySelectorAll("p, div, span, td").forEach(el => {
    if (semaiIsMobileSigEl(el)) el.remove();
  });

  // 11. Remove trailing empty wrapper blocks left behind by Outlook markup.
  semaiTrimTrailingEmptyBlocks(clone);

  return clone.innerHTML.trim();
}

function semaiStripTrailingDisclaimers(container) {
  const DISCLAIMER_RE = /(privileged|confidential|protected health information|\bPHI\b|not the intended recipient|strictly prohibited|notify the sender by return e-mail)/i;

  function removeFromAndAfter(el) {
    let sib = el;
    while (sib) { const next = sib.nextElementSibling; sib.remove(); sib = next; }
  }

  const blocks = Array.from(container.querySelectorAll("div, p, span, td"));
  for (let i = blocks.length - 1; i >= 0; i--) {
    const text = (blocks[i].innerText || blocks[i].textContent || "").trim();
    if (!text) continue;
    if (!DISCLAIMER_RE.test(text)) continue;
    if (text.length < 120) continue;
    removeFromAndAfter(blocks[i]);
    return;
  }
}

function semaiStripQuotedReplyHeaders(container) {
  const HEADER_LINE_RE = /^(from|date|sent|to|cc|subject)\s*:/i;
  const HEADER_PAIR_RE = /(from|date|sent|to|cc|subject)\s*:.*\n.*(from|date|sent|to|cc|subject)\s*:/is;
  const WROTE_LINE_RE = /^on .+wrote:\s*$/i;

  function removeFromAndAfter(el) {
    let sib = el;
    while (sib) { const next = sib.nextElementSibling; sib.remove(); sib = next; }
  }

  const blocks = Array.from(container.querySelectorAll("div, p, blockquote, section, article, td"));
  for (const block of blocks) {
    const blockText = (block.innerText || block.textContent || "").trim();
    if (!blockText) continue;

    const lines = blockText.split(/\n+/).map((line) => line.trim()).filter(Boolean);
    const headerLines = lines.filter((line) => HEADER_LINE_RE.test(line));

    if (headerLines.length >= 2 || HEADER_PAIR_RE.test(blockText)) {
      removeFromAndAfter(block);
      return;
    }

    if (WROTE_LINE_RE.test(blockText)) {
      removeFromAndAfter(block);
      return;
    }
  }

  const walker = document.createTreeWalker(container, NodeFilter.SHOW_TEXT);
  let textNode;
  while ((textNode = walker.nextNode())) {
    const text = (textNode.textContent || "").trim();
    if (!HEADER_LINE_RE.test(text) && !WROTE_LINE_RE.test(text)) continue;

    let block = textNode.parentElement;
    while (block && block !== container && !SEMAI_BLOCK_TAGS.has(block.tagName)) {
      block = block.parentElement;
    }
    if (!block || block === container) continue;

    const blockText = (block.innerText || block.textContent || "").trim();
    const lines = blockText.split(/\n+/).map((line) => line.trim()).filter(Boolean);
    const headerLines = lines.filter((line) => HEADER_LINE_RE.test(line));
    if (headerLines.length >= 1 || lines.some((line) => WROTE_LINE_RE.test(line))) {
      removeFromAndAfter(block);
      return;
    }
  }

  const quoteBlocks = Array.from(container.querySelectorAll("blockquote"));
  for (const quoteBlock of quoteBlocks) {
    const previous = quoteBlock.previousElementSibling;
    const previousText = (previous?.innerText || previous?.textContent || "").trim();
    if (previous && WROTE_LINE_RE.test(previousText)) {
      previous.remove();
      quoteBlock.remove();
      return;
    }

    const parentText = (quoteBlock.parentElement?.innerText || "").trim();
    if (WROTE_LINE_RE.test(parentText.split(/\n+/)[0] || "")) {
      quoteBlock.parentElement?.remove();
      return;
    }
  }
}

// Remove a trailing "Best,\nName" or "Thanks,\nName" if it's the last
// content in the container — even if there's no contact card after it.
function semaiStripTrailingSignOff(container, senderFirstName) {
  let scope = container;
  while (scope.children.length === 1 && SEMAI_BLOCK_TAGS.has(scope.children[0].tagName)) {
    scope = scope.children[0];
  }
  const kids = Array.from(scope.children);
  if (kids.length === 0) return;

  let last = kids.length - 1;
  while (last >= 0 && !(kids[last].innerText || kids[last].textContent || "").trim()) last--;
  if (last < 0) return;

  const lastRaw = (kids[last].innerText || kids[last].textContent || "").trim();
  const lastLines = lastRaw.split(/\r?\n|\r/).map(l => l.trim()).filter(Boolean);

  // Pattern: single element with "Best,\nMichael"
  if (lastLines.length === 2 && SEMAI_CLOSING_RE.test(lastLines[0])) {
    const nameWords = lastLines[1].split(/\s+/);
    if (nameWords.length <= 3 && nameWords.every(w => /^[A-Z]/.test(w))) {
      if (!senderFirstName || semaiFoldNameForMatch(lastLines[1]).startsWith(semaiFoldNameForMatch(senderFirstName))) {
        kids[last].remove();
        return;
      }
    }
  }

  // Pattern: "Michael" preceded by "Best,"
  if (lastLines.length === 1 && last > 0) {
    const words = lastRaw.split(/\s+/);
    if (words.length <= 3 && lastRaw.length <= 30 && words.every(w => /^[A-Z]/.test(w))) {
      if (!senderFirstName || semaiFoldNameForMatch(lastRaw).startsWith(semaiFoldNameForMatch(senderFirstName))) {
        const prevRaw = (kids[last - 1].innerText || kids[last - 1].textContent || "").trim();
        if (SEMAI_CLOSING_RE.test(prevRaw)) {
          kids[last].remove();
          kids[last - 1].remove();
          return;
        }
      }
    }
  }

  // Pattern: just a closing phrase at the end (no name)
  if (lastLines.length === 1 && SEMAI_CLOSING_RE.test(lastRaw)) {
    kids[last].remove();
  }
}
