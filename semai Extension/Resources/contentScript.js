// ===== UTIL: find the compose/body element =====
function getComposeElement() {
  const candidates = [
    'div[aria-label="Message body"][contenteditable="true"]',
    'div[role="textbox"][contenteditable="true"]'
  ];

  for (const sel of candidates) {
    const el = document.querySelector(sel);
    if (el) return el;
  }

  const fallbacks = document.querySelectorAll(
    'div[role="textbox"][contenteditable="true"]'
  );
  if (fallbacks.length > 0) {
    return fallbacks[fallbacks.length - 1];
  }

  return null;
}

const SEMAI_DEBUG = false;

let semaiSavedSelection = null;

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

  semaiLog("[semai] Panel toggled", { collapsed: isCollapsed });
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
        <div class="semai-title">semai</div>
      </div>
      <button
        class="semai-toggle-btn"
        type="button"
        aria-label="Collapse semai"
      >
        ▴
      </button>
    </div>
    <div class="semai-body">
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
    </div>
  `;

  panel.addEventListener("click", (e) => {
    const target = e.target;
    if (!(target instanceof HTMLButtonElement)) return;

    // Handle collapse/expand toggle
    if (target.classList.contains("semai-toggle-btn")) {
      toggleSemaiPanel();
      return;
    }

    const mode = target.dataset.mode;
    if (!mode) return;

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
  semaiLog("[semai] Panel created");
}

// ===== SIGNATURE STRIPPING (reading view) =====

const SEMAI_SIG_SHORT_LINE_MAX = 60;
const SEMAI_SIG_MIN_LINES = 5;
const SEMAI_PHONE_RE = /\+?[\d][\d\s\-\.\(\)]{6,}/;
const SEMAI_URL_RE = /https?:\/\/|www\./i;
const SEMAI_SOCIAL_RE = /linkedin\.com|twitter\.com|facebook\.com|instagram\.com/i;
const SEMAI_BLOCK_TAGS = new Set([
  "P","DIV","BLOCKQUOTE","LI","H1","H2","H3","H4","H5","H6",
  "SECTION","ARTICLE","HEADER","FOOTER","MAIN","TABLE","TR","TD","TH"
]);
const SEMAI_CLOSING_RE = /^(best|regards|thanks|thank you|cheers|sincerely|warmly|warm regards|kind regards|best regards|yours|cordially|take care|many thanks|with gratitude|respectfully|yours truly|talk soon|looking forward)[,.]?\s*$/i;

function semaiLooksLikeSig(el) {
  const text = el.innerText || el.textContent || "";
  const lines = text.split("\n").map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length < SEMAI_SIG_MIN_LINES) return false;
  const shortLines = lines.filter(l => l.length <= SEMAI_SIG_SHORT_LINE_MAX).length;
  if (shortLines / lines.length < 0.8) return false;
  return SEMAI_PHONE_RE.test(text) || SEMAI_URL_RE.test(text) || SEMAI_SOCIAL_RE.test(text);
}

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

// Wraps startEl and all its following siblings (within their shared parent)
// in a hidden div and inserts a toggle button before it.
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

// Walk text nodes within a container and return the nearest block-level element
// whose entire trimmed text content matches pattern.
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

// Try to extract the sender's first name from Outlook's email header UI.
// Casts a wide net of selectors then walks up to 12 ancestor levels looking
// for known sender-name elements. Returns a lowercase first name or null.
function semaiGetSenderFirstName(bodyEl) {
  const selectors = [
    // Outlook Web (office.com / office365.com) — Fluent UI persona
    '.ms-Persona-primaryText',
    '[class*="personaName" i]',
    '[class*="persona-name" i]',
    // Outlook Web — reading pane header
    '[data-testid="senderName"]',
    '[data-testid="sender-name"]',
    '[class*="senderName" i]',
    '[class*="sender-name" i]',
    '[class*="fromAddress" i]',
    // Aria labels on contact buttons ("Michael T. Treadway")
    'button[aria-label]:not([aria-label=""])',
  ];

  let ancestor = bodyEl.parentElement;
  for (let d = 0; d < 12 && ancestor; d++, ancestor = ancestor.parentElement) {
    for (const sel of selectors) {
      try {
        const found = ancestor.querySelector(sel);
        if (!found || found.contains(bodyEl)) continue;
        const raw = (found.getAttribute("aria-label") || found.innerText || found.textContent || "").trim();
        // Strip "From: " prefix if present
        const cleaned = raw.replace(/^from[:\s]+/i, "");
        const firstName = cleaned.split(/[\s,<(@]+/)[0];
        if (firstName && firstName.length >= 2 && /^[A-Za-z]/.test(firstName)) {
          return firstName.toLowerCase();
        }
      } catch (e) { /* ignore invalid selectors */ }
    }
  }
  return null;
}

// Primary sender-name anchor strategy.
// Scans the last portion of the email body BOTTOM→TOP so it always finds
// the most-recent (innermost) sign-off and never catches an earlier
// "Thanks, John" from a quoted reply higher up in the thread.
//
// Returns the element to START collapsing from, or null.
//
// Patterns (checked bottom→top, first match wins):
//   A — "Best,\nMichael" in one element  → collapse from element AFTER
//   B — name element preceded by closing  → collapse from element AFTER name
//   C — "Michael T. Treadway, PhD\n..."  → collapse FROM this element
//       (only fires when all lines are short — not a paragraph starting with name)
//
// Without a sender name, Patterns A & B still fire on any closing+capitalised-name
// pair. Pattern C always requires a known sender name.
function semaiFindSenderAnchor(body, senderName) {
  let scope = body;
  while (scope.children.length === 1 && SEMAI_BLOCK_TAGS.has(scope.children[0].tagName)) {
    scope = scope.children[0];
  }
  const kids = Array.from(scope.children);
  // Only look at the last 8 elements — signatures are always near the bottom
  const startAt = Math.max(0, kids.length - 8);

  // ── Sweep bottom→top ──
  for (let i = kids.length - 1; i >= startAt; i--) {
    const raw = (kids[i].innerText || kids[i].textContent || "").trim();
    if (!raw) continue;
    const lines = raw.split(/\r?\n|\r/).map(l => l.trim()).filter(Boolean);
    const firstLine = lines[0] || "";

    // ── Pattern A: closing + name packed into one element via <br> ──
    // e.g. <p>Best,<br>Michael</p>
    if (lines.length === 2 && SEMAI_CLOSING_RE.test(lines[0])) {
      const nameWords = lines[1].split(/\s+/);
      const nameOk = nameWords.length <= 3 && nameWords.every(w => /^[A-Z]/.test(w));
      const nameMatches = !senderName || lines[1].toLowerCase().startsWith(senderName);
      if (nameOk && nameMatches && i + 1 < kids.length) {
        return kids[i + 1];
      }
    }

    // ── Pattern B: name element immediately preceded by a closing sibling ──
    // e.g. <p>Best,</p><p>Michael</p>
    if (lines.length === 1) {
      const words = raw.split(/\s+/);
      const isShortName = words.length >= 1 && words.length <= 3 && raw.length <= 30 && words.every(w => /^[A-Z]/.test(w));
      const nameMatches = !senderName || raw.toLowerCase().startsWith(senderName);
      if (isShortName && nameMatches && i > 0) {
        const prevRaw = (kids[i - 1].innerText || kids[i - 1].textContent || "").trim();
        if (SEMAI_CLOSING_RE.test(prevRaw)) {
          // Return the element right after the name — that's where the contact card starts
          return i + 1 < kids.length ? kids[i + 1] : null;
        }
      }
    }

    // ── Pattern C: sender's full name is the first line of a contact block ──
    // e.g. <p>Michael T. Treadway, PhD<br>Winship Distinguished...<br>...</p>
    // Guard: all lines must be short so we don't match a paragraph that happens
    // to start with the sender's name.
    if (!senderName) continue;
    if (firstLine.toLowerCase().startsWith(senderName)) {
      const charAfter = firstLine[senderName.length];
      if (!charAfter || /[\s,.]/.test(charAfter)) {
        if (lines.length >= 2 && lines.every(l => l.length <= 80)) {
          return kids[i];
        }
      }
    }
  }
  return null;
}

function semaiStripSignature(body) {
  if (body.dataset.semaiSigStripped) return;
  body.dataset.semaiSigStripped = "true";

  // Fetch sender name once — used by multiple strategies below
  const senderName = semaiGetSenderFirstName(body);
  semaiLog("[semai] Sender name", { senderName });

  // ── Strategy 1: Outlook's labelled signature div ──────────────────────
  // Handles <div id="ms-outlook-mobile-signature">, <div id="Signature">, etc.
  // Looks inside the div for a separator so sign-off lines stay visible.
  const outlookSig = body.querySelector(
    '[id="Signature"], [id*="signature" i], [class*="signature" i]'
  );
  if (outlookSig && body.contains(outlookSig)) {
    const sepInSig =
      semaiFirstSeparatorEl(outlookSig, /^_{4,}$/) ||
      semaiFirstSeparatorEl(outlookSig, /^--\s*$/) ||
      outlookSig.querySelector("hr");
    if (sepInSig) {
      semaiCollapseFrom(sepInSig);
    } else {
      // No separator inside the sig div — use sender name anchor if available
      const anchor = semaiFindSenderAnchor(outlookSig, senderName);
      if (anchor) {
        semaiCollapseFrom(anchor);
      } else {
        semaiHideEl(outlookSig);
      }
    }
    semaiLog("[semai] Signature hidden via Outlook sig div");
    return;
  }

  // ── Strategy 2: Sender name anchor (primary heuristic) ───────────────
  // Uses the sender's first name from Outlook's UI to locate their name
  // in the body as a sign-off or contact-block header.
  const anchor = semaiFindSenderAnchor(body, senderName);
  if (anchor) {
    semaiCollapseFrom(anchor);
    semaiLog("[semai] Signature hidden via sender name anchor", { senderName });
    return;
  }

  // ── Strategy 3: RFC "-- " delimiter ──────────────────────────────────
  const dashEl = semaiFirstSeparatorEl(body, /^--\s*$/);
  if (dashEl) {
    semaiCollapseFrom(dashEl);
    semaiLog("[semai] Signature hidden via -- delimiter");
    return;
  }

  // ── Strategy 4: Outlook underscore separator ──────────────────────────
  // <hr> is intentionally excluded: it's also used between quoted replies
  // in email threads and causes entire conversations to be hidden.
  const underscoreEl = semaiFirstSeparatorEl(body, /^_{4,}$/);
  if (underscoreEl) {
    semaiCollapseFrom(underscoreEl);
    semaiLog("[semai] Signature hidden via underscore separator");
    return;
  }

  // ── Strategy 5: Heuristic — contact-card block near the bottom ───────
  // Last resort: a block with 5+ short lines containing phone/URL/social.
  const children = Array.from(body.children);
  for (let i = children.length - 1; i >= Math.max(0, children.length - 6); i--) {
    if (semaiLooksLikeSig(children[i])) {
      semaiCollapseFrom(children[i]);
      semaiLog("[semai] Signature hidden via heuristic", { childIndex: i });
      break;
    }
  }
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
  document.addEventListener("selectionchange", semaiSaveSelectionFromCompose);

  const observer = new MutationObserver(() => {
    if (!document.getElementById("semai-panel")) {
      semaiLog("[semai] Panel missing, recreating");
      createPanel();
    }
  });

  observer.observe(document.documentElement, {
    childList: true,
    subtree: true
  });
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", setupWhenReady, { once: true });
} else {
  setupWhenReady();
}
