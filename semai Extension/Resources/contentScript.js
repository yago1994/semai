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
const SEMAI_CALIBRATION_STORAGE_KEY = "semaiSenderCalibration";

let semaiSavedSelection = null;
let semaiCalibrationState = null;

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
      <button class="semai-chat-toggle-btn" type="button" style="display:none">💬 Chat view</button>
      <button class="semai-calibrate-btn" type="button">Calibrate senders</button>
      <p id="semai-calibration-status" class="semai-calibration-status"></p>
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

    // Handle chat view toggle
    if (target.classList.contains("semai-chat-toggle-btn")) {
      if (semaiChatViewActive) {
        semaiDeactivateChatView();
      } else {
        semaiActivateChatView();
      }
      return;
    }

    if (target.classList.contains("semai-calibrate-btn")) {
      semaiStartCalibration();
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
  const calibration = semaiGetCalibration();
  semaiUpdateCalibrationStatus(
    calibration?.senderSelector
      ? "Sender calibration loaded."
      : "Calibration recommended: click your sender, then another sender."
  );
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
  // Save original HTML before we mutate it — used by chat view
  body.dataset.semaiOriginalHtml = body.innerHTML;
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

// ===== CHAT VIEW ============================================================

let semaiChatViewActive = false;
let semaiCurrentUser = null; // { name, email, initials }

// Deterministic avatar colour from name — 8-colour palette
const SEMAI_AVATAR_COLORS = [
  "#6366f1","#8b5cf6","#ec4899","#f59e0b","#10b981","#3b82f6","#ef4444","#14b8a6"
];
function semaiNameColor(name) {
  let h = 0;
  for (let i = 0; i < name.length; i++) h = name.charCodeAt(i) + ((h << 5) - h);
  return SEMAI_AVATAR_COLORS[Math.abs(h) % SEMAI_AVATAR_COLORS.length];
}

function semaiInitials(name) {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "?";
  if (parts.length === 1) return parts[0][0].toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
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

function semaiUpdateCalibrationStatus(message) {
  const status = document.getElementById("semai-calibration-status");
  if (status) status.textContent = message;
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
  semaiCalibrationState = null;
  semaiCurrentUser = null;
  semaiGetCurrentUser();
  semaiUpdateCalibrationStatus("Sender calibration saved.");
}

function semaiHandleCalibrationClick(event) {
  if (!semaiCalibrationState) return;

  const target = event.target instanceof Element ? event.target.closest("span, button, div") : null;
  if (!target) return;

  const text = (target.innerText || target.textContent || "").trim();
  if (!text || text.length > 160) return;

  event.preventDefault();
  event.stopPropagation();

  if (semaiCalibrationState.step === "self") {
    semaiCalibrationState.selfLabel = text;
    semaiCalibrationState.selector = semaiBuildSenderSelector(target);
    semaiCalibrationState.step = "other";
    semaiUpdateCalibrationStatus("Click another sender label in the thread.");
    return;
  }

  semaiFinishCalibration(
    semaiCalibrationState.selfLabel,
    text,
    semaiCalibrationState.selector || semaiBuildSenderSelector(target)
  );
  document.removeEventListener("click", semaiHandleCalibrationClick, true);
}

function semaiStartCalibration() {
  document.removeEventListener("click", semaiHandleCalibrationClick, true);
  semaiCalibrationState = {
    step: "self",
    selfLabel: "",
    selector: null
  };

  semaiUpdateCalibrationStatus("Click your sender label in the thread.");
  document.addEventListener("click", semaiHandleCalibrationClick, true);
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

function semaiGetSenderLabelNearBody(bodyEl) {
  const bodyContainer = bodyEl.closest('[data-test-id="mailMessageBodyContainer"]');
  if (!bodyContainer || !bodyContainer.parentElement) return null;
  const calibration = semaiGetCalibration();

  let sibling = bodyContainer.previousElementSibling;
  while (sibling) {
    if (calibration?.senderSelector) {
      if (sibling.matches?.(calibration.senderSelector)) {
        const text = (sibling.innerText || sibling.textContent || "").trim();
        if (text) return text;
      }

      const calibratedLabel = sibling.querySelector?.(calibration.senderSelector);
      if (calibratedLabel) {
        const text = (calibratedLabel.innerText || calibratedLabel.textContent || "").trim();
        if (text) return text;
      }
    }

    if (sibling.matches(".OZZZK")) {
      const text = (sibling.innerText || sibling.textContent || "").trim();
      if (text) return text;
    }

    const directLabel = sibling.querySelector?.(".OZZZK");
    if (directLabel) {
      const text = (directLabel.innerText || directLabel.textContent || "").trim();
      if (text) return text;
    }

    const text = (sibling.innerText || sibling.textContent || "").trim();
    if (text && text.length <= 120 && /^[A-Za-z]/.test(text)) {
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

  // ── Strategy 0: Config-defined name (most reliable) ──
  if (typeof SEMAI_USER_NAME === "string" && SEMAI_USER_NAME.trim().length >= 2) {
    trySet(SEMAI_USER_NAME);
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

  console.log("[semai] Current user detection failed — set SEMAI_USER_NAME in semaiConfig.js");
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
    if (normalized.name.length >= 2 && /^[A-Za-z]/.test(normalized.name)) {
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
        const raw = (found.getAttribute("aria-label") || found.innerText || found.textContent || "").trim();
        const normalized = semaiNormalizeSenderLabel(raw);
        if (normalized.name.length >= 2 && /^[A-Za-z]/.test(normalized.name)) {
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

// Clone a message body and strip signatures + quoted-reply blocks.
// senderFirstName: lowercase first name of this message's sender (for sign-off detection).
function semaiCleanBodyClone(bodyEl, senderFirstName) {
  const clone = document.createElement("div");
  if (bodyEl.dataset.semaiOriginalHtml) {
    clone.innerHTML = bodyEl.dataset.semaiOriginalHtml;
  } else {
    clone.innerHTML = bodyEl.innerHTML;
  }

  // ── 1. Remove Outlook reply/forward header blocks ──
  clone.querySelectorAll(
    '#divRplyFwdMsg, div[id*="divRplyFwdMsg"], div[id*="appendonsend"]'
  ).forEach(el => {
    let sib = el;
    while (sib) { const next = sib.nextElementSibling; sib.remove(); sib = next; }
  });

  // ── 2. Remove Outlook mobile reference messages (quoted replies) ──
  clone.querySelectorAll(
    '#mail-editor-reference-message-container, [class*="reference-message" i]'
  ).forEach(el => el.remove());

  // ── 3. Remove signature wrapper divs ──
  clone.querySelectorAll(
    '[id*="signature" i], [class*="signature" i]'
  ).forEach(el => el.remove());

  // ── 4. Remove "From:/Sent:/To:" quoted-reply text blocks ──
  const walker = document.createTreeWalker(clone, NodeFilter.SHOW_TEXT);
  let textNode;
  while ((textNode = walker.nextNode())) {
    if (/^from:\s/im.test(textNode.textContent)) {
      let block = textNode.parentElement;
      while (block && block !== clone && !SEMAI_BLOCK_TAGS.has(block.tagName)) {
        block = block.parentElement;
      }
      if (block && block !== clone) {
        const blockText = (block.innerText || "").trim();
        if (/^from:\s/im.test(blockText) && /sent:\s/im.test(blockText)) {
          let sib = block;
          while (sib) { const next = sib.nextElementSibling; sib.remove(); sib = next; }
          break;
        }
      }
    }
  }

  // ── 5. Strip separator lines (-- , ____ , <hr>) and everything after ──
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

  // ── 6. Strip sign-off + sender name contact block ──
  // Uses the same anchor logic as the reading-view signature stripper.
  const anchor = semaiFindSenderAnchor(clone, senderFirstName);
  if (anchor) removeFromAndAfter(anchor);

  // ── 7. Strip closing phrase + name even when no contact block follows ──
  // e.g. "Best,\nMichael" at the very end — remove the closing and name too.
  semaiStripTrailingSignOff(clone, senderFirstName);

  // ── 8. Remove trailing empty elements ──
  while (clone.lastElementChild) {
    const last = clone.lastElementChild;
    if (!(last.innerText || last.textContent || "").trim()) {
      last.remove();
    } else {
      break;
    }
  }

  return clone.innerHTML.trim();
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

  // Check the last non-empty element
  let last = kids.length - 1;
  while (last >= 0 && !(kids[last].innerText || kids[last].textContent || "").trim()) last--;
  if (last < 0) return;

  const lastRaw = (kids[last].innerText || kids[last].textContent || "").trim();
  const lastLines = lastRaw.split(/\r?\n|\r/).map(l => l.trim()).filter(Boolean);

  // Pattern: single element with "Best,\nMichael"
  if (lastLines.length === 2 && SEMAI_CLOSING_RE.test(lastLines[0])) {
    const nameWords = lastLines[1].split(/\s+/);
    if (nameWords.length <= 3 && nameWords.every(w => /^[A-Z]/.test(w))) {
      if (!senderFirstName || lastLines[1].toLowerCase().startsWith(senderFirstName)) {
        kids[last].remove();
        return;
      }
    }
  }

  // Pattern: "Michael" preceded by "Best,"
  if (lastLines.length === 1 && last > 0) {
    const words = lastRaw.split(/\s+/);
    if (words.length <= 3 && lastRaw.length <= 30 && words.every(w => /^[A-Z]/.test(w))) {
      if (!senderFirstName || lastRaw.toLowerCase().startsWith(senderFirstName)) {
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

// Detect if a message is from the current user
function semaiIsCurrentUser(senderName, senderEmail) {
  const user = semaiGetCurrentUser();
  if (!user) return false;

  // Email match (strongest signal)
  if (user.email && senderEmail && senderEmail.toLowerCase() === user.email) return true;

  const sLower = senderName.toLowerCase().trim();
  if (!sLower) return false;

  // Exact full name match
  if (sLower === user.nameLower) return true;

  // Last name match as fallback (Outlook sometimes shows "Lastname, Firstname")
  const userParts = user.nameLower.split(/\s+/);
  if (userParts.length >= 2 && sLower.includes(",")) {
    const userLast = userParts[userParts.length - 1];
    const userFirst = userParts[0];
    if (sLower.startsWith(userLast + ",") && sLower.includes(userFirst)) return true;
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
    const senderFirstName = (sender.name.split(/\s+/)[0] || "").toLowerCase();
    const cleanHtml = semaiCleanBodyClone(bodyEl, senderFirstName);
    const isMe = semaiIsCurrentUser(sender.name, sender.email);
    return { sender, timestamp, cleanHtml, isMe };
  });
}

// Get the conversation subject
function semaiGetThreadSubject() {
  const selectors = [
    '[data-testid="subjectLine"]',
    '[class*="subjectLine" i]',
    'h2[class*="subject" i]',
    '[role="heading"][class*="subject" i]',
  ];
  for (const sel of selectors) {
    const el = document.querySelector(sel);
    if (el) {
      const text = (el.innerText || el.textContent || "").trim();
      if (text) return text;
    }
  }
  return document.title.replace(/- Outlook.*$/i, "").trim() || "Conversation";
}

// Build the chat overlay DOM
function semaiCreateChatOverlay(messages, subject) {
  const overlay = document.createElement("div");
  overlay.id = "semai-chat-overlay";

  // Header bar
  const header = document.createElement("div");
  header.className = "semai-chat-header";
  header.innerHTML = `
    <span class="semai-chat-subject">${subject.replace(/</g, "&lt;")}</span>
    <button class="semai-chat-close" type="button">✕ Exit chat view</button>
  `;
  header.querySelector(".semai-chat-close").addEventListener("click", semaiDeactivateChatView);
  overlay.appendChild(header);

  // Scrollable message area
  const scroll = document.createElement("div");
  scroll.className = "semai-chat-scroll";

  messages.forEach((msg) => {
    if (!msg.cleanHtml) return; // skip empty bodies

    const row = document.createElement("div");
    row.className = `semai-chat-row ${msg.isMe ? "semai-chat-me" : "semai-chat-them"}`;

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

    if (msg.isMe) {
      row.appendChild(bubble);
      row.appendChild(avatar);
    } else {
      row.appendChild(avatar);
      row.appendChild(bubble);
    }
    scroll.appendChild(row);
  });

  overlay.appendChild(scroll);
  return overlay;
}

// Find the Outlook reading pane container (scrollable ancestor of message bodies)
function semaiGetReadingPane() {
  const firstBody = document.querySelector('[aria-label="Message body"]:not([contenteditable])');
  if (!firstBody) return null;

  let el = firstBody.parentElement;
  for (let d = 0; d < 25 && el; d++, el = el.parentElement) {
    if (el === document.body || el === document.documentElement) break;
    const style = window.getComputedStyle(el);
    const overflow = style.overflowY;
    if ((overflow === "auto" || overflow === "scroll") && el.clientHeight > 200) {
      return el;
    }
  }
  // Fallback: role="main" or the complementary region
  return document.querySelector('[role="main"]') ||
    document.querySelector('[role="complementary"]') || null;
}

function semaiActivateChatView() {
  if (semaiChatViewActive) return;

  // Ensure we have the current user
  if (!semaiGetCurrentUser()) {
    alert("semai couldn't identify your account.\nSet SEMAI_USER_NAME in semaiConfig.js.");
    return;
  }

  const messages = semaiExtractThreadMessages();
  if (messages.length < 2) {
    alert("semai chat view needs a thread with at least 2 messages.");
    return;
  }

  const subject = semaiGetThreadSubject();
  const overlay = semaiCreateChatOverlay(messages, subject);

  // Contain the overlay within the reading pane
  const readingPane = semaiGetReadingPane();
  if (readingPane) {
    const rpStyle = window.getComputedStyle(readingPane);
    if (rpStyle.position === "static") readingPane.style.position = "relative";
    readingPane.appendChild(overlay);
  } else {
    document.body.appendChild(overlay);
  }

  semaiChatViewActive = true;
  semaiUpdateChatToggleBtn();

  // Scroll to bottom after the overlay is in the DOM and painted
  const scrollEl = overlay.querySelector(".semai-chat-scroll");
  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      scrollEl.scrollTop = scrollEl.scrollHeight;
    });
  });

  semaiLog("[semai] Chat view activated", { messageCount: messages.length });
}

function semaiDeactivateChatView() {
  const overlay = document.getElementById("semai-chat-overlay");
  if (overlay) overlay.remove();
  semaiChatViewActive = false;
  semaiUpdateChatToggleBtn();
  semaiLog("[semai] Chat view deactivated");
}

function semaiUpdateChatToggleBtn() {
  const btn = document.querySelector(".semai-chat-toggle-btn");
  if (!btn) return;
  btn.textContent = semaiChatViewActive ? "✕ Exit chat view" : "💬 Chat view";
}

// Show/hide the chat toggle based on whether we're looking at a thread
function semaiUpdateChatToggleVisibility() {
  const btn = document.querySelector(".semai-chat-toggle-btn");
  if (!btn) return;
  const bodies = document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])');
  btn.style.display = bodies.length >= 2 ? "" : "none";
}

// Auto-deactivate when Outlook navigates to a different email
let semaiLastReadingPaneSignature = "";
function semaiWatchForNavigation() {
  const check = () => {
    const bodies = document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])');
    const sig = Array.from(bodies).map(b => b.dataset.semaiSigStripped || "").join("|");
    if (semaiChatViewActive && sig !== semaiLastReadingPaneSignature) {
      semaiDeactivateChatView();
    }
    semaiLastReadingPaneSignature = sig;
    semaiUpdateChatToggleVisibility();
  };

  const obs = new MutationObserver(check);
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
  document.addEventListener("selectionchange", semaiSaveSelectionFromCompose);

  if (!semaiGetCalibration()) {
    window.setTimeout(() => {
      if (!semaiGetCalibration() && !semaiCalibrationState) {
        semaiStartCalibration();
      }
    }, 1200);
  }

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
