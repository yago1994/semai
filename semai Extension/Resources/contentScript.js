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

function semaiIsVisibleElement(el) {
  if (!(el instanceof Element)) return false;
  const rect = el.getBoundingClientRect();
  const style = window.getComputedStyle(el);
  return rect.width > 0 && rect.height > 0 && style.visibility !== "hidden" && style.display !== "none";
}

function semaiWaitForComposeElement(timeoutMs = 6000) {
  return new Promise((resolve, reject) => {
    const startedAt = Date.now();

    const check = () => {
      const composeEl = getComposeElement();
      if (composeEl) {
        resolve(composeEl);
        return;
      }

      if (Date.now() - startedAt >= timeoutMs) {
        reject(new Error("Outlook reply box did not open in time."));
        return;
      }

      window.setTimeout(check, 120);
    };

    check();
  });
}

function semaiFindReplyAllButton() {
  const selector = [
    'button[aria-label*="Reply all" i]',
    '[role="button"][aria-label*="Reply all" i]',
    'button[title*="Reply all" i]',
    '[role="button"][title*="Reply all" i]',
    '[data-testid*="replyall" i]',
    '[name*="replyall" i]'
  ].join(", ");

  const matches = Array.from(document.querySelectorAll(selector)).filter(semaiIsVisibleElement);
  if (matches.length > 0) return matches[matches.length - 1];

  const textMatches = Array.from(document.querySelectorAll('button, [role="button"]'))
    .filter(semaiIsVisibleElement)
    .filter((el) => /reply all/i.test(el.getAttribute("aria-label") || el.textContent || ""));

  return textMatches[textMatches.length - 1] || null;
}

function semaiFindSendButton() {
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

  const matches = Array.from(document.querySelectorAll(selector))
    .filter(semaiIsVisibleElement)
    .filter((el) => {
      const label = el.getAttribute("aria-label") || "";
      const title = el.getAttribute("title") || "";
      return !/send to/i.test(label) && !/schedule/i.test(label) && !/schedule/i.test(title);
    });

  if (matches.length > 0) return matches[matches.length - 1];

  const textMatches = Array.from(document.querySelectorAll('button, [role="button"]'))
    .filter(semaiIsVisibleElement)
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

async function semaiOpenReplyAllCompose() {
  let composeEl = getComposeElement();
  if (composeEl) return composeEl;

  const replyAllBtn = semaiFindReplyAllButton();
  if (!replyAllBtn) {
    throw new Error("Reply all button not found in Outlook.");
  }

  replyAllBtn.click();
  composeEl = await semaiWaitForComposeElement();
  return composeEl;
}

function semaiInsertComposeText(composeEl, text) {
  composeEl.focus();

  const lines = text.split(/\n/);
  const fragment = document.createDocumentFragment();

  lines.forEach((line, index) => {
    const block = document.createElement("div");
    if (line) {
      block.textContent = line;
    } else {
      block.appendChild(document.createElement("br"));
    }
    fragment.appendChild(block);

    if (index === lines.length - 1 && !line) {
      block.appendChild(document.createElement("br"));
    }
  });

  composeEl.innerHTML = "";
  composeEl.appendChild(fragment);
  composeEl.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText", data: text }));
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
    semaiInsertComposeText(composeEl, draft);

    if (status) status.textContent = "Reply all draft inserted into Outlook.";
  } catch (err) {
    if (status) status.textContent = err.message;
  } finally {
    if (draftBtn) draftBtn.disabled = false;
    if (sendBtn) sendBtn.disabled = false;
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
    const composeEl = await semaiOpenReplyAllCompose();
    semaiInsertComposeText(composeEl, draft);

    const sendButton = semaiFindSendButton();
    if (sendButton) {
      sendButton.click();
    } else {
      semaiTriggerComposeSend(composeEl);
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

const SEMAI_DEBUG = false;
const SEMAI_AI_AGENT_ENABLED = false;
const SEMAI_CALIBRATION_STORAGE_KEY = "semaiSenderCalibration";
const SEMAI_PANEL_POSITION_STORAGE_KEY = "semaiPanelPosition";

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
      <button
        class="semai-toggle-btn"
        type="button"
        aria-label="Collapse REMOU"
      >
        ▴
      </button>
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
  semaiUpdateCalibrationStatus(
    calibration?.senderSelector
      ? "Sender detection is trained for this Outlook layout."
      : "Quick setup: click your name in a thread, then click another sender.",
    calibration?.senderSelector ? "success" : "neutral"
  );
  semaiLog("[semai] Panel created");
}

// ===== SIGNATURE STRIPPING (reading view) =====

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

function semaiLooksLikeSig(el) {
  const text = el.innerText || el.textContent || "";
  const lines = text.split("\n").map(l => l.trim()).filter(l => l.length > 0);
  if (lines.length < SEMAI_SIG_MIN_LINES) return false;
  const shortLines = lines.filter(l => l.length <= SEMAI_SIG_SHORT_LINE_MAX).length;
  if (shortLines / lines.length < 0.8) return false;
  return SEMAI_PHONE_RE.test(text) || SEMAI_URL_RE.test(text) || SEMAI_SOCIAL_RE.test(text);
}

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

function semaiNormalizeNameLine(text) {
  return (text || "")
    .replace(/[^\p{L}\s.'-]/gu, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
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

  if (SEMAI_PHONE_RE.test(raw)) signalKinds.add("phone");
  if (SEMAI_EMAIL_RE.test(raw)) signalKinds.add("email");
  if (SEMAI_URL_RE.test(raw) || SEMAI_SOCIAL_RE.test(raw)) signalKinds.add("url");
  if (SEMAI_SIGNATURE_TITLE_RE.test(raw)) signalKinds.add("title");
  if (SEMAI_SIGNATURE_ADDRESS_RE.test(raw)) signalKinds.add("address");

  return signalKinds.size;
}

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

// Detect when the entire body is a compact branded signature with no actual message.
// Returns true if the content starts with the sender's name and has URL/image signals
// but no long-sentence body text — i.e., the whole email IS the signature.
function semaiIsEntireBodySignature(clone, senderFirstName) {
  if (!senderFirstName) return false;

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
  if (!firstThreeText.includes(senderFirstName.toLowerCase())) return false;

  // Must have at least one URL/image signal in the raw HTML
  const hasUrl = SEMAI_URL_RE.test(text) || !!clone.querySelector("img, a[href]");
  if (!hasUrl) return false;

  // Must NOT have any long sentence (real message body would have sentences > 60 chars)
  const hasLongSentence = lines.some(l => l.length > 60 && /\s/.test(l));
  if (hasLongSentence) return false;

  return true;
}

function semaiFindRepeatedNameSignatureAnchor(container, senderFirstName) {
  const blocks = Array.from(container.querySelectorAll("p, div, table, td"))
    .filter((el) => semaiLooksLikeCompactBlock(el));
  if (blocks.length < 3) return null;

  const startAt = Math.max(0, blocks.length - 24);
  for (let i = startAt; i < blocks.length; i++) {
    const currentText = (blocks[i].innerText || blocks[i].textContent || "").trim();
    if (!semaiLooksLikeNameLine(currentText, senderFirstName)) continue;

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
        if (semaiLooksLikeContactCardBlock(blocks[k])) {
          contactCardSeen = true;
        }
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

// Extracts the true first name from a display name in either "First Last" or
// "Last, First" format (the latter is common in Outlook when contacts are
// stored with surname-first ordering, e.g. "Lam, Wilbur").
function semaiFirstNameFromDisplayName(displayName) {
  const name = (displayName || "").trim();
  if (/,/.test(name)) {
    // "Last, First [Middle]" → take the token right after the comma
    const afterComma = (name.split(/\s*,\s*/)[1] || "").split(/\s+/)[0];
    if (afterComma && afterComma.length >= 2 && /^[A-Za-z]/.test(afterComma)) {
      semaiNativeLog(`[semai-sig] firstNameFromDisplay: "${name}" → "Last, First" format → "${afterComma.toLowerCase()}"`);
      return afterComma.toLowerCase();
    }
  }
  // "First [Middle] Last" → take the first whitespace-separated token
  const first = (name.split(/[\s,<(@]+/)[0] || "").toLowerCase();
  semaiNativeLog(`[semai-sig] firstNameFromDisplay: "${name}" → "First Last" format → "${first}"`);
  return first;
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
        const firstName = semaiFirstNameFromDisplayName(cleaned);
        semaiNativeLog(`[semai-sig] getSenderFirstName: selector="${sel}" raw="${raw}" → firstName="${firstName}"`);
        if (firstName && firstName.length >= 2 && /^[A-Za-z]/.test(firstName)) {
          semaiNativeLog(`[semai-sig] getSenderFirstName: resolved to "${firstName}"`);
          return firstName;
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
    // Guard: lines must not look like a prose paragraph.  We allow lines up to
    // 120 chars (academic/medical dept names can exceed 80) but cap at 2 long
    // lines — a real paragraph would have far more text.
    if (!senderName) continue;
    if (firstLine.toLowerCase().startsWith(senderName)) {
      const charAfter = firstLine[senderName.length];
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

function semaiStripSignature(body) {
  if (body.dataset.semaiSigStripped) return;
  // Save original HTML before we mutate it — used by chat view
  body.dataset.semaiOriginalHtml = body.innerHTML;
  body.dataset.semaiSigStripped = "true";

  // Fetch sender name once — used by multiple strategies below
  const senderName = semaiGetSenderFirstName(body);
  semaiNativeLog(`[semai-sig] stripSignature: senderName="${senderName}"`);

  // ── Strategy 0: Entire body is a compact branded signature ──────────────
  // Handles emails that have no message body at all — just name/title/logo.
  // Must run after saving semaiOriginalHtml but before any DOM mutation.
  if (semaiIsEntireBodySignature(body, senderName)) {
    semaiNativeLog(`[semai-sig] Strategy 0 (compact branded sig): collapsing entire body`);
    const firstChild = body.firstElementChild;
    if (firstChild) semaiCollapseFrom(firstChild);
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 0 skipped (body has real content)`);

  // ── Strategy 1: Outlook's labelled signature div ──────────────────────
  // Handles <div id="ms-outlook-mobile-signature">, <div id="Signature">, etc.
  // Looks inside the div for a separator so sign-off lines stay visible.
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
      // No separator inside the sig div — use sender name anchor if available
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

  // ── Strategy 2: Sender name anchor (primary heuristic) ───────────────
  // Uses the sender's first name from Outlook's UI to locate their name
  // in the body as a sign-off or contact-block header.
  const anchor = semaiFindSenderAnchor(body, senderName);
  if (anchor) {
    semaiNativeLog(`[semai-sig] Strategy 2 (sender anchor): collapsing from [${(anchor).tagName || ""}${(anchor).className ? "." + (anchor).className.toString().split(" ")[0] : ""}]`);
    semaiCollapseFrom(anchor);
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 2 skipped (no sender anchor found for "${senderName}")`);

  // ── Strategy 3: RFC "-- " delimiter ──────────────────────────────────
  const dashEl = semaiFirstSeparatorEl(body, /^--\s*$/);
  if (dashEl) {
    semaiNativeLog(`[semai-sig] Strategy 3 (-- delimiter): collapsing from [${(dashEl).tagName || ""}${(dashEl).className ? "." + (dashEl).className.toString().split(" ")[0] : ""}]`);
    semaiCollapseFrom(dashEl);
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 3 skipped (no -- delimiter)`);

  // ── Strategy 4: Outlook underscore separator ──────────────────────────
  // <hr> is intentionally excluded: it's also used between quoted replies
  // in email threads and causes entire conversations to be hidden.
  const underscoreEl = semaiFirstSeparatorEl(body, /^_{4,}$/);
  if (underscoreEl) {
    semaiNativeLog(`[semai-sig] Strategy 4 (____ separator): collapsing from [${(underscoreEl).tagName || ""}${(underscoreEl).className ? "." + (underscoreEl).className.toString().split(" ")[0] : ""}]`);
    semaiCollapseFrom(underscoreEl);
    return;
  }

  // ── Strategy 5: Specific nested div contact-card signature ────────────
  const nestedDivSig = semaiFindNestedDivSignature(body);
  if (nestedDivSig) {
    semaiNativeLog(`[semai-sig] Strategy 5 (nested div[dir=ltr]): hiding [${(nestedDivSig).tagName || ""}${(nestedDivSig).className ? "." + (nestedDivSig).className.toString().split(" ")[0] : ""}]`);
    semaiHideEl(nestedDivSig);
    return;
  }
  semaiNativeLog(`[semai-sig] Strategy 5 skipped (no nested-div signature found)`);

  // ── Strategy 6: Heuristic — contact-card block near the bottom ───────
  // Last resort: a block with 5+ short lines containing phone/URL/social.
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
  if (!strategy6Hit) semaiNativeLog(`[semai-sig] Strategy 6 skipped (no heuristic match) — signature NOT detected`);
}

// ===== CHAT VIEW ============================================================

let semaiChatViewActive = false;
let semaiCurrentUser = null; // { name, email, initials }
let semaiReportHoverRow = null;
let semaiReportModeOverlay = null;

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

function semaiEscapeHtml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function semaiBuildGitHubIssueBody(message, subject) {
  const senderName = message.sender?.name || "Unknown";
  const senderEmail = message.sender?.email || "Unknown";
  const timestamp = message.timestamp || "Unknown";
  const excerpt = (message.cleanHtml || "")
    .replace(/<[^>]+>/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 500);

  return [
    "## Reported from REMOU",
    "",
    `- Subject: ${subject || "Conversation"}`,
    `- Sender: ${senderName}`,
    `- Sender Email: ${senderEmail}`,
    `- Timestamp: ${timestamp}`,
    `- Page URL: ${window.location.href}`,
    "",
    "## Excerpt",
    "",
    excerpt || "No text extracted.",
    "",
    "## Clean HTML",
    "",
    "```html",
    message.cleanHtml || "",
    "```",
    "",
    "## Original HTML",
    "",
    "```html",
    message.rawHtml || "",
    "```"
  ].join("\n");
}

async function semaiCreateGitHubIssue(message, subject) {
  if (!REMOU_GITHUB_TOKEN) {
    throw new Error("Missing GitHub token in secrets.js.");
  }

  if (!REMOU_GITHUB_REPO) {
    throw new Error("Missing GitHub repo in secrets.js.");
  }

  const titleParts = [
    "UI issue",
    subject || "Conversation",
    message.sender?.name || "Unknown sender"
  ];
  const response = await fetch(`https://api.github.com/repos/${REMOU_GITHUB_REPO}/issues`, {
    method: "POST",
    headers: {
      "Accept": "application/vnd.github+json",
      "Authorization": `Bearer ${REMOU_GITHUB_TOKEN}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify({
      title: titleParts.join(" | ").slice(0, 240),
      body: semaiBuildGitHubIssueBody(message, subject)
    })
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    throw new Error(errorData.message || `GitHub API error ${response.status}`);
  }

  return response.json();
}

function semaiClearReportHover() {
  if (semaiReportHoverRow) {
    semaiReportHoverRow.classList.remove("semai-chat-row-report-hover");
    semaiReportHoverRow = null;
  }
}

function semaiSetReportModeStatus(overlay, message, tone = "neutral") {
  const status = overlay?.querySelector("#semai-chat-reply-status");
  if (!status) return;
  status.textContent = message;
  status.dataset.tone = tone;
}

function semaiSetReportModeHint(overlay, message = "") {
  const hint = overlay?.querySelector("#semai-chat-report-hint");
  if (!hint) return;

  hint.textContent = message;
  hint.hidden = !message;
}

function semaiHandleReportModeKeydown(event) {
  if (event.key !== "Escape" || !semaiReportModeOverlay) return;

  event.preventDefault();
  semaiExitReportMode(semaiReportModeOverlay);
}

function semaiExitReportMode(overlay, statusMessage, tone = "neutral") {
  if (!overlay) return;

  overlay.classList.remove("semai-chat-report-mode");
  overlay.dataset.reportMode = "inactive";
  document.removeEventListener("keydown", semaiHandleReportModeKeydown, true);
  semaiReportModeOverlay = null;
  semaiClearReportHover();
  semaiSetReportModeHint(overlay, "");

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

  semaiSetReportModeHint(
    overlay,
    "Hover an email, click to report it, or press Esc to cancel."
  );

  semaiSetReportModeStatus(
    overlay,
    "Report mode is on.",
    "report"
  );
}

function semaiToggleReportMode(overlay) {
  if (!overlay) return;

  if (overlay.dataset.reportMode === "active") {
    semaiExitReportMode(
      overlay,
      overlay.dataset.viewMode === "real"
        ? "The original Outlook thread is visible above the reply box. Use the eye button to switch back to chat bubbles."
        : "Chat view is on. Use the eye button to switch only the thread view above this reply box."
    );
    return;
  }

  semaiEnterReportMode(overlay);
}

function semaiGetReportRowFromEventTarget(target) {
  if (!(target instanceof Element)) return null;
  return target.closest(".semai-chat-row[data-report-index]");
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

async function semaiHandleReportRowClick(event) {
  const overlay = semaiReportModeOverlay;
  if (!overlay || overlay.dataset.reportMode !== "active") return;

  const row = semaiGetReportRowFromEventTarget(event.target);
  if (!row) return;

  event.preventDefault();
  event.stopPropagation();

  const reportButton = overlay.querySelector("#semai-chat-report-issue-btn");
  const index = Number(row.dataset.reportIndex);
  const message = overlay._semaiMessages?.[index];
  const subject = overlay._semaiSubject || "Conversation";
  if (!message) return;

  if (reportButton) {
    reportButton.disabled = true;
  }

  semaiSetReportModeStatus(overlay, "Creating GitHub issue…", "report");

  try {
    const issue = await semaiCreateGitHubIssue(message, subject);
    semaiExitReportMode(
      overlay,
      `Issue #${issue.number} created for ${message.sender?.name || "this message"}.`,
      "success"
    );
  } catch (error) {
    semaiSetReportModeStatus(overlay, error.message || "Failed to create GitHub issue.", "error");
    if (reportButton) {
      reportButton.disabled = false;
    }
  }
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

function semaiFindCalibrationTarget(startEl) {
  if (!(startEl instanceof Element)) return null;

  const candidate = startEl.closest(
    '.OZZZK, [data-testid="senderName"], [class*="senderName" i], [class*="sender-name" i], .ms-Persona-primaryText, span, button, div'
  );
  if (!candidate) return null;

  const text = (candidate.innerText || candidate.textContent || "").trim();
  if (!text || text.length > 160) return null;
  if (!/[A-Za-z]/.test(text)) return null;

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
  semaiCalibrationState = null;
  semaiCurrentUser = null;
  semaiGetCurrentUser();
  semaiClearCalibrationHover();
  document.body.classList.remove("semai-calibrating");
  semaiUpdateCalibrationStatus(`Saved. Using "${selfSender.name}" as you.`, "success");
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
      `Captured you as "${semaiNormalizeSenderLabel(text).name}". Now click a different sender.`,
      "other"
    );
    return;
  }

  semaiFinishCalibration(
    semaiCalibrationState.selfLabel,
    text,
    semaiCalibrationState.selector || semaiBuildSenderSelector(target)
  );
  document.removeEventListener("mousemove", semaiHandleCalibrationHover, true);
  document.removeEventListener("click", semaiHandleCalibrationClick, true);
}

function semaiStartCalibration() {
  document.removeEventListener("click", semaiHandleCalibrationClick, true);
  document.removeEventListener("mousemove", semaiHandleCalibrationHover, true);
  semaiClearCalibrationHover();
  semaiCalibrationState = {
    step: "self",
    selfLabel: "",
    selector: null
  };

  document.body.classList.add("semai-calibrating");
  semaiUpdateCalibrationStatus("Setup step 1 of 2: hover and click your sender label.", "self");
  document.addEventListener("mousemove", semaiHandleCalibrationHover, true);
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

  // ── 0a. Short-circuit: if the entire body is a compact branded signature, return empty ──
  semaiNativeLog(`[semai-sig] cleanBodyClone: senderFirstName="${senderFirstName}" bodyLen=${clone.textContent.length}`);
  if (semaiIsEntireBodySignature(clone, senderFirstName)) {
    semaiNativeLog(`[semai-sig] cleanBodyClone: Step 0a → entire body is a signature, returning empty`);
    return "";
  }

  // ── 0. Remove Outlook "external sender" warning banners ──
  //    These are injected by Exchange/Outlook and contain links to aka.ms/LearnAboutSenderIdentification
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

  // ── 4. Remove quoted-reply header blocks like From / Date / Sent / To / Subject ──
  semaiStripQuotedReplyHeaders(clone);

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
  if (anchor) {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 6 (sender anchor): removing from [${(anchor).tagName || ""}${(anchor).className ? "." + (anchor).className.toString().split(" ")[0] : ""}]`);
    removeFromAndAfter(anchor);
  } else {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 6: no sender anchor for "${senderFirstName}"`);
  }

  // ── 6b. Strip repeated-name contact cards with 2+ contact signals ──────
  const repeatedNameAnchor = semaiFindRepeatedNameSignatureAnchor(clone, senderFirstName);
  if (repeatedNameAnchor) {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 6b (repeated-name anchor): removing from [${(repeatedNameAnchor).tagName || ""}${(repeatedNameAnchor).className ? "." + (repeatedNameAnchor).className.toString().split(" ")[0] : ""}]`);
    removeFromAndAfter(repeatedNameAnchor);
  } else {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 6b: no repeated-name anchor found`);
  }

  // ── 7. Strip specific nested-div contact-card signatures directly ──────
  const nestedDivSig = semaiFindNestedDivSignature(clone);
  if (nestedDivSig) {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 7 (nested div[dir=ltr]): removing [${(nestedDivSig).tagName || ""}${(nestedDivSig).className ? "." + (nestedDivSig).className.toString().split(" ")[0] : ""}]`);
    nestedDivSig.remove();
  } else {
    semaiNativeLog(`[semai-sig] cleanBodyClone step 7: no nested-div signature found`);
  }

  // ── 8. Strip closing phrase + name even when no contact block follows ──
  // e.g. "Best,\nMichael" at the very end — remove the closing and name too.
  semaiStripTrailingSignOff(clone, senderFirstName);

  // ── 9. Remove trailing empty elements ──
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

function semaiStripQuotedReplyHeaders(container) {
  const HEADER_LINE_RE = /^(from|date|sent|to|cc|subject)\s*:/i;
  const HEADER_PAIR_RE = /(from|date|sent|to|cc|subject)\s*:.*\n.*(from|date|sent|to|cc|subject)\s*:/is;
  const WROTE_LINE_RE = /^on .+wrote:\s*$/i;

  function removeFromAndAfter(el) {
    let sib = el;
    while (sib) {
      const next = sib.nextElementSibling;
      sib.remove();
      sib = next;
    }
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
    const senderFirstName = semaiFirstNameFromDisplayName(sender.name);
    semaiNativeLog(`[semai-sig] extractThreadMessages: sender.name="${sender.name}" → senderFirstName="${senderFirstName}"`);
    const cleanHtml = semaiCleanBodyClone(bodyEl, senderFirstName);
    const rawHtml = bodyEl.dataset.semaiOriginalHtml || bodyEl.innerHTML;
    const isMe = semaiIsCurrentUser(sender.name, sender.email);
    return { sender, timestamp, cleanHtml, rawHtml, isMe };
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
  overlay.dataset.viewMode = "chat";
  overlay.dataset.reportMode = "inactive";
  overlay._semaiMessages = messages;
  overlay._semaiSubject = subject;

  // Header bar
  const header = document.createElement("div");
  header.className = "semai-chat-header";
  header.innerHTML = `
    <span class="semai-chat-subject">${semaiEscapeHtml(subject)}</span>
    <button class="semai-chat-close" type="button">✕ Hide chat view</button>
  `;
  header.querySelector(".semai-chat-close").addEventListener("click", semaiDeactivateChatView);
  overlay.appendChild(header);

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

    if (msg.isMe) {
      row.appendChild(bubble);
      row.appendChild(avatar);
    } else {
      row.appendChild(avatar);
      row.appendChild(bubble);
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
      placeholder="Type a reply-all draft for the latest message…"
    ></textarea>
    <div class="semai-chat-composer-footer">
      <div id="semai-chat-reply-status" class="semai-chat-reply-status">
        Chat view is on. Use the eye button to switch only the thread view above this reply box.
      </div>
      <div class="semai-chat-reply-actions">
        <button
          id="semai-chat-report-issue-btn"
          class="semai-chat-report-issue-btn"
          type="button"
        >
          Report issue
        </button>
        <span id="semai-chat-report-hint" class="semai-chat-report-hint" hidden></span>
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
        <button id="semai-chat-reply-draft-btn" class="semai-chat-reply-draft-btn" type="button">
          Draft
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
  const draftBtn = composer.querySelector("#semai-chat-reply-draft-btn");
  const replyBtn = composer.querySelector("#semai-chat-reply-send-btn");

  reportIssueBtn.addEventListener("click", () => {
    semaiToggleReportMode(overlay);
  });
  viewToggleBtn.addEventListener("click", () => {
    semaiToggleOverlayView(overlay);
  });
  draftBtn.addEventListener("click", semaiDraftReplyAllFromChat);
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
      ? "Chat view is on. Use the eye button to switch only the thread view above this reply box."
      : "The original Outlook thread is visible above the reply box. Use the eye button to switch back to chat bubbles.";
    delete status.dataset.tone;
  }

  if (overlay._semaiReadingPane) {
    semaiUpdateReadingPaneBottomClearance(overlay._semaiReadingPane, overlay);
  }
}

function semaiToggleOverlayView(overlay) {
  if (!overlay) return;

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
  const readingPane = overlay?.parentElement;
  document.removeEventListener("keydown", semaiHandleReportModeKeydown, true);
  semaiReportModeOverlay = null;
  semaiClearReportHover();
  overlay?._semaiResizeObserver?.disconnect();
  if (overlay) overlay.remove();
  semaiRemoveReadingPaneBottomClearance(readingPane);
  semaiChatViewActive = false;
  const bodies = document.querySelectorAll('[aria-label="Message body"]:not([contenteditable])');
  semaiAutoOpenSuppressedSignature = Array.from(bodies).map(b => b.dataset.semaiSigStripped || "").join("|");
  semaiUpdateChatToggleBtn();
  semaiLog("[semai] Chat view deactivated");
}

function semaiUpdateChatToggleBtn() {
  const btn = document.querySelector(".semai-chat-toggle-btn");
  if (!btn) return;
  btn.textContent = semaiChatViewActive ? "Hide chat view" : "Turn on chat view";
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
    if (sig !== semaiLastReadingPaneSignature) {
      semaiAutoOpenSuppressedSignature = "";
    }
    semaiLastReadingPaneSignature = sig;
    semaiUpdateChatToggleVisibility();
    if (
      !semaiChatViewActive &&
      !semaiCalibrationState &&
      bodies.length >= 2 &&
      sig &&
      sig !== semaiAutoOpenSuppressedSignature
    ) {
      semaiActivateChatView();
    }
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
  window.addEventListener("resize", () => {
    const panel = document.getElementById("semai-panel");
    if (panel) semaiEnsurePanelVisible(panel, false);
  });

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
