# remou (semai)

Safari/Chrome MV3 extension that transforms Outlook Web email threads into a chat-like interface.

The extension lives under [`semai Extension/Resources/`](./semai%20Extension/Resources/). The Xcode project (`semai.xcodeproj`) wraps it as a Safari Web Extension; the same `Resources/` folder also loads as an unpacked Chrome MV3 extension.

---

## Reply All via REST API (no draft side effect)

### Why this exists

When the user hits "Send" in our custom chat overlay, we previously opened Outlook's native Reply All compose UI, filled it, and clicked Send. Outlook autosaves drafts the moment compose opens — so every reply we sent left a phantom empty draft behind, and any cleanup attempt triggered a "confirm deletion" dialog. That UX was unacceptable.

The fix: skip the compose UI entirely. Capture an OAuth bearer token off Outlook's own network traffic, then `POST` to `outlook.office.com/api/v2.0/me/messages/{id}/replyAll` ourselves. No compose UI is ever opened, so no autosave draft is ever created.

### Architecture (4 files cooperating)

```
┌────────────────────┐  injects   ┌──────────────────────┐  CustomEvent  ┌──────────────────────┐
│ content.js         │──────────▶ │ pageWorldHook.js     │──────────────▶│ contentScript.js     │
│ (isolated world,   │ at         │ (page main world,    │ "semai-       │ (isolated world,     │
│  document_start)   │ document-  │  hooks fetch + XHR)  │  outlook-     │  document_idle)      │
│                    │ _start     │                      │  token")      │                      │
└────────────────────┘            └──────────────────────┘               └──────────┬───────────┘
                                                                                    │
                                                                       chrome.runtime.sendMessage
                                                                       { type:'OUTLOOK_API_CALL' }
                                                                                    │
                                                                                    ▼
                                                                         ┌──────────────────────┐
                                                                         │ background.js        │
                                                                         │ (service worker;     │
                                                                         │  cross-origin fetch  │
                                                                         │  to outlook.office   │
                                                                         │  .com/api/v2.0)      │
                                                                         └──────────────────────┘
```

1. **`pageWorldHook.js`** — Runs in the page's main world (NOT the isolated content-script world). Wraps `window.fetch` and `XMLHttpRequest.prototype.open / setRequestHeader` to read the `Authorization: Bearer …` header off every request Outlook makes to its own backends. Publishes captured tokens via `document.dispatchEvent(new CustomEvent('semai-outlook-token', { detail: { token, at } }))`. Idempotent via `window.__semaiTokenHookInstalled`.

2. **`content.js`** — Runs at `document_start` in the isolated world. Reads `chrome.runtime.getURL('pageWorldHook.js')` and injects it as a `<script src=…>` into `<head>` before Outlook's SPA bundles execute. (Has to be a real `<script src=>` because content scripts can't reach into the page's main world directly.) `pageWorldHook.js` is exposed via `web_accessible_resources` in `manifest.json`.

3. **`contentScript.js`** — Runs at `document_idle` in the isolated world. Listens for `semai-outlook-token` events and caches the latest token in `semaiCachedOutlookToken`. When the user sends a reply through the chat overlay, calls `semaiTryReplyAllViaRestApi(draft)`:
   - **Resolve the message ID** by listing the 50 most recent messages (`GET /me/messages?$top=50&$orderby=ReceivedDateTime desc`), matching one against an existing chat-overlay message via a 50-char `BodyPreview` snippet, then pivoting via the matched message's `ConversationId` to find the latest message in the thread.
   - **Send** via `POST /me/messages/{id}/replyAll` with `{ Comment: <plaintext> }`.
   - On any failure, falls back to the legacy compose-UI path (still has the draft side effect, but at least the reply lands).

4. **`background.js`** — Service worker. Handles `OUTLOOK_API_CALL` messages from `contentScript.js` and performs the actual cross-origin `fetch` to `outlook.office.com/api/v2.0`. This indirection exists because Safari's content-script context blocks `Authorization`-bearing fetches to `outlook.office.com` (CORS policy for content scripts is stricter than for service workers with `host_permissions`).

### Why we pivot via `ConversationId` instead of `$search`-ing for the subject

- `ConversationTopic` doesn't exist on every tenant — querying it returns `Could not find a property named 'ConversationTopic' on type 'Microsoft.OutlookServices.Message'`.
- `$search` against just-sent messages hits Outlook's search-index lag (5–60s) and returns 0 results.
- Modern Outlook DOM doesn't stamp message IDs anywhere reliably.

So instead: list the 50 most recent messages (always indexed, no lag), match one of them against an *older* message visible in our chat overlay (guaranteed to be indexed because it's old), grab its `ConversationId`, then take the most-recent message in the same conversation. The chain is robust because we never depend on the just-sent message being searchable.

### Debug panel

`.semai-chat-reply-debug` in the chat overlay is a scrollable, copy-pastable text region that shows a running log of the REST attempt: token capture, smoke test, list query, body-snippet matches, ConversationId, POST response. Backed by a module-level `semaiDebugLogBuffer` (500 lines max, append-only) plus a `MutationObserver` that replays the buffered history into any newly-rendered debug panel — so the log survives Outlook's overlay rebuilds.

### Feature flag

`SEMAI_USE_REST_API_REPLY = true` at the top of `contentScript.js`. Flip to `false` to disable the REST path and force the legacy compose-UI flow.

---

## Known limitations & plan of action

The REST path works, but five known issues will hit real users eventually. None of them are addressed yet — this section is the punch list.

### #4 — Token expiration after long idle

**Problem.** Outlook bearer tokens have ~60-minute lifetime. We cache the most-recently-captured token in `semaiCachedOutlookToken` and reuse it. If the user wakes their laptop after >1h without Outlook making any background traffic, the cached token is stale; the `replyAll` POST returns `401 Unauthorized`; we fall back to the compose-UI path and leave a draft.

**Plan.**
1. On `401` from `OUTLOOK_API_CALL`, don't fall back immediately. Instead, mark the cached token as invalid (`semaiCachedOutlookToken = null`) and trigger a token refresh by issuing a benign request to a known Outlook endpoint (e.g. fetch the user's own profile via the OWA frontdoor) — this provokes Outlook's SPA to mint a new token, which our hook captures.
2. Retry the original POST once with the new token.
3. Only fall back to compose UI if the retry also fails.
4. Optional: stash the token's `exp` claim (decode the JWT payload — base64 of segment 1) and proactively treat the cache as empty when within 60s of expiry.

**Files.** `contentScript.js` (`semaiCallOutlookApi`, `semaiTryReplyAllViaRestApi`).

---

### #5 — Token audience mismatch

**Problem.** Outlook's SPA mints multiple tokens with different audiences. Observed lengths in the wild: ~5241/5242 chars (outlook.office.com), ~10151 chars (likely Microsoft Graph). Our hook captures *all* `Bearer …` headers regardless of destination URL, so whichever token was minted most recently wins — even if it's the wrong audience for `outlook.office.com/api/v2.0`. Result: sporadic `401 Invalid audience` errors that don't repro on retry because by then the right token has been re-captured.

**Plan.**
1. In `pageWorldHook.js`, scope token capture by destination URL. Only publish a token if the request URL matches `^https://outlook\.office\.com/(api|owa)/`. Drop tokens captured from `graph.microsoft.com`, `substrate.office.com`, etc.
2. Augment the CustomEvent detail with `{ token, audience: 'outlook.office.com', at }` so `contentScript.js` can keep separate caches per audience if we ever need Graph too.
3. Add a sanity check in `contentScript.js`: decode the JWT, verify `payload.aud === 'https://outlook.office.com'` before caching.

**Files.** `pageWorldHook.js`, `contentScript.js`.

---

### #7 — Consumer Outlook.com / outlook.live.com support

**Problem.** All REST URLs are hardcoded to `https://outlook.office.com/api/v2.0/…`. That endpoint serves Microsoft 365 / commercial tenants. Consumer accounts (`@outlook.com`, `@hotmail.com`, `@live.com`) live on a separate frontdoor — `outlook.live.com` for OWA, but the API uses Microsoft Graph (`graph.microsoft.com/v1.0/me/messages`) with a different token audience. Currently, a consumer user's REST POST will likely 404 or 401, then fall back to compose UI (draft side effect).

**Plan.**
1. Detect tenant flavor at runtime: inspect `location.hostname` of the active tab (already known to background via the sender). `outlook.office.com` / `outlook.office365.com` / `outlook.cloud.microsoft` → commercial; `outlook.live.com` → consumer.
2. For consumer, switch the entire REST path to Graph: `https://graph.microsoft.com/v1.0/me/messages?$top=50&$orderby=receivedDateTime desc` and `POST /me/messages/{id}/replyAll` with `{ comment }` (note camelCase, lowercase property names — Graph's convention differs from Outlook REST v2.0).
3. Token capture must also scope by audience for Graph (see #5).
4. Add `https://graph.microsoft.com/*` to `host_permissions` in `manifest.json`.
5. Test on a personal `@outlook.com` account before shipping.

**Files.** `manifest.json`, `contentScript.js`, `pageWorldHook.js`.

---

### #8 — False `BodyPreview` matches (MAJOR — wrong-thread reply risk)

**Problem.** We resolve the target message ID by matching a 50-char `BodyPreview` snippet against the chat overlay. If two unrelated threads share similar body text (the canonical example: a one-word "thanks!" reply, or boilerplate "Sent from my iPhone" footers), we may pivot to the wrong `ConversationId` and reply-all into a completely unrelated thread. This is a privacy and correctness disaster — a reply meant for thread A could land in thread B with thread B's recipient list.

**Current mitigations.** We skip the first 20 chars (avoids "Hi Yago," / "Thanks," / signature openers) and take a 50-char window from the middle. Helps with greetings, doesn't help with one-line replies.

**Plan.**
1. **Sender email cross-check.** Every chat-overlay message carries the sender's email address in the DOM. Before accepting a `BodyPreview` match, verify the candidate message's `From.EmailAddress.Address` equals the overlay message's sender. If not, skip and try the next candidate.
2. **Date proximity check.** Each overlay message has a timestamp. Only accept matches within ±5 minutes of the overlay message's parsed date.
3. **Multi-message confirmation.** Require at least 2 overlay messages to match within the same `ConversationId` before pivoting. A coincidental match on one message is plausible; a coincidental match on two is vanishingly unlikely.
4. **Refuse to send if confidence is low.** If we can match only one overlay message with a generic snippet (e.g. snippet contains common phrases "thanks", "ok", "got it", or is shorter than 30 chars after the skip), abort the REST path and fall back to compose UI. Better to leave a draft than misroute a reply.
5. Add a unit test fixture covering the "two threads with identical 'thanks!' reply" case.

**Files.** `contentScript.js` (`semaiResolveMessageIdViaRest`, `semaiBodySnippet`).

---

### #9 — SPA reload race on first install

**Problem.** `pageWorldHook.js` only intercepts tokens for fetches that happen *after* it executes. On first install (or after extension update), the extension activates mid-session — Outlook's SPA bundles already loaded, the initial token-minting fetch already happened. Result: no token captured until the user manually reloads the Outlook tab. Until they do, every reply falls back to the compose-UI path with the draft side effect, and the user has no idea why.

**Plan.**
1. **Detect missing-token state.** When `semaiTryReplyAllViaRestApi` finds `semaiCachedOutlookToken === null`, surface a one-time toast in the chat overlay: "Reload the Outlook tab to enable seamless replies."
2. **Auto-reload on first install.** In `background.js`, listen for `chrome.runtime.onInstalled` with `reason === 'install' || reason === 'update'`. Use `chrome.tabs.query({ url: ['https://outlook.office.com/*', …] })` and `chrome.tabs.reload(tabId)` for any matching open tab. Caveats: this loses unsaved compose state in Outlook's UI — gate it behind a confirm prompt or only do it on `install`, not `update`.
3. **Provoke a token-minting request.** Alternative to reload: from `contentScript.js`, after install, fire a benign no-op fetch via the page world (e.g. `await fetch('/owa/service.svc?action=GetUserConfiguration', { credentials: 'include' })`). Outlook's SPA will attach a `Bearer` header → our hook captures it. Doesn't reload the tab. Need to verify Outlook actually attaches Authorization to that endpoint.
4. **Persist last-known-good token.** Store the captured token (encrypted, with `exp`) in `chrome.storage.session` so a service-worker restart doesn't lose it. Doesn't solve first-install but does smooth over background-script wake-ups.

**Files.** `background.js`, `contentScript.js`, `pageWorldHook.js`.

---

## Repo layout

| Path | Purpose |
|------|---------|
| `semai Extension/Resources/` | The actual extension. Loads as Safari + Chrome MV3 from the same files. |
| `semai Extension/Resources/manifest.json` | MV3 manifest. |
| `semai Extension/Resources/content.js` | document_start injector for `pageWorldHook.js`. |
| `semai Extension/Resources/pageWorldHook.js` | Page main-world fetch/XHR hook for token capture. |
| `semai Extension/Resources/contentScript.js` | Main content script — chat overlay + reply orchestration. |
| `semai Extension/Resources/background.js` | Service worker — patches feed + cross-origin REST proxy. |
| `semai Extension/Resources/styles.css` | Chat overlay styles. |
| `semai.xcodeproj/` | Xcode wrapper for the Safari Web Extension. |
| `docs/` | Signature-detection edge cases + patches feed source. |
| `scripts/`, `tests/` | Build helpers and unit tests. |
