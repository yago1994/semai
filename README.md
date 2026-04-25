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

## Resolved defenses

### Safe-match gate (was gotcha #8 — wrong-thread reply risk)

**Risk.** Two unrelated threads can share `BodyPreview` text ("thanks!", "Sent from my iPhone", boilerplate replies). Pivoting on body-preview alone could misroute a reply-all into the wrong thread with the wrong recipients. This is a privacy/correctness disaster the system MUST NOT allow.

**Implementation** — `semaiResolveMessageIdViaRest` in `contentScript.js` now applies six independent gates before it will return a `messageId`:

| # | Gate | Threshold |
|---|------|-----------|
| 1 | Snippet length | `≥ SEMAI_MIN_SNIPPET_LEN` (30 chars after the 20-char greeting skip) |
| 2 | Snippet must not be dominated by a generic phrase | "thanks", "ok", "got it", "sounds good", "regards", "sent from my iphone", etc. — match if any phrase covers >70% of the snippet |
| 3 | Sender email cross-check | candidate's `From.EmailAddress.Address` (or `Sender.EmailAddress.Address` for delegated send) must equal the overlay message's parsed sender email when both are known |
| 4 | Date proximity | candidate's `ReceivedDateTime` within `SEMAI_MAX_DATE_DELTA_MIN` (5 min) of the overlay's parsed timestamp; skipped only if overlay timestamp is unparseable |
| 5 | Multi-message confirmation | `≥ SEMAI_MIN_CONFIRMING_MATCHES` (2) **distinct** overlay messages must survive gates 1-4 against the SAME `ConversationId` |
| 6 | Single-match override | a lone match is accepted ONLY if its snippet is `≥ SEMAI_STRICT_SNIPPET_LEN` (40 chars), sender email is verified-equal, AND date is within `SEMAI_STRICT_DATE_DELTA_MIN` (2 min) |

If any of these fail, `semaiResolveMessageIdViaRest` throws and `semaiTryReplyAllViaRestApi` falls back to the compose-UI path. The fallback leaves a draft (UX wart) but **cannot** misroute the reply, which is the invariant we protect.

The full evaluation trail (per-overlay snippet, sender email, body hits, sender rejects, date rejects, accepted candidates, dominant convId, distinct confirmations) is dumped into the chat-overlay debug panel under `SAFE:` lines so any field failure can be diagnosed without instrumenting builds.

**Tunables** are constants at the top of the SAFE-MATCH GATE block in `contentScript.js`. `SEMAI_GENERIC_PHRASES` is also there for additions.

**Open follow-up.** A unit-test fixture for the "two threads with identical 'thanks!' reply" case still needs to be added under `tests/` — the gate logic is testable in isolation against a JSON `recent` fixture and a synthetic overlay-messages array.

---

### Token lifecycle (was gotcha #4 — token expiration after long idle)

**Risk.** Outlook bearer tokens are JWTs with a ~60min lifetime. The page-world hook re-captures fresh tokens whenever Outlook polls its own backends, but on wake-from-sleep / long-idle, the cached token can be stale. A stale-token POST returns 401, the REST path aborts, and the user sees a phantom draft as the compose-UI fallback runs.

**Implementation** — `contentScript.js` now layers two defenses inside `semaiCallOutlookApi`:

| Layer | What it does | Where |
|-------|--------------|-------|
| Proactive expiry check | Decodes the JWT `exp` claim of the cached token; if it expires within `SEMAI_TOKEN_EXPIRY_BUFFER_MS` (60s), treats the cache as empty and refuses the call until a fresh token arrives | `semaiGetUsableToken` / `semaiDecodeJwtExp` |
| Reactive 401 retry | On 401, wipes the cache, polls for up to `SEMAI_TOKEN_REFRESH_TIMEOUT_MS` (5s) waiting for the page-world hook to publish a *different* token, then retries the original request once | `semaiCallOutlookApi` / `semaiWaitForFreshToken` |

The reactive path works because Outlook's SPA polls its own backends every few seconds (mail/presence/focused-inbox). Wiping our cache doesn't affect Outlook's behavior — within 1-3 seconds it issues another Bearer-tagged fetch and the hook publishes a fresh token via the `semai-outlook-token` CustomEvent. The retry uses that new token.

If the retry also returns 401 (extension permanently revoked, or no fresh token within 5s), the function returns the 401 response and the caller falls back to compose-UI. The user still sends successfully; they just get the draft side effect.

**Tunables.** `SEMAI_TOKEN_EXPIRY_BUFFER_MS`, `SEMAI_TOKEN_REFRESH_TIMEOUT_MS`, `SEMAI_TOKEN_REFRESH_POLL_MS` at the top of the TOKEN LIFECYCLE block in `contentScript.js`.

**Open follow-up.** Persisting the captured token in `chrome.storage.session` would also smooth over service-worker restarts (currently we only cache in content-script module memory, which is per-page-load). Worth doing alongside #5 (audience scoping) so we don't accidentally persist a wrong-audience token.

---

### Install-race recovery (was gotcha #9 — SPA reload race on first install)

**Risk.** When the extension is installed into a browser that already has Outlook tabs open, `pageWorldHook.js` only injects after the user's next chrome.runtime activation. By then Outlook's SPA bundles have already executed and minted their initial tokens. Until Outlook's next background poll, no token is cached and the first reply falls back to compose-UI silently — leaving a draft with no explanation.

**Implementation** — two layers:

1. **Auto-reload on first install** (`background.js`, `reloadOpenOutlookTabs`). The `chrome.runtime.onInstalled` listener now checks `details.reason === 'install'` and force-reloads any open Outlook tabs. We deliberately do NOT reload on `'update'` / `'browser_update'` / `'shared_module_update'` — auto-updates shouldn't kick a user out of an in-progress compose. Uses `chrome.tabs.query` filtered by the same host patterns the manifest declares (no `tabs` permission needed since `host_permissions` covers these origins).

2. **One-shot inline hint banner** (`contentScript.js`, `semaiShowMissingTokenHint`). If a reply attempt finds `semaiCachedOutlookToken === ""` and falls back to compose-UI, we surface a yellow `.semai-chat-token-hint` banner above the reply input: *"Tip: reload this Outlook tab to send replies without leaving a draft."* The banner has a dismiss button and auto-clears after 30s. The `semaiMissingTokenHintShown` module flag ensures it shows at most once per page session — we don't nag.

These cover different failure modes:
- The auto-reload fixes the genuine first-install race (extension installed mid-Outlook-session).
- The hint covers cases where auto-reload didn't run (manual install via Safari devmode, content-script restart without page reload, Safari-specific edge cases) — and tells the user exactly what to do.

**Open follow-up.** A "provoke" path that fires a benign authenticated fetch from the page world to force Outlook to attach a Bearer header (without reloading) would eliminate the install-time disruption entirely. Needs an Outlook endpoint that's safe to hit without side effects and that we've verified attaches Authorization. Lower priority now that auto-reload covers the common path.

---

## Known limitations & plan of action

Two known issues will hit real users eventually. Neither is addressed yet — this section is the punch list.

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
