// Background service worker — relays log messages from content scripts to the
// native host so they appear in the Xcode console via os_log.
browser.runtime.onMessage.addListener((message) => {
  if (message && message.type === "semaiLog") {
    browser.runtime.sendNativeMessage("yam.team.remou", { log: message.text });
  }
});
