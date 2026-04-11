function show(enabled, useSettingsInsteadOfPreferences) {
    if (useSettingsInsteadOfPreferences) {
        document.getElementsByClassName('state-on')[0].innerText = "remou’s extension is currently on. You can turn it off in the Extensions section of Safari Settings.";
        document.getElementsByClassName('state-off')[0].innerText = "remou’s extension is currently off. You can turn it on in the Extensions section of Safari Settings.";
        document.getElementsByClassName('state-unknown')[0].innerText = "You can turn on remou’s extension in the Extensions section of Safari Settings.";
        document.getElementsByClassName('open-preferences')[0].innerText = "Quit and Open Safari Settings…";
    }

    if (typeof enabled === "boolean") {
        document.body.classList.toggle(`state-on`, enabled);
        document.body.classList.toggle(`state-off`, !enabled);
    } else {
        document.body.classList.remove(`state-on`);
        document.body.classList.remove(`state-off`);
    }
}

function openPreferences() {
    webkit.messageHandlers.controller.postMessage("open-preferences");
}

const reportIssueInput = document.getElementById("report-issue-input");
const reportIssueButton = document.querySelector("button.report-issue-btn");
const reportIssueStatus = document.getElementById("report-issue-status");

function setReportIssueStatus(message, tone = "neutral") {
    reportIssueStatus.textContent = message || "";
    reportIssueStatus.classList.remove("is-success", "is-error");

    if (tone === "success") {
        reportIssueStatus.classList.add("is-success");
    } else if (tone === "error") {
        reportIssueStatus.classList.add("is-error");
    }
}

function setReportIssueBusy(isBusy) {
    reportIssueButton.disabled = isBusy;
    reportIssueInput.disabled = isBusy;
    reportIssueButton.textContent = isBusy ? "Reporting…" : "Report an issue";
}

function submitReportIssue() {
    const reason = reportIssueInput.value.trim();
    if (!reason) {
        setReportIssueStatus("Add a short description before sending the report.", "error");
        reportIssueInput.focus();
        return;
    }

    setReportIssueBusy(true);
    setReportIssueStatus("Creating GitHub issue…");
    webkit.messageHandlers.controller.postMessage({
        command: "report-issue",
        reason,
        extensionEnabled: document.body.classList.contains("state-on"),
        pageUrl: window.location.href,
        userAgent: window.navigator.userAgent
    });
}

function reportIssueResult(success, message) {
    setReportIssueBusy(false);
    setReportIssueStatus(
        message || (success ? "Issue reported successfully." : "Failed to report issue."),
        success ? "success" : "error"
    );

    if (success) {
        reportIssueInput.value = "";
    }
}

document.querySelector("button.open-preferences").addEventListener("click", openPreferences);
reportIssueButton.addEventListener("click", submitReportIssue);
