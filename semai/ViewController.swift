//
//  ViewController.swift
//  remou
//
//  Created by Yago Arconada on 11/13/25.
//

import Cocoa
import Foundation
import SafariServices
import WebKit

let extensionBundleIdentifier = "yam.team.remou.Extension"

class ViewController: NSViewController, WKNavigationDelegate, WKScriptMessageHandler {

    @IBOutlet var webView: WKWebView!

    struct GitHubSecrets {
        let token: String
        let repository: String
    }

    override func viewDidLoad() {
        super.viewDidLoad()

        self.webView.navigationDelegate = self

        self.webView.configuration.userContentController.add(self, name: "controller")

        self.webView.loadFileURL(Bundle.main.url(forResource: "Main", withExtension: "html")!, allowingReadAccessTo: Bundle.main.resourceURL!)
    }

    func webView(_ webView: WKWebView, didFinish navigation: WKNavigation!) {
        SFSafariExtensionManager.getStateOfSafariExtension(withIdentifier: extensionBundleIdentifier) { (state, error) in
            guard let state = state, error == nil else {
                // Insert code to inform the user that something went wrong.
                return
            }

            DispatchQueue.main.async {
                if #available(macOS 13, *) {
                    webView.evaluateJavaScript("show(\(state.isEnabled), true)")
                } else {
                    webView.evaluateJavaScript("show(\(state.isEnabled), false)")
                }
            }
        }
    }

    func userContentController(_ userContentController: WKUserContentController, didReceive message: WKScriptMessage) {
        if let command = message.body as? String {
            if command != "open-preferences" {
                return
            }

            SFSafariApplication.showPreferencesForExtension(withIdentifier: extensionBundleIdentifier) { error in
                if let error {
                    NSLog("Failed to open Safari extension preferences: \(error.localizedDescription)")
                }
            }
            return
        }

        guard
            let payload = message.body as? [String: Any],
            let command = payload["command"] as? String
        else {
            return
        }

        if command == "report-issue" {
            let reason = (payload["reason"] as? String)?.trimmingCharacters(in: .whitespacesAndNewlines) ?? ""
            let extensionEnabled = payload["extensionEnabled"] as? Bool ?? false
            let pageURL = payload["pageUrl"] as? String ?? ""
            let userAgent = payload["userAgent"] as? String ?? ""
            submitOnboardingIssue(reason: reason, extensionEnabled: extensionEnabled, pageURL: pageURL, userAgent: userAgent)
        }
    }

    private func submitOnboardingIssue(reason: String, extensionEnabled: Bool, pageURL: String, userAgent: String) {
        guard !reason.isEmpty else {
            sendReportIssueResult(success: false, message: "Add a short description before sending the report.")
            return
        }

        Task {
            do {
                let secrets = try loadGitHubSecrets()
                let body = buildOnboardingIssueBody(
                    reason: reason,
                    extensionEnabled: extensionEnabled,
                    pageURL: pageURL,
                    userAgent: userAgent
                )
                try await createGitHubIssue(
                    title: "Onboarding issue | Safari setup screen",
                    body: body,
                    secrets: secrets
                )
                sendReportIssueResult(
                    success: true,
                    message: "Issue reported successfully. Thank you."
                )
            } catch {
                sendReportIssueResult(
                    success: false,
                    message: error.localizedDescription
                )
            }
        }
    }

    private func sendReportIssueResult(success: Bool, message: String) {
        DispatchQueue.main.async {
            let successLiteral = success ? "true" : "false"
            let escapedMessage = Self.javaScriptStringLiteral(message)
            self.webView.evaluateJavaScript("reportIssueResult(\(successLiteral), \(escapedMessage));")
        }
    }

    private func loadGitHubSecrets() throws -> GitHubSecrets {
        guard
            let pluginsURL = Bundle.main.builtInPlugInsURL
        else {
            throw NSError(domain: "remou", code: 1, userInfo: [NSLocalizedDescriptionKey: "Could not locate the bundled Safari extension."])
        }

        let pluginURLs = try FileManager.default.contentsOfDirectory(
            at: pluginsURL,
            includingPropertiesForKeys: nil
        )

        guard
            let extensionBundleURL = pluginURLs.first(where: {
                Bundle(url: $0)?.bundleIdentifier == extensionBundleIdentifier
            }),
            let extensionBundle = Bundle(url: extensionBundleURL),
            let secretsURL = extensionBundle.url(forResource: "secrets", withExtension: "js")
        else {
            throw NSError(domain: "remou", code: 2, userInfo: [NSLocalizedDescriptionKey: "Could not load remou issue-report configuration."])
        }

        let source = try String(contentsOf: secretsURL, encoding: .utf8)
        let token = try extractJavaScriptConstant(named: "REMOU_GITHUB_TOKEN", from: source)
        let repository = try extractJavaScriptConstant(named: "REMOU_GITHUB_REPO", from: source)
        return GitHubSecrets(token: token, repository: repository)
    }

    private func extractJavaScriptConstant(named name: String, from source: String) throws -> String {
        let pattern = "const\\s+\(NSRegularExpression.escapedPattern(for: name))\\s*=\\s*\"([^\"]+)\""
        let regex = try NSRegularExpression(pattern: pattern)
        let range = NSRange(source.startIndex..<source.endIndex, in: source)

        guard
            let match = regex.firstMatch(in: source, options: [], range: range),
            let valueRange = Range(match.range(at: 1), in: source)
        else {
            throw NSError(domain: "remou", code: 3, userInfo: [NSLocalizedDescriptionKey: "Missing \(name) in extension secrets.js."])
        }

        return String(source[valueRange])
    }

    private func buildOnboardingIssueBody(reason: String, extensionEnabled: Bool, pageURL: String, userAgent: String) -> String {
        [
            "## Reported from REMOU onboarding",
            "",
            "- Extension Enabled: \(extensionEnabled ? "Yes" : "No")",
            "- Page URL: \(pageURL)",
            "- macOS: \(ProcessInfo.processInfo.operatingSystemVersionString)",
            "- User Agent: \(userAgent)",
            "",
            "## Reason It's An Issue",
            "",
            reason
        ].joined(separator: "\n")
    }

    private func createGitHubIssue(title: String, body: String, secrets: GitHubSecrets) async throws {
        guard let url = URL(string: "https://api.github.com/repos/\(secrets.repository)/issues") else {
            throw NSError(domain: "remou", code: 4, userInfo: [NSLocalizedDescriptionKey: "Could not construct the GitHub issues URL."])
        }

        var request = URLRequest(url: url)
        request.httpMethod = "POST"
        request.setValue("application/vnd.github+json", forHTTPHeaderField: "Accept")
        request.setValue("Bearer \(secrets.token)", forHTTPHeaderField: "Authorization")
        request.setValue("application/json", forHTTPHeaderField: "Content-Type")
        request.httpBody = try JSONSerialization.data(withJSONObject: [
            "title": title,
            "body": body
        ])

        let (data, response) = try await URLSession.shared.data(for: request)
        guard let httpResponse = response as? HTTPURLResponse else {
            throw NSError(domain: "remou", code: 5, userInfo: [NSLocalizedDescriptionKey: "GitHub did not return a valid response."])
        }

        guard (200...299).contains(httpResponse.statusCode) else {
            let message = parseGitHubError(from: data) ?? "GitHub API error \(httpResponse.statusCode)."
            throw NSError(domain: "remou", code: httpResponse.statusCode, userInfo: [NSLocalizedDescriptionKey: message])
        }
    }

    private func parseGitHubError(from data: Data) -> String? {
        guard
            let json = try? JSONSerialization.jsonObject(with: data) as? [String: Any]
        else {
            return nil
        }

        let message = json["message"] as? String
        let errors = (json["errors"] as? [Any])?.map { error in
            if let stringError = error as? String {
                return stringError
            }

            if
                let dictionaryError = error as? [String: Any],
                let detail = dictionaryError["message"] as? String
            {
                let field = dictionaryError["field"] as? String
                return field.map { "\($0): \(detail)" } ?? detail
            }

            return nil
        }
        .compactMap { $0 }
        .joined(separator: "; ")

        return [message, errors].compactMap { value in
            guard let value, !value.isEmpty else { return nil }
            return value
        }
        .joined(separator: " — ")
    }

    private static func javaScriptStringLiteral(_ value: String) -> String {
        let data = try? JSONSerialization.data(withJSONObject: [value], options: [])
        guard
            let data,
            let json = String(data: data, encoding: .utf8),
            json.count >= 2
        else {
            return "\"\""
        }

        return String(json.dropFirst().dropLast())
    }

}
