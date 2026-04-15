//
//  AppDelegate.swift
//  semai
//
//  Created by Yago Arconada on 11/13/25.
//

import Cocoa
import Foundation
import SafariServices

private let analyticsURLString = "https://script.google.com/macros/s/AKfycbxTfl5yMdqcHAmXjRmvGIBQ_ILMmR6XX7vyUolxjvR2h17f0cCbsUAFnEdqcw9GMNfL/exec"
private let dailyAnalyticsDateKey = "last_daily_active_ping_date"
private let safariExtensionBundleIdentifier = "yam.team.remou.Extension"

@main
class AppDelegate: NSObject, NSApplicationDelegate {

    func applicationDidFinishLaunching(_ notification: Notification) {
        sendPing(event: "launch")
        sendDailyPingIfNeeded()
    }

    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool {
        return true
    }

    private func loadOrCreateInstallID() -> String {
        let key = "install_id"
        if let existing = UserDefaults.standard.string(forKey: key) {
            return existing
        }

        let newID = UUID().uuidString
        UserDefaults.standard.set(newID, forKey: key)
        return newID
    }

    private func sendDailyPingIfNeeded() {
        let today = Self.analyticsDayString(for: Date())
        if UserDefaults.standard.string(forKey: dailyAnalyticsDateKey) == today {
            return
        }

        SFSafariExtensionManager.getStateOfSafariExtension(withIdentifier: safariExtensionBundleIdentifier) { state, error in
            guard error == nil, state?.isEnabled == true else {
                return
            }

            self.sendPing(event: "daily_active")
            UserDefaults.standard.set(today, forKey: dailyAnalyticsDateKey)
        }
    }

    private func sendPing(event: String) {
        guard let url = URL(string: analyticsURLString) else {
            return
        }

        let payload: [String: Any] = [
            "install_id": loadOrCreateInstallID(),
            "app_version": Bundle.main.infoDictionary?["CFBundleShortVersionString"] as? String ?? "unknown",
            "app_build": Bundle.main.infoDictionary?["CFBundleVersion"] as? String ?? "unknown",
            "platform": ProcessInfo.processInfo.operatingSystemVersionString,
            "event": event
        ]

        var request = URLRequest(url: url)
        request.httpMethod = "POST"
        request.setValue("application/json", forHTTPHeaderField: "Content-Type")
        request.httpBody = try? JSONSerialization.data(withJSONObject: payload)

        URLSession.shared.dataTask(with: request) { _, _, error in
            if let error {
                print("Ping failed:", error)
            }
        }.resume()
    }

    private static func analyticsDayString(for date: Date) -> String {
        let formatter = DateFormatter()
        formatter.calendar = Calendar(identifier: .gregorian)
        formatter.locale = Locale(identifier: "en_US_POSIX")
        formatter.timeZone = .current
        formatter.dateFormat = "yyyy-MM-dd"
        return formatter.string(from: date)
    }
}
