# Privacy Policy — Teams Message Export

**Last updated:** 12 March 2026

## Overview

Teams Message Export is a Chrome extension that lets you select and export Microsoft Teams messages to Markdown or HTML. This privacy policy explains what data the extension accesses, how it is used, and how it is stored.

## Data Collection

**Teams Message Export does not collect, transmit, or store any personal data.** The extension operates entirely within your browser and does not communicate with any external servers, analytics services, or third-party APIs.

## Data Access

The extension accesses the following data solely within your browser tab:

- **Microsoft Teams page content** — the extension reads the visible DOM of Microsoft Teams web pages (`teams.microsoft.com` and `teams.cloud.microsoft`) to extract message text, author names, timestamps, reactions, and quoted replies.
- **Active tab information** — the extension uses the `tabs` permission to inject its content script into Microsoft Teams tabs.

All accessed data remains local to your browser session. No data is sent to any remote server.

## Data Storage

The extension does not persist any data between browser sessions. Exported files (Markdown or HTML) are saved to your local device through the browser's standard download mechanism. No data is stored in browser local storage, sync storage, cookies, or any cloud service.

## Data Sharing

The extension does not share any data with third parties. There is no telemetry, analytics, crash reporting, or usage tracking of any kind.

## Permissions

| Permission | Purpose |
|---|---|
| `tabs` | Detect when a Microsoft Teams tab is active so the extension icon can trigger the content script |
| Host access to `teams.microsoft.com` and `teams.cloud.microsoft` | Inject the content script that enables message selection and export on Microsoft Teams pages |

## Changes to This Policy

If this privacy policy is updated, the changes will be posted to this page with an updated "Last updated" date. Continued use of the extension after changes constitutes acceptance of the revised policy.

## Contact

If you have questions about this privacy policy, please open an issue on the [GitHub repository](https://github.com/Maxim-Mazurok/teams-selected-messages-export/issues).
