# Teams Selected Messages Export

A Chrome extension that lets you select and export Microsoft Teams messages to Markdown or HTML — right from the Teams web app.

> **Disclaimer:** This project was originally created using the OpenAI Codex app with GPT 5.4 and then finalized using Claude Opus 4.6 copilot agent mode in VS Code.

## Features

- **Click-to-select messages** — click individual messages or shift-click to select a range
- **Visible checkboxes** with shift-click range support
- **Copy to clipboard** — quick-copy selected messages as Markdown
- **Download exports** — save selected messages as Markdown or HTML files
- **Full chat history export** — scroll-harvest the entire conversation and download as Markdown or HTML
- **Rich formatting** — inline `@mentions`, reply blockquotes, reactions, and image placeholders
- **Theme-aware UI** — follows Teams light/dark theme automatically
- **Non-intrusive** — export controls embed into the Teams title bar

## Screenshots

![Selecting messages](artifacts/screenshots/screenshot-1.png)
![Export panel](artifacts/screenshots/screenshot-2.png)
![Markdown export](artifacts/screenshots/screenshot-3.png)
![HTML export](artifacts/screenshots/screenshot-4.png)

## Install

1. Download the latest `teams-selected-messages-export-extension.zip` from [Releases](../../releases/latest).
2. Unzip the archive.
3. Open `chrome://extensions` in Chrome.
4. Enable **Developer mode** (top-right toggle).
5. Click **Load unpacked** and select the unzipped folder.
6. Open [Microsoft Teams](https://teams.microsoft.com) in Chrome.

## Usage

1. Click the **Export** button in the Teams title bar (top-right, next to settings/avatar).
2. The export panel opens and selection mode activates automatically.
3. Click messages to select them. Hold **Shift** and click to select a range.
4. Use the panel actions:

| Action | Description |
|--------|-------------|
| **Copy MD** | Copy selected messages as Markdown to clipboard |
| **Clear Selection** | Deselect all messages |
| **Download MD** | Download selected messages as a Markdown file |
| **Download HTML** | Download selected messages as an HTML file |
| **Full chat MD** | Export the entire chat history as Markdown |
| **Full chat HTML** | Export the entire chat history as HTML |

Copying Markdown closes the panel and stops selection mode, but preserves your selection. Reopening the panel restores the selected messages.

## Current limitations

- Reactions show emoji, name, and count but actor names are best-effort (DOM-scraped).
- Export depends on the visible Teams DOM and current row selectors.
- Full chat history export relies on DOM scrolling rather than a network-level data source.
