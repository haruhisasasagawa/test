# CLAUDE.md

This file provides guidance for AI assistants working with this codebase.

## Project Overview

A simple browser-based reminder application (シンプルリマインダーアプリ) written in Japanese. Users can set reminders with titles and datetime, receive browser notifications when reminders trigger, and manage their reminder list. All data persists in `localStorage`.

## Repository Structure

```
.
├── index.html    # Main HTML — app entry point, single-page UI
├── style.css     # All styles — no preprocessor, plain CSS with CSS custom properties
├── app.js        # All application logic — no build step, plain ES6 JavaScript
└── README.md     # User-facing documentation (Japanese)
```

## Technology Stack

- **Pure HTML/CSS/JavaScript** — no frameworks, no bundler, no package manager
- **No dependencies** — no `node_modules`, no `package.json`
- **Browser APIs used:**
  - `localStorage` for persistence
  - `Notification` API for browser notifications
  - `setInterval` for polling reminder times (1-second tick)
  - DOM manipulation via `getElementById`, `innerHTML`

## Architecture

### Data Model

Reminders are stored as a JSON array in `localStorage` under the key `reminders`. Each object has:

```js
{
  id: Date.now(),      // Unix timestamp used as unique ID
  title: string,       // User-provided reminder title
  time: string,        // ISO 8601 datetime string (from datetime-local input)
  notified: boolean    // Whether the notification has already fired
}
```

### Key Functions (`app.js`)

| Function | Purpose |
|---|---|
| `loadReminders()` | Reads from `localStorage` on page load |
| `saveReminders()` | Writes current state to `localStorage` |
| `addReminder()` | Validates input, creates new reminder object, persists |
| `deleteReminder(id)` | Filters out by `id`, persists, re-renders |
| `displayReminders()` | Sorts by time, renders HTML into `#remindersList` |
| `checkReminders()` | Called every second; fires notifications for due reminders |
| `showNotification(reminder)` | Sends browser `Notification` + `alert()` fallback |
| `escapeHtml(text)` | XSS prevention via `div.textContent` trick |
| `formatDateTime(date)` | Formats Date to Japanese locale string |

### Event Flow

1. `DOMContentLoaded` → `loadReminders()` → `displayReminders()` → `requestNotificationPermission()`
2. 1-second interval starts: `checkReminders()` compares `Date.now()` against each reminder
3. User input → `addReminder()` → validates → pushes to array → `saveReminders()` → `displayReminders()`
4. Delete click → `deleteReminder(id)` → filters array → `saveReminders()` → `displayReminders()`

## Development Workflow

### Running the App

Open `index.html` directly in a browser — no server required:

```bash
open index.html        # macOS
xdg-open index.html    # Linux
start index.html       # Windows
```

Or serve locally to avoid any browser security restrictions on `file://`:

```bash
python3 -m http.server 8080
# Then open http://localhost:8080
```

### No Build Step

There is no compilation, transpilation, or bundling. Edit files and reload the browser.

### Testing

There is no test suite. Manual browser testing is the only option currently.

## Coding Conventions

- **Language:** Comments and variable names are in English; UI text is in Japanese
- **No semicolons policy:** Not enforced — `app.js` uses semicolons
- **DOM manipulation:** Direct `innerHTML` assignment (not `createElement`) for list rendering
- **XSS prevention:** User-supplied text must always go through `escapeHtml()` before insertion into `innerHTML`
- **ID-based lookup:** Elements accessed via `document.getElementById()` with hardcoded IDs from `index.html`
- **State:** A single module-level `reminders` array is the source of truth; always call `saveReminders()` after mutation
- **No modules:** All code is in one file (`app.js`), loaded at the bottom of `<body>`

## Key Constraints

- **No future-date bypass:** `addReminder()` rejects times in the past — keep this validation
- **Notification fallback:** `showNotification()` always fires `alert()` as a fallback even when browser notifications are granted — this is intentional
- **`notified` flag:** Once set to `true`, a reminder never fires again even if `checkReminders()` keeps running; never reset this flag
- **Default time:** On load and after each add, `#reminderTime` is set to "current time + 1 hour, rounded to :00 minutes"

## Browser Compatibility

Targets modern evergreen browsers (Chrome, Firefox, Safari, Edge). The `Notification` API requires explicit user permission; some browsers block it on `file://` URLs.

## Git

Single commit history. Development branch: `claude/claude-md-mlr4t5hkp9dfnlwl-UizLc`. The `master` branch is the main branch.
