# Profit Dashboard – Project Cost & Profit Management System

A professional **offline-first desktop application** built with Electron, Node.js, ExcelJS, Bootstrap 5, and Chart.js for managing project costs, revenues, and profitability – all stored in Excel files.

---

## Table of Contents

1. [Features](#features)
2. [Project Structure](#project-structure)
3. [Bug Fixes](#bug-fixes)
4. [Installation & Development](#installation--development)
5. [Building the Application](#building-the-application)
6. [Auto-Update System (GitHub Releases)](#auto-update-system-github-releases)
7. [Recommended Improvements](#recommended-improvements)

---

## Features

- 📊 **Interactive Dashboard** – four live Chart.js charts (cost vs payment, profit per project, monthly trend, overhead scatter)
- 📝 **Add / Edit / Delete projects** – full CRUD with live calculation preview
- 💾 **Excel storage** – one `.xlsx` file per month, stored in the OS user-data folder
- 📤 **Export & Backup** – export to any location; auto-timestamped backups
- 🔄 **Auto-updates** – via `electron-updater` + GitHub Releases
- 🖨️ **Print / PDF** – optimised print stylesheet
- 🌐 **Offline-first** – no internet required for core functionality

---

## Project Structure

```
profit-dashboard/
├── src/
│   ├── main/
│   │   ├── main.js          # Electron main process
│   │   ├── preload.js       # Context-bridge (secure IPC surface)
│   │   ├── ipcHandlers.js   # All IPC handler registrations
│   │   └── updater.js       # electron-updater setup & events
│   ├── renderer/
│   │   ├── index.html       # App UI shell
│   │   ├── app.js           # Renderer logic (DataManager, UIController, etc.)
│   │   └── styles.css       # All application styles
│   └── services/
│       └── excelService.js  # ExcelJS read/write/backup service
├── assets/
│   └── icons/               # icon.ico / icon.icns / icon.png
├── package.json
└── README.md
```

---

## Bug Fixes

The following bugs were identified and fixed in this improved version:

| # | Location | Bug | Fix |
|---|----------|-----|-----|
| 1 | `excelService.js` – `saveProject` | `engineerCost` formula was `=C<n>` (just salary), completely ignoring `numEngineers` | Changed to `=B<n>*C<n>` |
| 2 | `excelService.js` – `saveProject` | Row number for formulas used `projectsSheet.rowCount` *before* `addRow()`, producing the wrong row number when projects already existed | Call `addRow()` first, then read `newRow.number` |
| 3 | `excelService.js` – `readProjectsFromWorkbook` | Formula-result cells return `{ formula, result }` objects; code read them as raw values causing `NaN` / `[object Object]` in the UI | Added `_cellVal()` / `_cellNum()` helpers to unwrap result objects |
| 4 | `excelService.js` – `deleteProject` | Only the *last* row matching the project name was deleted (loop didn't break after first match) | Track first match only; throw if not found |
| 5 | `excelService.js` – `updateSummarySheet` | Cross-sheet formula references became stale after `spliceRows()`, so Summary showed wrong totals | Write direct VALUES to Summary, not formulas |
| 6 | `excelService.js` – `createBackup` / `exportFile` | Used sync `fs.copyFileSync` without await; possible race conditions under load | Replaced with `fsP.copyFile` (Promise-based) |
| 7 | `excelService.js` – workbook cache | Cache was not invalidated after `deleteProject` (`spliceRows`) so stale data persisted in memory | `delete this.cache[key]` and reload after destructive operations |
| 8 | `app.js` – `loadFromLocalStorage` | Even in Electron mode, trend data was loaded from `localStorage` (always empty in Electron) | `monthlyData` now lives in memory; no cross-environment `localStorage` reads |
| 9 | `app.js` – missing `updateProject` | Edit functionality was referenced but `updateProject` IPC handler / service method didn't exist | Implemented `updateProject` end-to-end |
| 10 | `preload.js` / `main.js` | `listBackups` was exposed in preload but had no IPC handler | Added handler + service method |
| 11 | `app.js` – `calculateProjectMetrics` | `engineerCost = engineerSalary` (same root cause as bug #1 in service layer) | `engineerCost = numEngineers × engineerSalary` |
| 12 | `index.html` | Inline `onclick="uiController.xxx()"` referenced a global that could be `undefined` | All actions now go through the single `APP` global |
| 13 | `main.js` | `BrowserWindow` shown immediately, causing a white flash | Added `show: false` + `ready-to-show` event |
| 14 | `main.js` | No single-instance lock – multiple app windows could open | Added `app.requestSingleInstanceLock()` |
| 15 | General | No input validation feedback in the UI | Bootstrap `was-validated` class + `invalid-feedback` elements |

---

## Installation & Development

### Prerequisites

- Node.js ≥ 18
- npm ≥ 9

### Setup

```bash
# Clone / extract the project
cd profit-dashboard

# Install dependencies
npm install

# Run in development mode (DevTools open automatically)
npm run dev

# Or run normally
npm start
```

---

## Building the Application

```bash
# Windows installer (NSIS)
npm run build:win

# macOS DMG
npm run build:mac

# Linux AppImage
npm run build:linux

# All platforms (on a macOS CI machine)
npm run build
```

Build output is placed in `dist/`.

### Icons

Place your icon files in `assets/icons/`:
- `icon.ico`  – Windows (256×256)
- `icon.icns` – macOS
- `icon.png`  – Linux (512×512)

You can generate them from a single PNG using [electron-icon-builder](https://www.npmjs.com/package/electron-icon-builder):

```bash
npx electron-icon-builder --input=assets/icons/icon.png --output=assets/icons
```

---

## Auto-Update System (GitHub Releases)

The app uses **`electron-updater`** to automatically check for, download, and install updates via GitHub Releases.

### One-time Setup

1. **Set your repo** in `package.json` → `build.publish`:
   ```json
   "publish": {
     "provider": "github",
     "owner":    "YOUR_GITHUB_USERNAME",
     "repo":     "profit-dashboard"
   }
   ```

2. **Create a GitHub Personal Access Token** with `repo` scope, then set it as an environment variable during builds:
   ```bash
   export GH_TOKEN=ghp_xxxxxxxxxxxxxxxxxxxx
   ```

### Publishing an Update

```bash
# 1. Bump the version
npm version patch   # or minor / major

# 2. Build the installers + publish to GitHub Releases
GH_TOKEN=ghp_xxx npm run build:win   # creates a draft release

# 3. Go to GitHub Releases, add release notes, then publish the release.
```

`electron-builder` automatically creates:
- The platform installer (`.exe` / `.dmg` / `.AppImage`)
- `latest.yml` (Windows), `latest-mac.yml`, `latest-linux.yml`

These files are attached to the GitHub Release and used by `electron-updater` on the client side.

### How updates work in the app

1. On startup (after 3 s), the app silently checks for updates.
2. If an update is found, a blue banner appears at the top.
3. The user clicks **Download** – a progress bar appears.
4. Once downloaded, the **Install & Restart** button appears.
5. The app quits, installs the update, and relaunches automatically.

---

## Recommended Improvements

| Priority | Suggestion |
|----------|------------|
| High | Add a CI pipeline (GitHub Actions) to automate builds and releases |
| High | Sign the Windows installer with a code-signing certificate to avoid SmartScreen warnings |
| Medium | Add unit tests for `excelService.js` calculation logic (Jest) |
| Medium | Add a "Restore from Backup" feature in the UI |
| Medium | Add multi-currency support (currently hard-coded to LKR) |
| Low | Dark mode support using `prefers-color-scheme` |
| Low | Add keyboard shortcuts (e.g., `Ctrl+S` to save, `Ctrl+E` to export) |
| Low | Add a Year Overview screen showing all 12 months on one page |
