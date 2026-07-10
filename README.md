# Project Cost Management Dashboard

A professional desktop dashboard for tracking project cost, profit, and monthly performance in one place. Built with Electron and ExcelJS, it is designed for teams that want a simple internal tool instead of manual spreadsheet work.

## Overview

Project Cost Management Dashboard helps teams record project-level cost inputs and instantly calculate the financial position for a selected month. It supports month-by-month tracking, automated Excel persistence, summary reporting, visual dashboards, backups, and GitHub Releases-based desktop app publishing.

## What It Does

- Manages multiple projects per month
- Calculates engineer cost, CE visit cost, direct cost, overhead cost, total cost, and profit automatically
- Displays dashboard totals for project count, revenue, costs, and net profit
- Shows charts for project performance and monthly profit trends
- Creates, updates, exports, and backs up Excel workbooks
- Uses GitHub Releases for desktop app updates and installer distribution
- Runs as a desktop app on Windows, macOS, and Linux

## Why It Matters

- Helps teams track project financials faster
- Reduces spreadsheet mistakes and manual calculations
- Gives managers a clear view of cost and profit data
- Fits internal business, operations, and reporting workflows

## Tech Stack

- Electron
- JavaScript
- HTML/CSS
- Bootstrap 5
- ExcelJS
- electron-builder
- electron-updater
- electron-log
- GitHub Actions

## Project Structure

```text
.
|-- app.js                  # Dashboard logic, calculations, UI updates, charts
|-- excelService.js         # Excel workbook creation, reading, writing, export, backup
|-- index.html              # Main application interface
|-- main.js                 # Electron main process and IPC handlers
|-- preload.js              # Secure bridge between Electron and renderer process
|-- styles.css              # Application styling
|-- updater.js              # GitHub Releases auto-update integration
|-- assets/                 # Application icon and assets
|-- dist/                   # Generated installer/build output
`-- .github/workflows/      # GitHub Actions release workflow
```

## Getting Started

### Prerequisites

- Node.js 20 or newer
- npm

### Install Dependencies

```bash
npm install
```

### Run Locally

```bash
npm start
```

## Build

Build a Windows installer:

```bash
npm run build-win
```

Build for the current platform:

```bash
npm run build
```

Generated installers are written to the `dist/` directory.

## Release Through GitHub

This project is configured to publish desktop installers through GitHub Releases.

The release workflow runs when a version tag is pushed:

```bash
git tag v1.1.5
git push origin v1.1.5
```

GitHub Actions then installs dependencies, builds the Electron app on Windows, macOS, and Linux, and publishes the release assets.

Release command used by the workflow:

```bash
npm run release
```

The publish target is configured in `package.json`:

```json
{
  "provider": "github",
  "owner": "shanakarajapakshe",
  "repo": "project-cost-management-dashboard"
}
```

## GitHub Release Token

The workflow expects a GitHub token to publish release assets:

```yaml
GH_TOKEN: ${{ secrets.GH_TOKEN }}
```

Before publishing releases, make sure the repository has a valid `GH_TOKEN` secret with permission to create releases and upload assets. Alternatively, the workflow can be adjusted to use the built-in `GITHUB_TOKEN`.

## Data Storage

For each selected month and year, the application creates or updates an Excel workbook using this naming pattern:

```text
Profit_Dashboard_<year>_<month>.xlsx
```

Each workbook includes project data and summary information. Backup files are also generated to protect monthly records.

## Available Scripts

```bash
npm start           # Run the Electron desktop app
npm run build       # Build the app for the current platform
npm run build-win   # Build the Windows installer
npm run build-mac   # Build the macOS package
npm run build-linux # Build the Linux package
npm run release     # Build and publish release assets
npm run draft       # Build without publishing
```

## Current Release Output

The repository currently contains generated Windows installer output in `dist/`, including versioned setup files for the Profit Dashboard application.

## License

MIT
