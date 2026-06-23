# Project Cost Management Dashboard

A desktop dashboard for tracking project cost, profit, and monthly performance in one place. Built with Electron and ExcelJS, it is designed for teams that want a simple internal tool instead of manual spreadsheet work.

## What It Does

- Manages multiple projects per month
- Calculates engineer cost, CE visit cost, overhead, total cost, and profit automatically
- Exports and updates data in Excel format
- Uses GitHub Releases for desktop app updates
- Runs as a desktop app on Windows, macOS, and Linux

## Why It Matters

- Helps teams track project financials faster
- Reduces spreadsheet mistakes and manual calculations
- Gives managers a clear view of cost and profit data
- Fits internal business, operations, and reporting workflows

## Tech Stack

- Electron
- ExcelJS
- electron-updater
- electron-log
- Bootstrap 5

## Run Locally

```bash
npm install
npm start
```

## Build and Release

Build installers for each platform:

```bash
npm run build-win
npm run build-mac
npm run build-linux
```

Create a release build:

```bash
npm run release
```

## Project Structure

```text
project-cost-management-dashboard/
├── app.js
├── excelService.js
├── main.js
├── preload.js
├── updater.js
├── index.html
├── styles.css
├── assets/
└── README.md
```

## Notes

- The app stores and reads data through Excel files.
- GitHub Releases is used for distribution and update delivery.

## License

MIT
