# Project Cost Management Dashboard

A professional desktop application for managing project costs, built with Electron and ExcelJS.

## Features

- Add and manage multiple projects per month
- Auto-calculates engineer cost, CE visit cost, overhead, total cost and profit
- Exports data to Excel (.xlsx)
- Auto-updates via GitHub Releases

## Getting Started

```bash
npm install
npm start
```

## Building a Release

```bash
# Bump version in package.json, then:
git add .
git commit -m "Release v1.x.x"
git tag v1.x.x
git push origin main
git push origin v1.x.x
```

GitHub Actions will automatically build and publish installers for Windows, macOS and Linux.

## Tech Stack

- [Electron](https://www.electronjs.org/)
- [ExcelJS](https://github.com/exceljs/exceljs)
- [electron-updater](https://www.electron.build/auto-update)
- Bootstrap 5
