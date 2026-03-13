# GitHub & Auto-Update Setup Guide

## Step 1 — Install the new dependency

Open a terminal in your project folder and run:

```bash
npm install electron-updater electron-log
```

## Step 2 — Push your code to GitHub

Run these commands in your project folder:

```bash
git init
git add .
git commit -m "Initial commit with auto-update support"
git branch -M main
git remote add origin https://github.com/shanakarajapakshe/project-cost-management-dashboard.git
git push -u origin main
```

## Step 3 — Create a GitHub Personal Access Token

1. Go to https://github.com/settings/tokens
2. Click **"Generate new token (classic)"**
3. Give it a name like `electron-release`
4. Check the **`repo`** scope (full control of private repositories)
5. Click **Generate token**
6. **Copy the token** — you won't see it again

## Step 4 — Add the token to GitHub Secrets

1. Go to your repo: https://github.com/shanakarajapakshe/project-cost-management-dashboard
2. Click **Settings → Secrets and variables → Actions**
3. Click **"New repository secret"**
4. Name: `GH_TOKEN`
5. Value: paste the token from Step 3
6. Click **Add secret**

## Step 5 — Release your first version

Every time you want to release an update:

```bash
# 1. Update version in package.json (e.g. "version": "1.0.1")
# 2. Commit your changes
git add .
git commit -m "Release v1.0.1"

# 3. Tag the release — this triggers GitHub Actions to build & publish
git tag v1.0.1
git push origin main --tags
```

GitHub Actions will automatically:
- Build the app for Windows (.exe), macOS (.dmg), and Linux (.AppImage)
- Create a GitHub Release with all installers attached
- Publish the `latest.yml` update manifest

## Step 6 — How auto-update works for users

When users run the installed app:
1. On startup, the app silently checks GitHub Releases for a newer version
2. If a new version is found, a blue bar appears at the bottom: **"Update v1.x.x available"**
3. User clicks **Download** — progress shown in the bar
4. When done: **"Restart Now"** button appears
5. App restarts and installs the update automatically

## Files changed in this update

| File | What changed |
|------|-------------|
| `package.json` | Added `electron-updater`, GitHub publish config |
| `updater.js` | **New** — handles all auto-update logic |
| `main.js` | Wires in `setupAutoUpdater()` on app ready |
| `preload.js` | Exposes updater IPC to renderer |
| `app.js` | Update status bar UI |
| `index.html` | Update status bar HTML |
| `.github/workflows/release.yml` | **New** — GitHub Actions build pipeline |
| `.gitignore` | **New** — excludes node_modules, dist |

## Troubleshooting

**Build fails with "GH_TOKEN not found"**
→ Make sure you added the secret in Step 4 with exactly the name `GH_TOKEN`

**Auto-update not working in development**
→ Auto-update only works in a packaged (installed) build, not `npm start`
→ Test by running `npm run build` and installing the output

**"Target not found" error on build**
→ Make sure you ran `npm install electron-updater electron-log` first
