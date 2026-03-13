'use strict';

/**
 * Auto-update module using electron-updater + GitHub Releases.
 *
 * Publishing flow:
 *   1. Bump version in package.json
 *   2. Run: npm run build:win  (or mac/linux)
 *   3. Create a GitHub Release tagged vX.Y.Z and attach the
 *      installer + latest.yml (auto-generated in dist/)
 *   4. On next app launch, this module checks for updates automatically.
 */

function setupUpdater(autoUpdater, mainWindow) {
  // ── Configuration ──────────────────────────────────────────────────────
  autoUpdater.autoDownload = false;   // Ask user before downloading
  autoUpdater.autoInstallOnAppQuit = true;

  // ── Helper: send status to renderer ────────────────────────────────────
  function sendStatus(event, data) {
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send('updater:status', { event, ...data });
    }
  }

  // ── Updater events ──────────────────────────────────────────────────────
  autoUpdater.on('checking-for-update', () => {
    sendStatus('checking');
  });

  autoUpdater.on('update-available', (info) => {
    sendStatus('available', { version: info.version, releaseNotes: info.releaseNotes });
  });

  autoUpdater.on('update-not-available', () => {
    sendStatus('not-available');
  });

  autoUpdater.on('download-progress', (progress) => {
    sendStatus('progress', {
      percent: Math.round(progress.percent),
      transferred: progress.transferred,
      total: progress.total,
      bytesPerSecond: progress.bytesPerSecond
    });
  });

  autoUpdater.on('update-downloaded', (info) => {
    sendStatus('downloaded', { version: info.version });
  });

  autoUpdater.on('error', (err) => {
    // Non-fatal – just log and notify renderer; don't crash the app
    console.error('[updater] error:', err.message);
    sendStatus('error', { message: err.message });
  });

  // ── IPC: renderer-initiated actions ─────────────────────────────────────
  const { ipcMain } = require('electron');

  ipcMain.handle('updater:check', async () => {
    try {
      await autoUpdater.checkForUpdates();
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('updater:download', async () => {
    try {
      await autoUpdater.downloadUpdate();
      return { success: true };
    } catch (err) {
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('updater:install', () => {
    autoUpdater.quitAndInstall(false, true);
    return { success: true };
  });

  // ── Initial check (delayed 3 s so the UI is settled) ───────────────────
  setTimeout(() => {
    autoUpdater.checkForUpdates().catch((err) => {
      console.warn('[updater] Initial check failed:', err.message);
    });
  }, 3000);
}

module.exports = { setupUpdater };
