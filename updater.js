const { autoUpdater } = require('electron-updater');
const { dialog, BrowserWindow } = require('electron');
const log = require('electron-log');

// Configure logging
autoUpdater.logger = log;
autoUpdater.logger.transports.file.level = 'info';

// Don't auto-download — ask user first
autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = true;

function setupAutoUpdater(mainWindow) {
    // Check for updates silently on startup (after 3 seconds)
    setTimeout(() => {
        autoUpdater.checkForUpdates().catch(err => {
            log.warn('Update check failed:', err.message);
        });
    }, 3000);

    // ── Events ──────────────────────────────────────────────────

    autoUpdater.on('checking-for-update', () => {
        log.info('Checking for update...');
        sendStatus(mainWindow, { type: 'checking' });
    });

    autoUpdater.on('update-available', (info) => {
        log.info('Update available:', info.version);
        sendStatus(mainWindow, { type: 'available', version: info.version });

        dialog.showMessageBox(mainWindow, {
            type: 'info',
            title: 'Update Available',
            message: `Version ${info.version} is available.`,
            detail: 'Would you like to download and install it now?',
            buttons: ['Download Now', 'Later'],
            defaultId: 0,
            cancelId: 1
        }).then(({ response }) => {
            if (response === 0) {
                autoUpdater.downloadUpdate();
                sendStatus(mainWindow, { type: 'downloading', percent: 0 });
            }
        });
    });

    autoUpdater.on('update-not-available', () => {
        log.info('App is up to date.');
        sendStatus(mainWindow, { type: 'up-to-date' });
    });

    autoUpdater.on('download-progress', (progress) => {
        const percent = Math.round(progress.percent);
        log.info(`Download progress: ${percent}%`);
        sendStatus(mainWindow, { type: 'downloading', percent });
    });

    autoUpdater.on('update-downloaded', (info) => {
        log.info('Update downloaded:', info.version);
        sendStatus(mainWindow, { type: 'downloaded', version: info.version });

        dialog.showMessageBox(mainWindow, {
            type: 'info',
            title: 'Update Ready',
            message: `Version ${info.version} has been downloaded.`,
            detail: 'The app will restart to apply the update.',
            buttons: ['Restart Now', 'Later'],
            defaultId: 0,
            cancelId: 1
        }).then(({ response }) => {
            if (response === 0) {
                autoUpdater.quitAndInstall();
            }
        });
    });

    autoUpdater.on('error', (err) => {
        log.error('Auto-updater error:', err);
        sendStatus(mainWindow, { type: 'error', message: err.message });
    });
}

function sendStatus(mainWindow, payload) {
    if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('updater:status', payload);
    }
}

module.exports = { setupAutoUpdater, autoUpdater };
