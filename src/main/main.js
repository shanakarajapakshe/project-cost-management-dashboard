'use strict';

const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');

// ── Auto-updater (gracefully skipped outside packaged builds) ──────────────
let autoUpdater = null;
try {
  autoUpdater = require('electron-updater').autoUpdater;
} catch (_) {
  // electron-updater not available in dev / uninstalled state
}

const ExcelService = require('../services/excelService');
const { setupUpdater } = require('./updater');
const { registerIpcHandlers } = require('./ipcHandlers');

// ── Globals ────────────────────────────────────────────────────────────────
let mainWindow = null;
let excelService = null;

// Prevent second instances
const gotTheLock = app.requestSingleInstanceLock();
if (!gotTheLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.focus();
    }
  });
}

// ── Window creation ────────────────────────────────────────────────────────
function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 900,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
      preload: path.join(__dirname, 'preload.js')
    },
    icon: path.join(__dirname, '../../assets/icons/icon.ico'),
    title: 'Profit Dashboard',
    backgroundColor: '#f5f7fa',
    show: false // Avoid white flash on startup
  });

  mainWindow.loadFile(path.join(__dirname, '../renderer/index.html'));

  // Show window gracefully once ready
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
    if (process.env.NODE_ENV === 'development') {
      mainWindow.webContents.openDevTools();
    }
  });

  // Intercept navigation to external links – open in browser, not in app
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// ── App lifecycle ──────────────────────────────────────────────────────────
app.whenReady().then(() => {
  // Initialise Excel storage directory
  const userDataPath = app.getPath('userData');
  const excelFolderPath = path.join(userDataPath, 'excel-files');

  try {
    fs.mkdirSync(excelFolderPath, { recursive: true });
  } catch (err) {
    console.error('[main] Failed to create excel-files directory:', err);
  }

  excelService = new ExcelService(excelFolderPath);

  // Register all IPC handlers
  registerIpcHandlers(ipcMain, { app, dialog, excelService, mainWindowGetter: () => mainWindow });

  createWindow();

  // Start auto-update check after window is ready
  if (autoUpdater) {
    mainWindow.once('ready-to-show', () => {
      setupUpdater(autoUpdater, mainWindow);
    });
  }

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// ── Crash / unhandled rejection guard ─────────────────────────────────────
process.on('uncaughtException', (err) => {
  console.error('[main] Uncaught exception:', err);
  dialog.showErrorBox('Unexpected Error', err.message || String(err));
});

process.on('unhandledRejection', (reason) => {
  console.error('[main] Unhandled rejection:', reason);
});
