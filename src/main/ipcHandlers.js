'use strict';

/**
 * Register all IPC handlers for the main process.
 * Keeping them in a dedicated module keeps main.js lean and testable.
 *
 * @param {Electron.IpcMain} ipcMain
 * @param {{ app, dialog, excelService, mainWindowGetter }} deps
 */
function registerIpcHandlers(ipcMain, { app, dialog, excelService, mainWindowGetter }) {

  // ── Helper: wrap every handler with consistent error shape ──────────────
  function safe(fn) {
    return async (event, ...args) => {
      try {
        const data = await fn(...args);
        return { success: true, data };
      } catch (err) {
        console.error('[ipc]', err);
        return { success: false, error: err.message || String(err) };
      }
    };
  }

  // ── Excel operations ─────────────────────────────────────────────────────
  ipcMain.handle('excel:loadOrCreate', safe(async (month, year) => {
    return await excelService.loadOrCreateMonthFile(month, year);
  }));

  ipcMain.handle('excel:saveProject', safe(async (month, year, projectData) => {
    await excelService.saveProject(month, year, projectData);
    return null; // success is enough
  }));

  ipcMain.handle('excel:updateProject', safe(async (month, year, originalName, projectData) => {
    await excelService.updateProject(month, year, originalName, projectData);
    return null;
  }));

  ipcMain.handle('excel:deleteProject', safe(async (month, year, projectName) => {
    await excelService.deleteProject(month, year, projectName);
    return null;
  }));

  ipcMain.handle('excel:getProjects', safe(async (month, year) => {
    return await excelService.getProjects(month, year);
  }));

  ipcMain.handle('excel:exportFile', async (event, month, year) => {
    try {
      const win = mainWindowGetter();
      const result = await dialog.showSaveDialog(win, {
        title: 'Export Excel File',
        defaultPath: `Profit_Dashboard_${year}_${String(month).padStart(2, '0')}.xlsx`,
        filters: [{ name: 'Excel Files', extensions: ['xlsx'] }]
      });

      if (result.canceled || !result.filePath) {
        return { success: false, canceled: true };
      }

      await excelService.exportFile(month, year, result.filePath);
      return { success: true, data: result.filePath };
    } catch (err) {
      console.error('[ipc] excel:exportFile', err);
      return { success: false, error: err.message };
    }
  });

  ipcMain.handle('excel:createBackup', safe(async (month, year) => {
    return await excelService.createBackup(month, year);
  }));

  ipcMain.handle('excel:listBackups', safe(async (month, year) => {
    return await excelService.listBackups(month, year);
  }));

  // ── Dialog helpers ───────────────────────────────────────────────────────
  ipcMain.handle('dialog:showMessage', async (event, options) => {
    const win = mainWindowGetter();
    return await dialog.showMessageBox(win, options);
  });

  ipcMain.handle('dialog:showError', async (event, title, content) => {
    dialog.showErrorBox(title, content);
    return null;
  });

  // ── App info ─────────────────────────────────────────────────────────────
  ipcMain.handle('app:getPath', (event, name) => app.getPath(name));
  ipcMain.handle('app:getVersion', () => app.getVersion());
}

module.exports = { registerIpcHandlers };
