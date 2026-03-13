'use strict';

const { contextBridge, ipcRenderer } = require('electron');

/**
 * Expose a strongly-typed API surface to the renderer process.
 * Nothing from Node / Electron is exposed directly – only these methods.
 */
contextBridge.exposeInMainWorld('electronAPI', {

  // ── Excel CRUD ───────────────────────────────────────────────────────────
  loadOrCreateExcel:  (month, year)                       => ipcRenderer.invoke('excel:loadOrCreate', month, year),
  saveProject:        (month, year, projectData)          => ipcRenderer.invoke('excel:saveProject',  month, year, projectData),
  updateProject:      (month, year, origName, projectData)=> ipcRenderer.invoke('excel:updateProject', month, year, origName, projectData),
  deleteProject:      (month, year, projectName)          => ipcRenderer.invoke('excel:deleteProject', month, year, projectName),
  getProjects:        (month, year)                       => ipcRenderer.invoke('excel:getProjects',  month, year),
  exportExcelFile:    (month, year)                       => ipcRenderer.invoke('excel:exportFile',   month, year),
  createBackup:       (month, year)                       => ipcRenderer.invoke('excel:createBackup', month, year),
  listBackups:        (month, year)                       => ipcRenderer.invoke('excel:listBackups',  month, year),

  // ── Dialogs ──────────────────────────────────────────────────────────────
  showMessage:        (options)          => ipcRenderer.invoke('dialog:showMessage', options),
  showError:          (title, content)   => ipcRenderer.invoke('dialog:showError',   title, content),

  // ── App info ─────────────────────────────────────────────────────────────
  getAppPath:         (name)  => ipcRenderer.invoke('app:getPath',    name),
  getAppVersion:      ()      => ipcRenderer.invoke('app:getVersion'),

  // ── Auto-updater ─────────────────────────────────────────────────────────
  checkForUpdates:    () => ipcRenderer.invoke('updater:check'),
  downloadUpdate:     () => ipcRenderer.invoke('updater:download'),
  installUpdate:      () => ipcRenderer.invoke('updater:install'),

  /** Listen for updater status pushes from the main process */
  onUpdaterStatus: (callback) => {
    ipcRenderer.on('updater:status', (_event, payload) => callback(payload));
  },

  // ── Utility ──────────────────────────────────────────────────────────────
  isElectron: () => true
});
