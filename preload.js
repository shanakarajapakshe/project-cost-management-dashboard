const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods that allow the renderer process to use
// the ipcRenderer without exposing the entire object
contextBridge.exposeInMainWorld('electronAPI', {
    // Excel operations
    loadOrCreateExcel: (month, year) => ipcRenderer.invoke('excel:loadOrCreate', month, year),
    saveProject: (month, year, projectData) => ipcRenderer.invoke('excel:saveProject', month, year, projectData),
    deleteProject: (month, year, projectName) => ipcRenderer.invoke('excel:deleteProject', month, year, projectName),
    getProjects: (month, year) => ipcRenderer.invoke('excel:getProjects', month, year),
    exportExcelFile: (month, year) => ipcRenderer.invoke('excel:exportFile', month, year),
    createBackup: (month, year) => ipcRenderer.invoke('excel:createBackup', month, year),
    
    // Dialog operations
    showMessage: (options) => ipcRenderer.invoke('dialog:showMessage', options),
    
    // App operations
    getAppPath: (name) => ipcRenderer.invoke('app:getPath', name),
    
    // Check if running in Electron
    isElectron: () => true
});
