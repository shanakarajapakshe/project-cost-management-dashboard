const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const ExcelService = require('./excelService');
const { setupAutoUpdater, autoUpdater } = require('./updater');

let mainWindow;
let excelService;

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 1400,
        height: 900,
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.join(__dirname, 'preload.js')
        },
        icon: path.join(__dirname, 'icon.png'),
        title: 'Profit Dashboard',
        backgroundColor: '#f5f7fa'
    });

    mainWindow.loadFile('index.html');

    // Open DevTools in development
    if (process.env.NODE_ENV === 'development') {
        mainWindow.webContents.openDevTools();
    }

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

app.whenReady().then(() => {
    // Initialize Excel Service
    const userDataPath = app.getPath('userData');
    const excelFolderPath = path.join(userDataPath, 'excel-files');
    
    // Create excel-files directory if it doesn't exist
    if (!fs.existsSync(excelFolderPath)) {
        fs.mkdirSync(excelFolderPath, { recursive: true });
    }

    excelService = new ExcelService(excelFolderPath);
    
    createWindow();
    setupAutoUpdater(mainWindow);

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});

// IPC Handlers for Excel Operations
ipcMain.handle('excel:loadOrCreate', async (event, month, year) => {
    try {
        const data = await excelService.loadOrCreateMonthFile(month, year);
        return { success: true, data };
    } catch (error) {
        console.error('Error loading/creating Excel file:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('excel:saveProject', async (event, month, year, projectData) => {
    try {
        console.log(`[Excel Service] Attempting to save project: ${projectData.projectName}`);
        console.log(`[Excel Service] Target file: ${excelService.getFilePath(month, year)}`);
        
        await excelService.saveProject(month, year, projectData);
        
        console.log(`[Excel Service] ✓ Project "${projectData.projectName}" saved successfully`);
        return { success: true };
    } catch (error) {
        console.error('[Excel Service] ✗ Error saving project:', error.message);
        console.error('[Excel Service] Stack trace:', error.stack);
        return { 
            success: false, 
            error: error.message,
            stack: error.stack
        };
    }
});

ipcMain.handle('excel:deleteProject', async (event, month, year, projectName) => {
    try {
        await excelService.deleteProject(month, year, projectName);
        return { success: true };
    } catch (error) {
        console.error('Error deleting project:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('excel:getProjects', async (event, month, year) => {
    try {
        const projects = await excelService.getProjects(month, year);
        return { success: true, data: projects };
    } catch (error) {
        console.error('Error getting projects:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('excel:exportFile', async (event, month, year) => {
    try {
        const result = await dialog.showSaveDialog(mainWindow, {
            title: 'Export Excel File',
            defaultPath: `Profit_Dashboard_${year}_${month}.xlsx`,
            filters: [
                { name: 'Excel Files', extensions: ['xlsx'] }
            ]
        });

        if (!result.canceled && result.filePath) {
            await excelService.exportFile(month, year, result.filePath);
            return { success: true, path: result.filePath };
        }
        
        return { success: false, canceled: true };
    } catch (error) {
        console.error('Error exporting file:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('excel:createBackup', async (event, month, year) => {
    try {
        const backupPath = await excelService.createBackup(month, year);
        return { success: true, path: backupPath };
    } catch (error) {
        console.error('Error creating backup:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('dialog:showMessage', async (event, options) => {
    return await dialog.showMessageBox(mainWindow, options);
});

ipcMain.handle('app:getPath', async (event, name) => {
    return app.getPath(name);
});

ipcMain.handle('app:getExcelPath', async (event, month, year) => {
    return excelService.getFilePath(month, year);
});

ipcMain.handle('app:openDataFolder', async (event) => {
    const userDataPath = app.getPath('userData');
    const excelFolderPath = path.join(userDataPath, 'excel-files');
    
    // Create folder if it doesn't exist
    if (!fs.existsSync(excelFolderPath)) {
        fs.mkdirSync(excelFolderPath, { recursive: true });
    }
    
    // Open folder in file explorer
    await shell.openPath(excelFolderPath);
    
    return { success: true, path: excelFolderPath };
});

// IPC: manual update check triggered from renderer
ipcMain.handle('updater:check', async () => {
    try {
        await autoUpdater.checkForUpdates();
        return { success: true };
    } catch (err) {
        return { success: false, error: err.message };
    }
});
