const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const Excel = require('exceljs');
const isDev = require('electron-is-dev');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js') // Load the preload script
    }
  });

  mainWindow.loadFile('index.html');

  // Open the DevTools in development mode
  if (isDev) {
    mainWindow.webContents.openDevTools();
  }
}

app.whenReady().then(() => {
  createWindow();

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

// Implement the readExcel function here
function readExcel(filePath) {
  const workbook = new Excel.Workbook();
  const data = [];

  return workbook.xlsx.readFile(filePath)
    .then(() => {
      const worksheet = workbook.worksheets[0];

      worksheet.eachRow((row, rowNumber) => {
        data.push(row.values);
      });

      return data;
    })
    .catch((error) => {
      console.error('Error reading Excel file:', error);
      return null;
    });
}

ipcMain.on('chooseExcelFile', (event) => {
  dialog.showOpenDialog(mainWindow, {
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  }).then((filePaths) => {
    if (!filePaths.canceled && filePaths.filePaths.length > 0) {
      const filePath = filePaths.filePaths[0];
      // Read the Excel file and send data back to the renderer process
      readExcel(filePath)
        .then((data) => {
          event.reply('excelData', data);
        })
        .catch((error) => {
          console.error('Error reading Excel file:', error);
          event.reply('excelData', null);
        });
    }
  });
});
