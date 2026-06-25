const { app, BrowserWindow } = require('electron');
const path = require('path');

function createWindow() {
  const win = new BrowserWindow({
    width: 1280,
    height: 800,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    },
    title: "Spreadsheet Pro"
  });

  // If we are in dev or if the user ran standard start, load local server
  if (process.env.ELECTRON_DEV === 'true') {
    win.loadURL('http://localhost:3000');
  } else {
    // Load built index.html from dist folder
    win.loadFile(path.join(__dirname, 'dist', 'index.html'));
  }
  
  // Remove default menu bar
  win.setMenuBarVisibility(true);
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
