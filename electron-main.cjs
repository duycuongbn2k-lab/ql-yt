const { app, BrowserWindow, Menu, ipcMain } = require('electron');
const path = require('path');
const fs = require('fs');
const https = require('https');
const { spawn } = require('child_process');

let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 800,
    minWidth: 1024,
    minHeight: 768,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.cjs')
    },
    title: "YT Audit Pro",
    autoHideMenuBar: true
  });

  // Tự động phóng to toàn màn hình khi mở
  mainWindow.maximize();

  // Load the built app's index.html
  const indexPath = path.join(__dirname, 'dist', 'index.html');
  mainWindow.loadFile(indexPath).catch((err) => {
    console.error("Failed to load index.html:", err);
  });

  // Open Developer Tools safely for troubleshooting
  mainWindow.webContents.once('did-finish-load', () => {
    mainWindow.webContents.openDevTools();
  });

  // Remove standard browser menu bar
  Menu.setApplicationMenu(null);

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

// --- IPC HANDLERS FOR AUTO-UPDATER ---
ipcMain.handle('get-app-version', () => {
  return app.getVersion();
});

ipcMain.handle('download-and-install-update', async (event, downloadUrl) => {
  const tempDir = app.getPath('temp');
  const installerPath = path.join(tempDir, 'YT_Audit_Pro_Update.exe');
  
  // Clean up any old installer if exists
  if (fs.existsSync(installerPath)) {
    try {
      fs.unlinkSync(installerPath);
    } catch (e) {
      console.error("Failed to delete old installer:", e);
    }
  }

  const file = fs.createWriteStream(installerPath);
  
  return new Promise((resolve, reject) => {
    function startDownload(url) {
      https.get(url, (response) => {
        // Handle redirects (GitHub uses 302 to AWS S3)
        if (response.statusCode === 302 || response.statusCode === 301) {
          startDownload(response.headers.location);
          return;
        }

        if (response.statusCode !== 200) {
          reject(`Server returned status code: ${response.statusCode}`);
          if (mainWindow) {
            mainWindow.webContents.send('update-error', `Lỗi tải file (Mã lỗi: ${response.statusCode})`);
          }
          return;
        }

        const totalLength = parseInt(response.headers['content-length'], 10) || 0;
        let downloaded = 0;
        
        response.pipe(file);
        
        response.on('data', (chunk) => {
          downloaded += chunk.length;
          if (totalLength > 0 && mainWindow) {
            const percent = Math.round((downloaded / totalLength) * 100);
            mainWindow.webContents.send('update-progress', percent);
          }
        });
        
        file.on('finish', () => {
          file.close(() => {
            if (mainWindow) {
              mainWindow.webContents.send('update-finished');
            }
            
            // Spawn installer and quit app
            setTimeout(() => {
              try {
                const child = spawn(installerPath, ['/S'], {
                  detached: true,
                  stdio: 'ignore'
                });
                child.unref();
                app.quit();
                resolve(true);
              } catch (e) {
                if (mainWindow) {
                  mainWindow.webContents.send('update-error', `Không thể khởi chạy bộ cài đặt: ${e.message}`);
                }
                reject(e.message);
              }
            }, 1500);
          });
        });
      }).on('error', (err) => {
        fs.unlink(installerPath, () => {});
        if (mainWindow) {
          mainWindow.webContents.send('update-error', `Lỗi mạng: ${err.message}`);
        }
        reject(err.message);
      });
    }

    startDownload(downloadUrl);
  });
});

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
