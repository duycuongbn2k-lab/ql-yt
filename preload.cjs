const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  getAppVersion: () => ipcRenderer.invoke('get-app-version'),
  downloadAndInstallUpdate: (url) => ipcRenderer.invoke('download-and-install-update', url),
  onUpdateProgress: (callback) => {
    const subscription = (event, percent) => callback(percent);
    ipcRenderer.on('update-progress', subscription);
    return () => ipcRenderer.removeListener('update-progress', subscription);
  },
  onUpdateError: (callback) => {
    const subscription = (event, message) => callback(message);
    ipcRenderer.on('update-error', subscription);
    return () => ipcRenderer.removeListener('update-error', subscription);
  },
  onUpdateFinished: (callback) => {
    const subscription = (event) => callback();
    ipcRenderer.on('update-finished', subscription);
    return () => ipcRenderer.removeListener('update-finished', subscription);
  }
});
