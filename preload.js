const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('dashboardAPI', {
  openFile: () => ipcRenderer.invoke('open-file-dialog'),
  loadDroppedFile: (filePath) => ipcRenderer.invoke('load-dropped-file', filePath),
  onDashboardReady: (callback) => {
    ipcRenderer.on('dashboard-ready', (_event, data) => callback(data));
  },
  onDashboardError: (callback) => {
    ipcRenderer.on('dashboard-error', (_event, data) => callback(data));
  },
  onLoadProgress: (callback) => {
    ipcRenderer.on('load-progress', (_event, data) => callback(data));
  },
});
