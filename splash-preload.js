const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('splashAPI', {
  onProgress: (callback) => {
    ipcRenderer.on('splash-progress', (_event, data) => callback(data));
  },
  onDone: (callback) => {
    ipcRenderer.on('splash-done', () => callback());
  },
});
