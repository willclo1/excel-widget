const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: () => ipcRenderer.invoke('open-file'),
  reloadContent: (outputPath) => ipcRenderer.invoke('reload-content', outputPath)
});