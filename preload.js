const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: () => ipcRenderer.invoke('open-file'),
  processWithTemplate: (data) => ipcRenderer.invoke('process-with-template', data),
  reloadContent: (outputPath) => ipcRenderer.invoke('reload-content', outputPath)
});