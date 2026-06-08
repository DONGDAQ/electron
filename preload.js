const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  openFile: (filePath) => ipcRenderer.send('open-file', filePath),
  openFileLocation: (filePath) => ipcRenderer.send('open-file-location', filePath),
  openFolder: (folderPath) => ipcRenderer.send('open-folder', folderPath),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  copyFileTo: (filePath) => ipcRenderer.invoke('copy-file-to', filePath),
  copyToClipboard: (filePath) => ipcRenderer.invoke('copy-to-clipboard', filePath),
  requestDrag: (filePath) => ipcRenderer.send('request-drag', filePath),
});
