const { contextBridge, ipcRenderer } = require("electron");

// Expose IPC communication to the renderer process
contextBridge.exposeInMainWorld("electron", {
  send: (channel, data) => {
    // Use IPC to send data to the main process
    ipcRenderer.send(channel, data);
  },
});
