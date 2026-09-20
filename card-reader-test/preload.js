const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("testApp", {
  onEvent: (callback) => ipcRenderer.on("test:event", (_e, msg) => callback(msg)),
  onLog: (callback) => ipcRenderer.on("test:log", (_e, msg) => callback(msg)),
});
