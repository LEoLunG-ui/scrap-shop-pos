const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("scrapAPI", {
  loadDataSync: () => ipcRenderer.sendSync("data:loadSync"),
  loadData: () => ipcRenderer.invoke("data:load"),
  saveData: (data) => ipcRenderer.invoke("data:save", data),
  exportCSV: (filename, content) => ipcRenderer.invoke("csv:export", { filename, content }),
  dataDir: () => ipcRenderer.invoke("app:dataDir"),
  openMonthlyFolder: () => ipcRenderer.invoke("app:openMonthlyFolder"),
  printReceipt: (mmWidth) => ipcRenderer.invoke("print:receipt", mmWidth),
  saveReceiptImage: (rect, billNo, dateISO) => ipcRenderer.invoke("receipt:saveImage", { rect, billNo, dateISO }),
  netInfo: () => ipcRenderer.invoke("net:info"),
  driveStatus: () => ipcRenderer.invoke("drive:status"),
  driveSetCredentials: (clientId, clientSecret) => ipcRenderer.invoke("drive:setCredentials", { clientId, clientSecret }),
  driveConnect: () => ipcRenderer.invoke("drive:connect"),
  driveDisconnect: () => ipcRenderer.invoke("drive:disconnect"),
  driveBackupNow: () => ipcRenderer.invoke("drive:backupNow"),
  onCardReaderEvent: (callback) => {
    const handler = (_e, msg) => callback(msg);
    ipcRenderer.on("cardreader:event", handler);
    return () => ipcRenderer.removeListener("cardreader:event", handler);
  },
  appVersion: () => ipcRenderer.invoke("app:version"),
  installUpdate: () => ipcRenderer.invoke("app:installUpdate"),
  onUpdateDownloaded: (callback) => {
    const handler = (_e, info) => callback(info);
    ipcRenderer.on("update:downloaded", handler);
    return () => ipcRenderer.removeListener("update:downloaded", handler);
  },
});
