// เครื่องมือทดสอบเครื่องอ่านบัตรประชาชนแบบเดี่ยวๆ (ไม่เชื่อมกับระบบ POS)
//
// ใช้สำหรับทดสอบว่าเครื่องอ่านบัตร/พอร์ต USB/บัตรใบไหน อ่านผ่านได้จริงหรือไม่ โดยไม่ต้อง
// เปิดโปรแกรม POS หลักหรือระบบอ่านบัตรแยกไปด้วย — รันตัวช่วย ThaiIdReader.exe เอง
//
// ⚠️ ห้ามเปิดพร้อมกับโปรแกรม POS หลัก หรือ card-reader-app เด็ดขาด (จะแย่งกันคุยกับเครื่อง
// อ่านบัตรจนอ่านไม่ออกเลยทั้งคู่ — เจอปัญหานี้มาแล้วจริง) ใช้เครื่องมือนี้แยกต่างหากเท่านั้น

const { app, BrowserWindow } = require("electron");
const path = require("path");
const fs = require("fs");
const readline = require("readline");
const { spawn } = require("child_process");

let mainWindow;
let child;

function exePath() {
  return path.join(__dirname, "..", "card-reader-bin", "ThaiIdReader.exe");
}

function send(channel, payload) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send(channel, payload);
  }
}

function startReader() {
  const p = exePath();
  if (!fs.existsSync(p)) {
    send("test:log", { level: "error", text: "ไม่พบไฟล์ ThaiIdReader.exe ที่ " + p });
    return;
  }
  child = spawn(p, ["--debug"], { windowsHide: true });

  const rl = readline.createInterface({ input: child.stdout });
  rl.on("line", (line) => {
    let msg;
    try {
      msg = JSON.parse(line);
    } catch (e) {
      return;
    }
    send("test:event", msg);
  });

  const rlErr = readline.createInterface({ input: child.stderr });
  rlErr.on("line", (line) => {
    send("test:log", { level: "debug", text: line });
  });

  child.on("exit", (code) => {
    send("test:log", { level: "error", text: `โปรแกรมอ่านบัตรหยุดทำงาน (code ${code})` });
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 700,
    height: 900,
    backgroundColor: "#141714",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  mainWindow.setMenu(null);
  mainWindow.loadFile(path.join(__dirname, "renderer", "index.html"));
}

app.whenReady().then(() => {
  createWindow();
  startReader();
});

app.on("window-all-closed", () => {
  if (child && !child.killed) {
    try {
      child.kill();
    } catch (e) {}
  }
  if (process.platform !== "darwin") app.quit();
});
