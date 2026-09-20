// ระบบอ่านบัตรประชาชน (แยกต่างหากจากโปรแกรม POS หลัก แต่เชื่อมข้อมูลเข้าระบบ POS ได้)
//
// โปรแกรมนี้ไม่ได้เปิด ThaiIdReader.exe ของตัวเอง — ให้โปรแกรม POS หลักเป็นตัวเดียวที่คุยกับ
// เครื่องอ่านบัตรจริง (กันไม่ให้สองโปรแกรมแย่งกันคุยกับเครื่องอ่านพร้อมกันจนอ่านไม่ออกเลย ซึ่ง
// เคยเจอปัญหานี้มาแล้วตอนเปิดทั้งสองโปรแกรมพร้อมกัน) แต่รับข้อมูลบัตรผ่าน WebSocket ของ
// โปรแกรม POS หลักแทน (ช่องทางเดียวกับที่ใช้ซิงก์ข้อมูลสมาชิก/สินค้า) — ต้องเปิดโปรแกรม POS
// หลักไว้ก่อนเสมอ ถึงจะได้รับข้อมูลบัตรได้

const { app, BrowserWindow } = require("electron");
const path = require("path");

function createWindow() {
  const win = new BrowserWindow({
    width: 640,
    height: 900,
    backgroundColor: "#141714",
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
    },
  });
  win.setMenu(null);
  win.loadFile(path.join(__dirname, "renderer", "index.html"));
}

app.whenReady().then(createWindow);

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});
