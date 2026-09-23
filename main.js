const { app, BrowserWindow, ipcMain, dialog, Menu, shell } = require("electron");
const path = require("path");
const fs = require("fs");
const os = require("os");
const http = require("http");
const readline = require("readline");
const { spawn } = require("child_process");
const express = require("express");
const { WebSocketServer } = require("ws");
const ExcelJS = require("exceljs");
const { google } = require("googleapis");
const { autoUpdater } = require("electron-updater");

const PORT = 4173;

const dataFilePath = () => path.join(app.getPath("userData"), "scrap-pos-data.json");
const excelFilePath = () => path.join(app.getPath("userData"), "scrap-pos-data.xlsx");
const monthlyDirPath = () => path.join(app.getPath("userData"), "รายงานรายเดือน");
const receiptRootDirPath = () => path.join(app.getPath("userData"), "ใบเสร็จ");
const receiptImageDirPath = (dateISO) => {
  const d = new Date(dateISO);
  const ym = isNaN(d) ? "อื่นๆ" : `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
  return path.join(receiptRootDirPath(), ym);
};
const driveConfigPath = () => path.join(app.getPath("userData"), "drive-config.json");
const driveTokenPath = () => path.join(app.getPath("userData"), "drive-token.json");

function loadDataSync() {
  try {
    const raw = fs.readFileSync(dataFilePath(), "utf-8");
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

// คิว serialize งานเขียนไฟล์ Excel กันบันทึกซ้อนกันเร็วเกินไปจนไฟล์ .tmp ชนกัน (ทำให้เขียนไฟล์ล้มเหลว)
let excelWriteQueue = Promise.resolve();
function queueExcelSync(data) {
  excelWriteQueue = excelWriteQueue
    .then(() => regenerateExcel(data))
    .then(() => regenerateMonthlyReports(data))
    .catch((e) => console.error("Excel sync failed:", e));
  return excelWriteQueue;
}

function saveDataSync(data) {
  const tmp = dataFilePath() + ".tmp";
  fs.writeFileSync(tmp, JSON.stringify(data), "utf-8");
  fs.renameSync(tmp, dataFilePath());
  queueExcelSync(data);
  scheduleDriveBackup();
}

function getLanIp() {
  const ifaces = os.networkInterfaces();
  for (const name of Object.keys(ifaces)) {
    for (const iface of ifaces[name] || []) {
      if (iface.family === "IPv4" && !iface.internal) return iface.address;
    }
  }
  return null;
}

/* ============================= DEFAULT SEED DATA =============================
 * ต้องตรงกับ DEFAULT_CATEGORIES / DEFAULT_CATALOG ใน renderer/index.html เสมอ
 * ใช้สำหรับ "เติม" รายการที่ยังไม่มีเท่านั้น (เพิ่มโดยเทียบ id ที่ยังไม่เคยมี)
 * ไม่มีการแก้ไขหรือลบข้อมูลเดิมของผู้ใช้เด็ดขาด — ป้องกันข้อมูลหายตอนอัปเดตโปรแกรม
 */
const DEFAULT_CATEGORIES = [
  { id: "scrap", label: "เศษเหล็ก", icon: "magnet" },
  { id: "paper", label: "กระดาษ", icon: "fileText" },
  { id: "plastic", label: "พลาสติก", icon: "bottle" },
  { id: "precious", label: "โลหะมีค่า", icon: "gem" },
  { id: "glass", label: "แก้ว", icon: "cup" },
  { id: "electronics", label: "เครื่องใช้ไฟฟ้า/อิเล็กทรอนิกส์", icon: "zap" },
  { id: "misc", label: "อื่นๆ", icon: "layers" },
];
const DEFAULT_CATALOG = [
  { id: "scrap-thick", code: "FE01", category: "scrap", name: "เศษเหล็กหนา", unit: "กก.", mode: "weight", priceRegular: 8.5, priceCasual: 8, quick: true },
  { id: "scrap-thin", code: "FE02", category: "scrap", name: "เศษเหล็กบาง", unit: "กก.", mode: "weight", priceRegular: 5, priceCasual: 4.5, quick: true },
  { id: "scrap-rebar", code: "FE03", category: "scrap", name: "เหล็กเส้น/เหล็กโครงสร้าง", unit: "กก.", mode: "weight", priceRegular: 7, priceCasual: 6.5, quick: false },
  { id: "scrap-cast", code: "FE04", category: "scrap", name: "เหล็กหล่อ", unit: "กก.", mode: "weight", priceRegular: 6, priceCasual: 5.5, quick: false },
  { id: "scrap-drum", code: "FE05", category: "scrap", name: "ถังเหล็ก/ถังแก๊ส", unit: "กก.", mode: "weight", priceRegular: 6.5, priceCasual: 6, quick: false },
  { id: "scrap-rusty", code: "FE06", category: "scrap", name: "เศษเหล็กปนสนิม", unit: "กก.", mode: "weight", priceRegular: 4, priceCasual: 3.5, quick: false },
  { id: "scrap-wheel", code: "FE07", category: "scrap", name: "ล้อแม็กเหล็ก", unit: "กก.", mode: "weight", priceRegular: 9, priceCasual: 8.5, quick: false },
  { id: "paper-box", code: "PA01", category: "paper", name: "กระดาษลัง", unit: "กก.", mode: "weight", priceRegular: 4.5, priceCasual: 4, quick: true },
  { id: "paper-bw", code: "PA02", category: "paper", name: "กระดาษขาวดำ", unit: "กก.", mode: "weight", priceRegular: 3.2, priceCasual: 2.8, quick: false },
  { id: "paper-color", code: "PA03", category: "paper", name: "กระดาษสี/กระดาษออฟฟิศ", unit: "กก.", mode: "weight", priceRegular: 3.5, priceCasual: 3, quick: false },
  { id: "paper-magazine", code: "PA04", category: "paper", name: "นิตยสาร/กระดาษมัน", unit: "กก.", mode: "weight", priceRegular: 3, priceCasual: 2.5, quick: false },
  { id: "paper-book", code: "PA05", category: "paper", name: "สมุด/หนังสือเก่า", unit: "กก.", mode: "weight", priceRegular: 2.5, priceCasual: 2, quick: false },
  { id: "plastic-pet", code: "PL01", category: "plastic", name: "ขวดใส PET", unit: "กก.", mode: "weight", priceRegular: 6.5, priceCasual: 6, quick: true },
  { id: "plastic-mix", code: "PL02", category: "plastic", name: "พลาสติกรวม", unit: "กก.", mode: "weight", priceRegular: 3, priceCasual: 2.5, quick: false },
  { id: "plastic-hdpe", code: "PL03", category: "plastic", name: "ขวดขุ่น HDPE", unit: "กก.", mode: "weight", priceRegular: 5, priceCasual: 4.5, quick: false },
  { id: "plastic-bag", code: "PL04", category: "plastic", name: "ถุงพลาสติก", unit: "กก.", mode: "weight", priceRegular: 2, priceCasual: 1.5, quick: false },
  { id: "plastic-crate", code: "PL05", category: "plastic", name: "ลังพลาสติก/เข่ง", unit: "กก.", mode: "weight", priceRegular: 4, priceCasual: 3.5, quick: false },
  { id: "plastic-pvc", code: "PL06", category: "plastic", name: "ท่อ PVC", unit: "กก.", mode: "weight", priceRegular: 3.5, priceCasual: 3, quick: false },
  { id: "copper", code: "NF01", category: "precious", name: "ทองแดงเบอร์ 1", unit: "กก.", mode: "weight", priceRegular: 235, priceCasual: 225, quick: true },
  { id: "copper2", code: "NF02", category: "precious", name: "ทองแดงเบอร์ 2 (ปนฉนวน)", unit: "กก.", mode: "weight", priceRegular: 200, priceCasual: 190, quick: false },
  { id: "brass", code: "NF03", category: "precious", name: "ทองเหลือง", unit: "กก.", mode: "weight", priceRegular: 132, priceCasual: 125, quick: false },
  { id: "aluminum", code: "NF04", category: "precious", name: "อลูมิเนียมเส้น/ลวด", unit: "กก.", mode: "weight", priceRegular: 38, priceCasual: 35, quick: true },
  { id: "aluminum-can", code: "NF05", category: "precious", name: "กระป๋องอลูมิเนียม", unit: "กก.", mode: "weight", priceRegular: 32, priceCasual: 30, quick: true },
  { id: "aluminum-wheel", code: "NF06", category: "precious", name: "ล้อแม็กอลูมิเนียม", unit: "กก.", mode: "weight", priceRegular: 45, priceCasual: 42, quick: false },
  { id: "stainless", code: "NF07", category: "precious", name: "สแตนเลส", unit: "กก.", mode: "weight", priceRegular: 25, priceCasual: 22, quick: false },
  { id: "zinc", code: "NF08", category: "precious", name: "สังกะสี", unit: "กก.", mode: "weight", priceRegular: 30, priceCasual: 28, quick: false },
  { id: "lead", code: "NF09", category: "precious", name: "ตะกั่ว", unit: "กก.", mode: "weight", priceRegular: 28, priceCasual: 26, quick: false },
  { id: "glass-beer", code: "GL01", category: "glass", name: "ขวดเบียร์", unit: "ใบ", mode: "unit", priceRegular: 1, priceCasual: 0.8, quick: true },
  { id: "glass-clear", code: "GL02", category: "glass", name: "ขวดแก้วใส", unit: "กก.", mode: "weight", priceRegular: 1.2, priceCasual: 1, quick: false },
  { id: "glass-color", code: "GL03", category: "glass", name: "ขวดแก้วสี", unit: "กก.", mode: "weight", priceRegular: 1, priceCasual: 0.8, quick: false },
  { id: "glass-mix", code: "GL04", category: "glass", name: "แก้วรวม", unit: "กก.", mode: "weight", priceRegular: 1, priceCasual: 0.8, quick: false },
  { id: "e-wire", code: "EL01", category: "electronics", name: "สายไฟ/สายเคเบิล", unit: "กก.", mode: "weight", priceRegular: 20, priceCasual: 17, quick: true },
  { id: "e-battery-car", code: "EL02", category: "electronics", name: "แบตเตอรี่รถยนต์", unit: "กก.", mode: "weight", priceRegular: 15, priceCasual: 13, quick: false },
  { id: "e-motor", code: "EL03", category: "electronics", name: "มอเตอร์ไฟฟ้า", unit: "กก.", mode: "weight", priceRegular: 12, priceCasual: 10, quick: false },
  { id: "e-pcb", code: "EL04", category: "electronics", name: "แผงวงจร/บอร์ดอิเล็กทรอนิกส์", unit: "กก.", mode: "weight", priceRegular: 25, priceCasual: 20, quick: false },
  { id: "e-fan", code: "EL05", category: "electronics", name: "พัดลมเก่า", unit: "ตัว", mode: "unit", priceRegular: 20, priceCasual: 15, quick: false },
  { id: "e-tv", code: "EL06", category: "electronics", name: "ทีวี/จอเก่า", unit: "เครื่อง", mode: "unit", priceRegular: 30, priceCasual: 25, quick: false },
  { id: "e-battery-phone", code: "EL07", category: "electronics", name: "แบตมือถือ/โน้ตบุ๊ค", unit: "ก้อน", mode: "unit", priceRegular: 5, priceCasual: 3, quick: false },
  { id: "misc-tire", code: "MS01", category: "misc", name: "ยางรถยนต์เก่า", unit: "เส้น", mode: "unit", priceRegular: 10, priceCasual: 8, quick: false },
  { id: "misc-cloth", code: "MS02", category: "misc", name: "เศษผ้า/เสื้อผ้าเก่า", unit: "กก.", mode: "weight", priceRegular: 3, priceCasual: 2.5, quick: false },
];

// ทำครั้งเดียวตอนโปรแกรมเริ่ม: เติมหมวดหมู่/สินค้าเริ่มต้นเฉพาะรายการที่ยังไม่มี (id ใหม่)
// ไม่แตะต้อง/ไม่ลบ/ไม่ทับข้อมูลเดิมของผู้ใช้เลย
function seedDefaultsIfMissing() {
  let data = loadDataSync();
  const isFirstRun = !data;
  if (!data) data = { customers: [], catalog: [], categories: [], bills: [], settings: {} };
  data.customers = data.customers || [];
  data.catalog = data.catalog || [];
  data.categories = data.categories || [];
  data.bills = data.bills || [];

  let changed = isFirstRun;
  const existingCatIds = new Set(data.categories.map((c) => c.id));
  DEFAULT_CATEGORIES.forEach((c) => {
    if (!existingCatIds.has(c.id)) {
      data.categories.push(c);
      changed = true;
    }
  });
  const existingItemIds = new Set(data.catalog.map((i) => i.id));
  DEFAULT_CATALOG.forEach((i) => {
    if (!existingItemIds.has(i.id)) {
      data.catalog.push(i);
      changed = true;
    }
  });

  // เติมรหัสสินค้า (code) ให้รายการเดิมที่ยังไม่มี โดยเทียบกับรายการเริ่มต้นเท่านั้น
  // ไม่แตะฟิลด์อื่นของผู้ใช้ (ชื่อ ราคา หมวดหมู่ ฯลฯ)
  const defaultCodeById = new Map(DEFAULT_CATALOG.map((i) => [i.id, i.code]));
  data.catalog.forEach((item) => {
    if (!item.code && defaultCodeById.has(item.id)) {
      item.code = defaultCodeById.get(item.id);
      changed = true;
    }
  });

  if (changed) saveDataSync(data);
}

/* ============================= EXCEL SYNC ============================= */
const typeLabel = (t) =>
  t === "cart" ? "ซาเล้งประจำ" : t === "factory" ? "โรงงาน" : "รายย่อยทั่วไป";
function catLabel(id, categories) {
  const found = (categories || []).find((c) => c.id === id);
  if (found) return found.label;
  const fallback = DEFAULT_CATEGORIES.find((c) => c.id === id);
  return fallback ? fallback.label : id;
}
const fmtDT = (iso) => {
  const d = new Date(iso);
  return isNaN(d) ? "" : d.toLocaleString("th-TH", { dateStyle: "short", timeStyle: "short" });
};
const fmtDateOnly = (iso) => {
  const d = new Date(iso);
  return isNaN(d) ? "" : d.toLocaleDateString("th-TH", { day: "2-digit", month: "2-digit", year: "numeric" });
};
function billStatsNode(bill) {
  let weight = 0, amount = 0;
  (bill.items || []).forEach((it) => {
    if (it.mode === "weight") weight += it.net;
    amount += it.amount;
  });
  return { weight, amount };
}

async function regenerateExcel(data) {
  const storeName = (data.settings && data.settings.storeName) || "ระบบ POS ค้าของเก่า by.leolung";
  const categories = data.categories || DEFAULT_CATEGORIES;
  const wb = new ExcelJS.Workbook();
  wb.creator = storeName;
  wb.created = new Date();

  const wsCatalog = wb.addWorksheet("สินค้า");
  wsCatalog.columns = [
    { header: "รหัสสินค้า", key: "code", width: 12 },
    { header: "หมวดหมู่", key: "cat", width: 22 },
    { header: "ชื่อสินค้า", key: "name", width: 26 },
    { header: "หน่วย", key: "unit", width: 10 },
    { header: "ราคาประจำ", key: "pr", width: 12 },
    { header: "ราคาจร", key: "pc", width: 12 },
    { header: "ปุ่มลัด", key: "quick", width: 10 },
  ];
  (data.catalog || []).forEach((i) => {
    wsCatalog.addRow({
      code: i.code || "", cat: catLabel(i.category, categories), name: i.name, unit: i.unit,
      pr: i.priceRegular, pc: i.priceCasual, quick: i.quick ? "ใช่" : "",
    });
  });
  wsCatalog.getRow(1).font = { bold: true };

  const wsCustomers = wb.addWorksheet("สมาชิก");
  wsCustomers.columns = [
    { header: "ชื่อ", key: "name", width: 24 },
    { header: "เบอร์โทร", key: "phone", width: 16 },
    { header: "ประเภท", key: "type", width: 16 },
    { header: "ยืนยันตัวตน", key: "verified", width: 14 },
    { header: "วันที่สมัคร", key: "created", width: 18 },
  ];
  (data.customers || []).forEach((c) => {
    wsCustomers.addRow({
      name: c.name, phone: c.phone, type: typeLabel(c.type),
      verified: c.idVerified ? "ยืนยันแล้ว" : "", created: fmtDT(c.createdAt),
    });
  });
  wsCustomers.getRow(1).font = { bold: true };

  const wsBills = wb.addWorksheet("รายการขาย");
  wsBills.columns = [
    { header: "เลขที่บิล", key: "no", width: 10 },
    { header: "วันที่เวลา", key: "date", width: 18 },
    { header: "ลูกค้า", key: "customer", width: 20 },
    { header: "ประเภทลูกค้า", key: "ctype", width: 16 },
    { header: "สินค้า", key: "item", width: 26 },
    { header: "น้ำหนัก/จำนวน", key: "qty", width: 14 },
    { header: "หน่วย", key: "unit", width: 8 },
    { header: "ประเภทราคา", key: "pmode", width: 12 },
    { header: "ราคาต่อหน่วย", key: "price", width: 12 },
    { header: "จำนวนเงิน", key: "amount", width: 12 },
  ];
  (data.bills || []).forEach((b) => {
    (b.items || []).forEach((it) => {
      wsBills.addRow({
        no: b.no, date: fmtDT(b.date), customer: b.customerName,
        ctype: b.customerType ? typeLabel(b.customerType) : "ทั่วไป",
        item: it.name, qty: it.mode === "weight" ? Number(it.net.toFixed(2)) : it.qty, unit: it.unit,
        pmode: it.priceMode === "regular" ? "ประจำ" : "จร", price: it.price, amount: it.amount,
      });
    });
  });
  wsBills.getRow(1).font = { bold: true };

  // สรุปยอดรายเดือน — ภาพรวมทุกเดือนในไฟล์เดียว
  const wsMonthly = wb.addWorksheet("สรุปรายเดือน");
  wsMonthly.columns = [
    { header: "เดือน", key: "month", width: 14 },
    { header: "จำนวนบิล", key: "count", width: 12 },
    { header: "น้ำหนักรวม (กก.)", key: "weight", width: 16 },
    { header: "ยอดขายรวม (บาท)", key: "amount", width: 16 },
  ];
  const byMonth = new Map();
  (data.bills || []).forEach((b) => {
    const key = (b.date || "").slice(0, 7);
    if (!key) return;
    if (!byMonth.has(key)) byMonth.set(key, { count: 0, weight: 0, amount: 0 });
    const s = billStatsNode(b);
    const m = byMonth.get(key);
    m.count += 1; m.weight += s.weight; m.amount += s.amount;
  });
  Array.from(byMonth.keys()).sort().reverse().forEach((key) => {
    const m = byMonth.get(key);
    wsMonthly.addRow({ month: key, count: m.count, weight: Number(m.weight.toFixed(2)), amount: Number(m.amount.toFixed(2)) });
  });
  wsMonthly.getRow(1).font = { bold: true };

  // รายรับเสริม
  const wsIncome = wb.addWorksheet("รายรับเสริม");
  wsIncome.columns = [
    { header: "วันที่", key: "date", width: 14 },
    { header: "รายการ", key: "desc", width: 30 },
    { header: "จำนวนเงิน", key: "amount", width: 14 },
  ];
  (data.extraIncome || []).forEach((i) => {
    wsIncome.addRow({ date: i.date, desc: i.description, amount: i.amount });
  });
  wsIncome.getRow(1).font = { bold: true };

  // พนักงาน
  const wsEmployees = wb.addWorksheet("พนักงาน");
  wsEmployees.columns = [
    { header: "ชื่อ", key: "name", width: 22 },
    { header: "ตำแหน่ง", key: "position", width: 20 },
    { header: "เงินเดือน", key: "salary", width: 14 },
  ];
  (data.employees || []).forEach((e) => {
    wsEmployees.addRow({ name: e.name, position: e.position || "", salary: e.monthlySalary || 0 });
  });
  wsEmployees.getRow(1).font = { bold: true };

  // เบิกเงิน
  const wsAdvances = wb.addWorksheet("เบิกเงิน");
  wsAdvances.columns = [
    { header: "วันที่", key: "date", width: 14 },
    { header: "พนักงาน", key: "emp", width: 22 },
    { header: "จำนวนเงิน", key: "amount", width: 14 },
    { header: "หมายเหตุ", key: "note", width: 26 },
  ];
  (data.advances || []).forEach((a) => {
    wsAdvances.addRow({ date: a.date, emp: a.employeeName, amount: a.amount, note: a.note || "" });
  });
  wsAdvances.getRow(1).font = { bold: true };

  const tmp = excelFilePath() + ".tmp";
  await wb.xlsx.writeFile(tmp);
  fs.renameSync(tmp, excelFilePath());
}

// สร้างไฟล์รายงานแยกรายเดือน — 1 ไฟล์ต่อเดือน จัดกลุ่มตามวันที่พร้อมยอดรวมย่อยรายวันและยอดรวมท้ายเดือน
async function regenerateMonthlyReports(data) {
  fs.mkdirSync(monthlyDirPath(), { recursive: true });
  const bills = data.bills || [];
  const byMonth = new Map();
  bills.forEach((b) => {
    const key = (b.date || "").slice(0, 7);
    if (!key) return;
    if (!byMonth.has(key)) byMonth.set(key, []);
    byMonth.get(key).push(b);
  });

  for (const [monthKey, monthBills] of byMonth.entries()) {
    const wb = new ExcelJS.Workbook();
    wb.created = new Date();
    const ws = wb.addWorksheet(`ขาย ${monthKey}`);
    ws.columns = [
      { header: "วันที่ / เลขบิล", key: "c1", width: 24 },
      { header: "ลูกค้า", key: "c2", width: 20 },
      { header: "สินค้า", key: "c3", width: 26 },
      { header: "น้ำหนัก/จำนวน", key: "c4", width: 14 },
      { header: "หน่วย", key: "c5", width: 8 },
      { header: "ราคาต่อหน่วย", key: "c6", width: 12 },
      { header: "จำนวนเงิน", key: "c7", width: 14 },
    ];
    ws.getRow(1).font = { bold: true };

    const byDate = new Map();
    monthBills
      .slice()
      .sort((a, b) => a.date.localeCompare(b.date))
      .forEach((b) => {
        const dKey = (b.date || "").slice(0, 10);
        if (!byDate.has(dKey)) byDate.set(dKey, []);
        byDate.get(dKey).push(b);
      });

    let monthWeight = 0, monthAmount = 0, monthBillCount = 0;
    Array.from(byDate.keys()).sort().forEach((dKey) => {
      const dayBills = byDate.get(dKey);
      const dateHeaderRow = ws.addRow({ c1: `วันที่ ${fmtDateOnly(dayBills[0].date)}` });
      dateHeaderRow.font = { bold: true };
      dateHeaderRow.eachCell((cell) => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE7E2D6" } }; });

      let dayWeight = 0, dayAmount = 0;
      dayBills.forEach((b) => {
        const s = billStatsNode(b);
        dayWeight += s.weight; dayAmount += s.amount;
        monthBillCount += 1;
        (b.items || []).forEach((it) => {
          ws.addRow({
            c1: `#${b.no}`, c2: b.customerName, c3: it.name,
            c4: it.mode === "weight" ? Number(it.net.toFixed(2)) : it.qty, c5: it.unit,
            c6: it.price, c7: it.amount,
          });
        });
      });
      monthWeight += dayWeight; monthAmount += dayAmount;

      const subtotalRow = ws.addRow({
        c1: `รวมวันที่ ${fmtDateOnly(dayBills[0].date)}`,
        c4: Number(dayWeight.toFixed(2)), c7: Number(dayAmount.toFixed(2)),
      });
      subtotalRow.font = { bold: true, italic: true };
      ws.addRow({});
    });

    const totalRow = ws.addRow({
      c1: "ยอดรวมทั้งเดือน", c2: `${monthBillCount} บิล`,
      c4: Number(monthWeight.toFixed(2)), c7: Number(monthAmount.toFixed(2)),
    });
    totalRow.font = { bold: true, size: 12 };
    totalRow.eachCell((cell) => { cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCE9DD" } }; });

    const fileName = `ขาย-${monthKey}.xlsx`;
    const tmp = path.join(monthlyDirPath(), fileName + ".tmp");
    await wb.xlsx.writeFile(tmp);
    fs.renameSync(tmp, path.join(monthlyDirPath(), fileName));
  }
}

/* ============================= GOOGLE DRIVE BACKUP ============================= */
function loadDriveConfig() {
  try { return JSON.parse(fs.readFileSync(driveConfigPath(), "utf-8")); } catch (e) { return null; }
}
function saveDriveConfig(cfg) {
  fs.writeFileSync(driveConfigPath(), JSON.stringify(cfg), "utf-8");
}
function loadDriveToken() {
  try { return JSON.parse(fs.readFileSync(driveTokenPath(), "utf-8")); } catch (e) { return null; }
}
function saveDriveToken(token) {
  fs.writeFileSync(driveTokenPath(), JSON.stringify(token), "utf-8");
}
function clearDriveToken() {
  try { fs.unlinkSync(driveTokenPath()); } catch (e) {}
}

const DRIVE_SCOPES = [
  "https://www.googleapis.com/auth/drive.file",
  "https://www.googleapis.com/auth/userinfo.email",
];

async function connectDrive() {
  const cfg = loadDriveConfig();
  if (!cfg || !cfg.clientId || !cfg.clientSecret) {
    return { ok: false, error: "ยังไม่ได้ตั้งค่า Client ID / Client Secret" };
  }

  return new Promise((resolve) => {
    const server = http.createServer();
    server.listen(0, "127.0.0.1", async () => {
      const port = server.address().port;
      const redirectUri = `http://127.0.0.1:${port}/oauth2callback`;
      const client = new google.auth.OAuth2(cfg.clientId, cfg.clientSecret, redirectUri);
      const authUrl = client.generateAuthUrl({
        access_type: "offline",
        prompt: "consent",
        scope: DRIVE_SCOPES,
      });

      let settled = false;
      const timeout = setTimeout(() => {
        if (settled) return;
        settled = true;
        server.close();
        resolve({ ok: false, error: "หมดเวลารอการอนุญาตจาก Google (5 นาที)" });
      }, 5 * 60 * 1000);

      server.on("request", async (req, res) => {
        if (!req.url.startsWith("/oauth2callback")) { res.end(); return; }
        const url = new URL(req.url, redirectUri);
        const code = url.searchParams.get("code");
        const errParam = url.searchParams.get("error");
        res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
        if (errParam) {
          res.end("<html><body style='font-family:sans-serif;padding:40px'><h2>ยกเลิกการเชื่อมต่อ</h2><p>ปิดหน้าต่างนี้ได้เลย</p></body></html>");
        } else {
          res.end("<html><body style='font-family:sans-serif;padding:40px'><h2>เชื่อมต่อสำเร็จ</h2><p>กลับไปที่โปรแกรมได้เลย ปิดหน้าต่างนี้ได้</p></body></html>");
        }
        if (settled) return;
        settled = true;
        clearTimeout(timeout);
        server.close();

        if (errParam || !code) {
          resolve({ ok: false, error: "ยกเลิกการเชื่อมต่อ" });
          return;
        }
        try {
          const { tokens } = await client.getToken(code);
          client.setCredentials(tokens);
          const oauth2 = google.oauth2({ version: "v2", auth: client });
          const info = await oauth2.userinfo.get();
          const email = info.data.email || "";
          saveDriveToken({ ...tokens, email, redirectUri });
          resolve({ ok: true, email });
        } catch (e) {
          resolve({ ok: false, error: "แลกโทเคนไม่สำเร็จ: " + e.message });
        }
      });

      await shell.openExternal(authUrl);
    });
    server.on("error", (e) => resolve({ ok: false, error: e.message }));
  });
}

async function getDriveClient() {
  const token = loadDriveToken();
  if (!token) return null;
  const cfg = loadDriveConfig();
  const client = new google.auth.OAuth2(cfg?.clientId, cfg?.clientSecret, token.redirectUri);
  client.setCredentials(token);
  client.on("tokens", (t) => {
    const merged = { ...token, ...t };
    if (!t.refresh_token) merged.refresh_token = token.refresh_token;
    saveDriveToken(merged);
  });
  return client;
}

async function findOrCreateFolder(drive, name, parentId) {
  const parentClause = parentId ? ` and '${parentId}' in parents` : "";
  const q = `name='${name.replace(/'/g, "\\'")}' and mimeType='application/vnd.google-apps.folder' and trashed=false${parentClause}`;
  const list = await drive.files.list({ q, fields: "files(id,name)" });
  if (list.data.files && list.data.files.length > 0) return list.data.files[0].id;
  const folder = await drive.files.create({
    requestBody: { name, mimeType: "application/vnd.google-apps.folder", parents: parentId ? [parentId] : undefined },
    fields: "id",
  });
  return folder.data.id;
}

async function upsertFile(drive, folderId, name, filePath, mimeType) {
  const q = `name='${name.replace(/'/g, "\\'")}' and '${folderId}' in parents and trashed=false`;
  const list = await drive.files.list({ q, fields: "files(id,name)" });
  const media = { mimeType, body: fs.createReadStream(filePath) };
  if (list.data.files && list.data.files.length > 0) {
    await drive.files.update({ fileId: list.data.files[0].id, media });
  } else {
    await drive.files.create({ requestBody: { name, parents: [folderId] }, media, fields: "id" });
  }
}

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

async function backupNow() {
  const client = await getDriveClient();
  if (!client) return { ok: false, error: "ยังไม่ได้เชื่อมต่อ Google Drive" };
  try {
    const drive = google.drive({ version: "v3", auth: client });
    const folderId = await findOrCreateFolder(drive, "ระบบ POS ค้าของเก่า by.leolung - สำรองข้อมูล", null);
    if (fs.existsSync(excelFilePath())) {
      await upsertFile(drive, folderId, "scrap-pos-data.xlsx", excelFilePath(), XLSX_MIME);
    }
    if (fs.existsSync(dataFilePath())) {
      await upsertFile(drive, folderId, "scrap-pos-data.json", dataFilePath(), "application/json");
    }
    if (fs.existsSync(monthlyDirPath())) {
      const monthlyFolderId = await findOrCreateFolder(drive, "รายงานรายเดือน", folderId);
      const files = fs.readdirSync(monthlyDirPath()).filter((f) => f.endsWith(".xlsx"));
      for (const f of files) {
        await upsertFile(drive, monthlyFolderId, f, path.join(monthlyDirPath(), f), XLSX_MIME);
      }
    }
    if (fs.existsSync(receiptRootDirPath())) {
      const receiptFolderId = await findOrCreateFolder(drive, "ใบเสร็จ", folderId);
      const monthDirs = fs
        .readdirSync(receiptRootDirPath())
        .filter((f) => fs.statSync(path.join(receiptRootDirPath(), f)).isDirectory());
      for (const monthDir of monthDirs) {
        const monthFolderId = await findOrCreateFolder(drive, monthDir, receiptFolderId);
        const monthPath = path.join(receiptRootDirPath(), monthDir);
        const imgFiles = fs.readdirSync(monthPath).filter((f) => f.endsWith(".png"));
        for (const f of imgFiles) {
          await upsertFile(drive, monthFolderId, f, path.join(monthPath, f), "image/png");
        }
      }
    }
    lastBackupAt = new Date().toISOString();
    return { ok: true, at: lastBackupAt };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

let lastBackupAt = null;
let backupDebounceTimer = null;
function scheduleDriveBackup() {
  const token = loadDriveToken();
  if (!token) return;
  clearTimeout(backupDebounceTimer);
  backupDebounceTimer = setTimeout(() => {
    backupNow().catch(() => {});
  }, 30 * 1000);
}
setInterval(() => {
  if (loadDriveToken()) backupNow().catch(() => {});
}, 15 * 60 * 1000);

/* ============================= LAN SYNC SERVER ============================= */
let mainWindow;
let wss;
const netInfo = { port: PORT, ip: null, url: null };

function stateMessage() {
  const data = loadDataSync() || {};
  return JSON.stringify({
    type: "state",
    customers: data.customers || [],
    catalog: data.catalog || [],
    categories: data.categories || [],
    bills: data.bills || [],
    extraIncome: data.extraIncome || [],
    employees: data.employees || [],
    advances: data.advances || [],
    settings: data.settings || null,
  });
}

function broadcastState() {
  if (!wss) return;
  const payload = stateMessage();
  wss.clients.forEach((client) => {
    if (client.readyState === 1) client.send(payload);
  });
}

// ส่งเหตุการณ์จากเครื่องอ่านบัตรประชาชนไปให้ทุกอุปกรณ์ที่เชื่อมต่ออยู่ (รวมถึงระบบอ่านบัตร
// แยกต่างหาก) ผ่านช่องทางเดียวกับที่ซิงก์ข้อมูลร้าน — กันไม่ให้อุปกรณ์อื่นต้องเปิดโปรแกรม
// อ่านบัตรของตัวเองซ้ำ (เคยเจอปัญหาเปิดพร้อมกัน 2 ตัวแล้วแย่งกันคุยกับเครื่องอ่านจนอ่านไม่ออก)
function broadcastCardEvent(msg) {
  if (!wss) return;
  const payload = JSON.stringify({ type: "cardEvent", event: msg });
  wss.clients.forEach((client) => {
    if (client.readyState === 1) client.send(payload);
  });
}

function applyMutation(msg) {
  const data = loadDataSync() || {};
  data.customers = data.customers || [];
  data.catalog = data.catalog || [];
  data.categories = data.categories || [];
  data.bills = data.bills || [];
  data.extraIncome = data.extraIncome || [];
  data.employees = data.employees || [];
  data.advances = data.advances || [];

  if (msg.type === "addCustomer" && msg.customer && msg.customer.id) {
    const idx = data.customers.findIndex((c) => c.id === msg.customer.id);
    if (idx >= 0) data.customers[idx] = msg.customer;
    else data.customers.unshift(msg.customer);
  } else if (msg.type === "deleteCustomer" && msg.id) {
    data.customers = data.customers.filter((c) => c.id !== msg.id);
  } else if (msg.type === "updateCatalogItem" && msg.item && msg.item.id) {
    const idx = data.catalog.findIndex((i) => i.id === msg.item.id);
    if (idx >= 0) data.catalog[idx] = msg.item;
    else data.catalog.push(msg.item);
  } else if (msg.type === "deleteCatalogItem" && msg.id) {
    data.catalog = data.catalog.filter((i) => i.id !== msg.id);
  } else if (msg.type === "updateCategory" && msg.category && msg.category.id) {
    const idx = data.categories.findIndex((c) => c.id === msg.category.id);
    if (idx >= 0) data.categories[idx] = msg.category;
    else data.categories.push(msg.category);
  } else if (msg.type === "deleteCategory" && msg.id) {
    data.categories = data.categories.filter((c) => c.id !== msg.id);
  } else if (msg.type === "addBill" && msg.bill && msg.bill.id) {
    const idx = data.bills.findIndex((b) => b.id === msg.bill.id);
    if (idx >= 0) data.bills[idx] = msg.bill;
    else data.bills.push(msg.bill);
  } else if (msg.type === "deleteBill" && msg.id) {
    data.bills = data.bills.filter((b) => b.id !== msg.id);
  } else if (msg.type === "updateSettings" && msg.settings) {
    data.settings = { ...(data.settings || {}), ...msg.settings };
  } else if (msg.type === "addIncome" && msg.item && msg.item.id) {
    const idx = data.extraIncome.findIndex((i) => i.id === msg.item.id);
    if (idx >= 0) data.extraIncome[idx] = msg.item;
    else data.extraIncome.unshift(msg.item);
  } else if (msg.type === "deleteIncome" && msg.id) {
    data.extraIncome = data.extraIncome.filter((i) => i.id !== msg.id);
  } else if (msg.type === "updateEmployee" && msg.employee && msg.employee.id) {
    const idx = data.employees.findIndex((e) => e.id === msg.employee.id);
    if (idx >= 0) data.employees[idx] = msg.employee;
    else data.employees.push(msg.employee);
  } else if (msg.type === "deleteEmployee" && msg.id) {
    data.employees = data.employees.filter((e) => e.id !== msg.id);
  } else if (msg.type === "addAdvance" && msg.item && msg.item.id) {
    const idx = data.advances.findIndex((a) => a.id === msg.item.id);
    if (idx >= 0) data.advances[idx] = msg.item;
    else data.advances.unshift(msg.item);
  } else if (msg.type === "deleteAdvance" && msg.id) {
    data.advances = data.advances.filter((a) => a.id !== msg.id);
  } else {
    return false;
  }
  saveDataSync(data);
  return true;
}

function startServer(onReady) {
  const expressApp = express();
  expressApp.use(express.static(path.join(__dirname, "renderer")));

  const server = http.createServer(expressApp);
  wss = new WebSocketServer({ server, path: "/ws" });

  wss.on("connection", (ws) => {
    ws.send(stateMessage());
    ws.on("message", (raw) => {
      let msg;
      try {
        msg = JSON.parse(raw);
      } catch (e) {
        return;
      }
      if (applyMutation(msg)) broadcastState();
    });
  });

  server.on("error", (err) => {
    console.error("LAN server error:", err);
  });

  server.listen(PORT, "0.0.0.0", () => {
    netInfo.ip = getLanIp();
    netInfo.url = netInfo.ip ? `http://${netInfo.ip}:${PORT}` : null;
    if (onReady) onReady();
  });
}

/* ============================= เครื่องอ่านบัตรประชาชน (Smart Card) =============================
 * ใช้โปรแกรมช่วย ThaiIdReader.exe (เขียนด้วย C#/.NET เรียก winscard.dll โดยตรง) เพราะไลบรารี
 * มาตรฐานฝั่ง Node.js (nfc-pcsc) ต้องคอมไพล์ native module ผ่าน Visual Studio C++ Build Tools
 * ซึ่งเครื่องพัฒนา/เครื่องร้านอาจไม่มีติดตั้งไว้ ตัว .exe เป็นแบบ self-contained ไม่ต้องลง .NET เพิ่ม
 * วนเช็คบัตรตลอดเวลา ส่งผลลัพธ์เป็น JSON ทีละบรรทัด (NDJSON) ทาง stdout
 *
 * สำคัญ: ให้โปรแกรม POS หลักเป็นตัวเดียวที่รันไฟล์นี้เท่านั้น (ห้ามมีโปรแกรมอื่นเปิดเครื่องอ่าน
 * บัตรของตัวเองพร้อมกัน เคยเจอปัญหาสองโปรเซสแย่งกันคุยกับเครื่องอ่านจนอ่านบัตรไม่ออกเลยทั้งคู่
 * มาแล้วจริง) โปรแกรมอื่นที่ต้องการข้อมูลบัตร ให้รับผ่าน WebSocket ของโปรแกรม POS หลักแทน
 * (ดู broadcastCardEvent ด้านล่าง — ส่ง {type:"cardEvent", event:...} ให้ทุกอุปกรณ์ที่เชื่อมอยู่)
 */
function cardReaderExePath() {
  const base = app.isPackaged ? process.resourcesPath : __dirname;
  return path.join(base, "card-reader-bin", "ThaiIdReader.exe");
}

let cardReaderChild = null;
let cardReaderQuitting = false;

// ปิดโปรแกรมอ่านบัตรลูกให้เรียบร้อยก่อนออกจากโปรแกรมหลักเสมอ กันไม่ให้ ThaiIdReader.exe
// ค้างอยู่เบื้องหลังหลังปิดโปรแกรม (เคยเจอปัญหาเปิดค้างไว้หลายตัวพร้อมกันแล้วแย่งกันคุยกับ
// เครื่องอ่านบัตรจนอ่านไม่ออกเลยสักตัว)
function stopCardReaderWatcher() {
  cardReaderQuitting = true;
  if (cardReaderChild && !cardReaderChild.killed) {
    try {
      cardReaderChild.kill();
    } catch (e) {}
  }
  cardReaderChild = null;
}

function startCardReaderWatcher() {
  if (process.platform !== "win32") return;
  const exePath = cardReaderExePath();
  if (!fs.existsSync(exePath)) {
    console.log("ไม่พบโปรแกรมอ่านบัตรประชาชน (card-reader-bin/ThaiIdReader.exe) — ข้ามการเชื่อมต่อเครื่องอ่านบัตร");
    return;
  }
  let child;
  try {
    child = spawn(exePath, [], { windowsHide: true });
  } catch (e) {
    console.error("เปิดโปรแกรมอ่านบัตรประชาชนไม่สำเร็จ:", e);
    return;
  }
  cardReaderChild = child;
  const rl = readline.createInterface({ input: child.stdout });
  rl.on("line", (line) => {
    let msg;
    try {
      msg = JSON.parse(line);
    } catch (e) {
      return;
    }
    if (
      msg.type === "card_data" ||
      msg.type === "card_removed" ||
      msg.type === "no_reader" ||
      msg.type === "read_error" ||
      msg.type === "fatal_error"
    ) {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send("cardreader:event", msg);
      }
      broadcastCardEvent(msg);
      if (msg.type === "fatal_error" || msg.type === "read_error") {
        console.error("card reader:", msg.message);
      }
    }
  });
  child.stderr.on("data", (d) => console.error("card reader stderr:", d.toString()));
  child.on("exit", (code) => {
    if (cardReaderQuitting) return;
    console.log("โปรแกรมอ่านบัตรประชาชนหยุดทำงาน (code " + code + ") กำลังลองเปิดใหม่...");
    setTimeout(startCardReaderWatcher, 5000);
  });
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1440,
    height: 920,
    minWidth: 1000,
    minHeight: 700,
    backgroundColor: "#141714",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  Menu.setApplicationMenu(null);
  mainWindow.loadURL(`http://localhost:${PORT}`);
}

/* ============================= อัปเดตอัตโนมัติ (ผ่าน GitHub Releases) =============================
 * ตรวจสอบเวอร์ชันใหม่ตอนเปิดโปรแกรม เฉพาะตอนเป็นเวอร์ชันที่ติดตั้งจริงเท่านั้น (ข้ามตอนรันแบบ dev
 * ด้วย npm start) ถ้ามีอัปเดตจะดาวน์โหลดเงียบๆ อยู่เบื้องหลัง แล้วแจ้งเตือนในโปรแกรมให้กดรีสตาร์ท
 * เพื่อติดตั้งได้ทันที — ถ้าไม่กด โปรแกรมจะติดตั้งให้อัตโนมัติตอนปิดโปรแกรมครั้งถัดไปอยู่ดี
 */
function autoUpdaterCheck() {
  return autoUpdater.checkForUpdates().catch((e) => console.error("checkForUpdates failed:", e));
}

function setupAutoUpdater() {
  if (!app.isPackaged) return;

  autoUpdater.autoDownload = true;
  autoUpdater.autoInstallOnAppQuit = true;

  const notify = (type, extra) => {
    if (mainWindow) mainWindow.webContents.send("update:status", { type, ...(extra || {}) });
  };

  autoUpdater.on("checking-for-update", () => notify("checking"));
  autoUpdater.on("update-available", (info) => notify("available", { version: info.version }));
  autoUpdater.on("update-not-available", () => notify("not-available"));
  autoUpdater.on("update-downloaded", (info) => notify("downloaded", { version: info.version }));
  autoUpdater.on("error", (err) => {
    const message = err && (err.stack || err.message) ? (err.stack || err.message) : String(err);
    console.error("Auto-update error:", message);
    notify("error", { message: (err && err.message) || String(err) });
  });

  setTimeout(autoUpdaterCheck, 5000);
  setInterval(autoUpdaterCheck, 4 * 60 * 60 * 1000); // เช็คซ้ำทุก 4 ชั่วโมง เผื่อเปิดโปรแกรมค้างไว้นาน
}

ipcMain.handle("app:installUpdate", () => {
  autoUpdater.quitAndInstall();
  return true;
});
ipcMain.handle("app:version", () => app.getVersion());
// ปุ่ม "เช็คอัปเดตตอนนี้" ในหน้าตั้งค่า — เรียกเช็คทันทีแทนที่จะรอรอบอัตโนมัติ (ทุก 4 ชม.)
ipcMain.handle("app:checkForUpdates", () => {
  if (!app.isPackaged) return { ok: false, reason: "not-packaged" };
  autoUpdaterCheck();
  return { ok: true };
});

app.whenReady().then(() => {
  seedDefaultsIfMissing();
  startServer(() => createWindow());
  startCardReaderWatcher();
  setupAutoUpdater();
});

app.on("window-all-closed", () => {
  stopCardReaderWatcher();
  if (process.platform !== "darwin") app.quit();
});
app.on("before-quit", () => {
  stopCardReaderWatcher();
});
app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

ipcMain.on("data:loadSync", (event) => {
  event.returnValue = loadDataSync();
});
ipcMain.handle("data:load", () => loadDataSync());
ipcMain.handle("data:save", (_e, data) => {
  saveDataSync(data);
  return true;
});
ipcMain.handle("data:saveLocal", (_e, localData) => {
  // บันทึกเฉพาะ draft/heldBills (ข้อมูลที่ยังไม่ผ่าน WebSocket) โดยอ่านไฟล์ล่าสุดจากดิสก์ก่อนแล้วค่อยแก้เฉพาะสองฟิลด์นี้
  // เพื่อไม่ให้ไปทับข้อมูล customers/catalog/bills ฯลฯ ที่อาจถูกอัปเดตผ่าน WebSocket ไปแล้วแต่ state ฝั่ง renderer ยังไม่ทันอัปเดตตาม
  const data = loadDataSync() || {};
  data.draft = localData.draft;
  data.heldBills = localData.heldBills;
  saveDataSync(data);
  return true;
});
ipcMain.handle("csv:export", async (_e, { filename, content }) => {
  const { canceled, filePath } = await dialog.showSaveDialog(mainWindow, {
    title: "บันทึกไฟล์ CSV",
    defaultPath: filename,
    filters: [{ name: "CSV", extensions: ["csv"] }],
  });
  if (canceled || !filePath) return { ok: false, reason: "cancelled" };
  fs.writeFileSync(filePath, "﻿" + content, "utf-8");
  return { ok: true, path: filePath };
});
ipcMain.handle("app:dataDir", () => app.getPath("userData"));
ipcMain.handle("print:receipt", (_e, mmWidth) => {
  if (!mainWindow) return { ok: false };
  const widthMicrons = (Number(mmWidth) || 80) * 1000;
  mainWindow.webContents.print(
    {
      // silent:true — ตอน silent:false เดิม เครื่องพิมพ์จะเด้งกล่องโต้ตอบของ Windows ขึ้นมาทุกครั้ง
      // ซึ่งกล่องนั้นใช้ "จำนวนชุด" ของตัวเองที่เคยตั้งไว้ก่อนหน้า (อาจยังเป็น 2 ค้างอยู่) แทนค่า
      // copies ที่โปรแกรมส่งไป ทำให้กดพิมพ์ 1 ใบแต่ได้ 2 ใบจริง — เปลี่ยนเป็นพิมพ์ตรงไม่ผ่านกล่อง
      // โต้ตอบ รับประกันว่าจำนวนชุดตรงกับที่ตั้งไว้ในโค้ดเสมอ (ใช้เครื่องพิมพ์ที่ตั้งเป็นค่าเริ่มต้น
      // ของ Windows) และยังเร็วขึ้นด้วยเพราะไม่ต้องกดยืนยันกล่องโต้ตอบทุกครั้ง
      silent: true,
      printBackground: true,
      copies: 1,
      collate: true,
      margins: { marginType: "none" },
      pageSize: { width: widthMicrons, height: 297000 }, // ความกว้างตามจุดที่พิมพ์ (จบการขาย/ประวัติ อาจตั้งค่าไม่เท่ากัน), สูง 297mm ไว้รองรับบิลยาว
      scaleFactor: 100,
    },
    (success, reason) => {
      if (!success && reason !== "cancelled") console.error("Print failed:", reason);
    }
  );
  return { ok: true };
});
ipcMain.handle("app:openMonthlyFolder", () => {
  fs.mkdirSync(monthlyDirPath(), { recursive: true });
  shell.openPath(monthlyDirPath());
  return true;
});
ipcMain.handle("receipt:saveImage", async (event, { rect, billNo, dateISO }) => {
  if (!mainWindow) return { ok: false, error: "no window" };
  try {
    await mainWindow.webContents.executeJavaScript(
      "new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)))"
    );
    const image = await mainWindow.webContents.capturePage(rect);
    const dir = receiptImageDirPath(dateISO);
    fs.mkdirSync(dir, { recursive: true });
    const safeTime = (dateISO || new Date().toISOString()).replace(/[:.]/g, "-");
    const filePath = path.join(dir, `บิล-${billNo}-${safeTime}.png`);
    fs.writeFileSync(filePath, image.toPNG());
    return { ok: true, path: filePath };
  } catch (e) {
    console.error("Save receipt image failed:", e);
    return { ok: false, error: String(e) };
  }
});
ipcMain.handle("net:info", () => netInfo);

ipcMain.handle("drive:status", () => {
  const cfg = loadDriveConfig();
  const token = loadDriveToken();
  return {
    hasCredentials: !!(cfg && cfg.clientId && cfg.clientSecret),
    connected: !!token,
    email: token ? token.email : null,
    lastBackupAt,
  };
});
ipcMain.handle("drive:setCredentials", (_e, { clientId, clientSecret }) => {
  saveDriveConfig({ clientId: (clientId || "").trim(), clientSecret: (clientSecret || "").trim() });
  clearDriveToken();
  return true;
});
ipcMain.handle("drive:connect", () => connectDrive());
ipcMain.handle("drive:disconnect", () => {
  clearDriveToken();
  return true;
});
ipcMain.handle("drive:backupNow", () => backupNow());
